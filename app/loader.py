"""Чтение листа «Прогноз выручки» и приведение значений к нужным типам.

Значения читает python-calamine: он быстрый и берёт вычисленные значения формул,
которые Excel сохранил в файле. Один недостаток: ошибки Excel (#VALUE!) calamine
отдаёт как пустую строку, а формулы без сохранённых значений не отличает от пустых
ячеек. Поэтому XML листа дополнительно просматривается потоком (scan_sheet_xml),
чтобы найти такие ячейки. Целиком в память лист при этом не загружается.

Колонки, скрытые в Excel, не используются: их не проверяем и не берём в таблицу.
Сведения о скрытых колонках тоже берутся из XML листа (тег <cols>).
"""

import codecs
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import BinaryIO, Callable
from xml.etree import ElementTree

import polars as pl
from python_calamine import CalamineError, CalamineWorkbook

from app.schema import (
    COLUMNS,
    DATE,
    FIRST_DATA_ROW,
    HEADER_ROW,
    INT,
    LAST_LETTER,
    NUMBER,
    ROW_CODE,
    SHEET_NAME,
    TEXT,
    THRESHOLD_CELL,
    Column,
    col_index,
    col_letter,
)
from app import validate
from app.validate import ColumnStats, FormatError

Source = str | Path | BinaryIO

POLARS_TYPES = {INT: pl.Int64, NUMBER: pl.Float64, TEXT: pl.String, DATE: pl.Date}

# Строка считается строкой данных, если заполнена хотя бы одна из этих колонок.
# Так отсекаются итоговая строка внизу (там только суммы) и черновые строки
# с одним номером заказа, но не теряются сделки с незаполненным годом.
ANCHOR_CODES = ("year", "month", "revenue_date", "customer")
# Колонка Z заполнена формулой во всех строках листа, даже пустых: её не учитываем,
# когда решаем, пустая ли строка.
IGNORED_FOR_EMPTINESS = {"passes_threshold"}
# Колонки, где проверяем, что значение — номер месяца 1–12.
MONTH_CODES = {"month"}


@dataclass
class LoadResult:
    df: pl.DataFrame  # строки данных; первая колонка — номер строки в Excel
    threshold: float | None  # порог вероятности из Z1
    columns: list[Column]  # колонки, попавшие в таблицу (видимые в Excel)
    hidden: list[Column]  # колонки из схемы, скрытые в Excel и пропущенные
    warnings: list[str] = field(default_factory=list)  # стоит показать пользователю
    info: list[str] = field(default_factory=list)  # справочно, например пустые суммы


@dataclass
class SheetScan:
    """Что нашёл просмотр XML листа."""

    hidden_letters: set[str]  # буквы скрытых колонок
    error_rows: dict[str, set[int]]  # буква колонки -> строки с ошибкой Excel
    formulas_without_values: int


# ---------- приведение типов ----------
# Каждая функция возвращает (значение, состояние), где состояние:
# "ok", "empty" (пусто или одни пробелы) или название встреченного типа, если значение не подошло.

EMPTY = "empty"
OK = "ok"


def type_name(value) -> str:
    if isinstance(value, bool):
        return "логическое"
    if isinstance(value, (int, float)):
        return "число"
    if isinstance(value, (date, datetime)):
        return "дата"
    if isinstance(value, str):
        return "текст"
    return "другое"


def to_int(value):
    if isinstance(value, str):
        text = _squeeze(value)
        if not text:
            return None, EMPTY
        try:
            return int(text), OK
        except ValueError:
            return None, "текст"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            return int(value), OK
        return None, "дробное число"
    return None, type_name(value)


def to_number(value):
    if isinstance(value, str):
        text = _squeeze(value).replace(",", ".")
        if not text:
            return None, EMPTY
        scale = 1.0
        if text.endswith("%"):  # «10,0%» -> 0.1
            text, scale = text[:-1], 0.01
        try:
            return float(text) * scale, OK
        except ValueError:
            return None, "текст"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value), OK
    return None, type_name(value)


def to_text(value):
    if isinstance(value, str):
        text = value.strip()
        return (text, OK) if text else (None, EMPTY)
    if isinstance(value, float) and value.is_integer():
        return str(int(value)), OK
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y"), OK
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y"), OK
    return str(value), OK


DATE_FORMATS = ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d")
EXCEL_EPOCH = date(1899, 12, 30)


def to_date(value):
    if isinstance(value, datetime):
        return value.date(), OK
    if isinstance(value, date):
        return value, OK
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None, EMPTY
        for fmt in DATE_FORMATS:
            try:
                return datetime.strptime(text, fmt).date(), OK
            except ValueError:
                pass
        return None, "текст"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Дата, сохранённая как число (серийный номер Excel), в разумных пределах 1955–2099.
        if 20000 <= value <= 73050 and float(value).is_integer():
            return EXCEL_EPOCH + timedelta(days=int(value)), OK
        return None, "число"
    return None, type_name(value)


CONVERTERS = {INT: to_int, NUMBER: to_number, TEXT: to_text, DATE: to_date}


def _squeeze(text: str) -> str:
    """Убирает все пробелы, в том числе неразрывные: '1 234,5' -> '1234,5'."""
    return re.sub(r"\s", "", text)


# ---------- основная функция ----------


def load(source: Source, progress: Callable[[float], None] | None = None) -> LoadResult:
    """Читает лист, проверяет формат и возвращает таблицу.

    source — путь к файлу или открытый файл (например, загруженный в Streamlit).
    progress — необязательная функция, которой передаётся доля прочитанного (0..1).
    При несоответствии формату выбрасывает FormatError с понятными сообщениями.
    """
    workbook = _open_workbook(source)
    try:
        if SHEET_NAME not in workbook.sheet_names:
            raise FormatError([validate.missing_sheet_message(workbook.sheet_names)])

        try:
            scan = scan_sheet_xml(source, SHEET_NAME)
        except (zipfile.BadZipFile, KeyError, ElementTree.ParseError) as exc:
            raise FormatError(
                ["Поддерживаются только файлы .xlsx (формат Excel 2007 и новее). "
                 "Откройте файл в Excel и сохраните как «Книга Excel (*.xlsx)»."]
            ) from exc
        if scan.formulas_without_values:
            raise FormatError(
                [validate.NO_FORMULA_VALUES.format(count=scan.formulas_without_values)]
            )

        sheet = workbook.get_sheet_by_name(SHEET_NAME)
        return _read_sheet(sheet, scan, progress)
    finally:
        workbook.close()


def _read_sheet(sheet, scan: SheetScan, progress) -> LoadResult:
    start = sheet.start
    if start is None:
        raise FormatError([f"Лист «{SHEET_NAME}» пустой."])
    first_row, first_col = start[0] + 1, start[1]  # строка в нумерации Excel, сдвиг колонок
    total_rows = max(sheet.height, 1)

    columns = [c for c in COLUMNS if c.letter not in scan.hidden_letters]
    hidden = [c for c in COLUMNS if c.letter in scan.hidden_letters]
    anchor_columns = [c for c in COLUMNS if c.code in ANCHOR_CODES]
    anchors = [k for k, c in enumerate(columns) if c.code in ANCHOR_CODES]
    if not anchors:
        raise FormatError([validate.no_anchor_columns_message(anchor_columns)])
    visible_anchors = [columns[k] for k in anchors]

    indexes = [col_index(c.letter) - first_col for c in columns]
    converters = [CONVERTERS[c.dtype] for c in columns]
    errors_by_col = [scan.error_rows.get(c.letter, set()) for c in columns]
    check_month = [c.code in MONTH_CODES for c in columns]
    skip_for_empty = [c.code in IGNORED_FOR_EMPTINESS for c in columns]
    threshold_index = col_index(THRESHOLD_CELL.rstrip("0123456789")) - first_col

    values: list[list] = [[] for _ in columns]
    row_numbers: list[int] = []
    stats = {c.code: ColumnStats() for c in columns}
    stats_list = [stats[c.code] for c in columns]
    header_row: list = []
    threshold = None
    skipped: list[int] = []

    for i, raw in enumerate(sheet.iter_rows()):
        excel_row = first_row + i
        if progress and i % 5000 == 0:
            progress(min(i / total_rows, 1.0))

        if excel_row == HEADER_ROW - 1:
            # Порог из Z1 — это настройка листа, читаем его, даже если колонка Z скрыта.
            threshold = _threshold(raw, threshold_index)
            continue

        cells = [raw[j] if 0 <= j < len(raw) else "" for j in indexes]

        if excel_row == HEADER_ROW:
            header_row = cells
            problems = validate.check_headers(columns, cells)
            if problems:
                raise FormatError(problems)
            continue
        if excel_row < FIRST_DATA_ROW:
            continue

        if all(_is_blank(cells[k]) for k in anchors):
            if any(not _is_blank(v) for v, skip in zip(cells, skip_for_empty) if not skip):
                skipped.append(excel_row)
            continue

        row_numbers.append(excel_row)
        for k, cell in enumerate(cells):
            value, state = converters[k](cell)
            st = stats_list[k]
            if state == OK:
                st.filled += 1
                if check_month[k] and not 1 <= value <= 12:
                    st.out_of_range_rows.append(excel_row)
            elif state == EMPTY:
                if excel_row in errors_by_col[k]:
                    st.error_rows.append(excel_row)
                else:
                    st.empty_rows.append(excel_row)
            else:
                st.filled += 1
                st.bad_rows.append(excel_row)
                st.bad_types.add(state)
            values[k].append(value)

    if progress:
        progress(1.0)

    if not header_row:
        raise FormatError([f"На листе «{SHEET_NAME}» нет строки заголовков (строка {HEADER_ROW})."])
    if not row_numbers:
        raise FormatError([validate.no_data_rows_message(visible_anchors)])

    errors, warnings, info = validate.check_types(columns, stats)
    if errors:
        raise FormatError(["Типы данных в колонках не соответствуют ожидаемым:"] + errors)
    if skipped:
        warnings.insert(0, validate.skipped_rows_message(skipped, visible_anchors))

    df = pl.DataFrame(
        [pl.Series(ROW_CODE, row_numbers, dtype=pl.Int32)]
        + [
            pl.Series(c.code, vals, dtype=POLARS_TYPES[c.dtype])
            for c, vals in zip(columns, values)
        ]
    )
    return LoadResult(
        df=df,
        threshold=threshold,
        columns=columns,
        hidden=hidden,
        warnings=warnings,
        info=info,
    )


def _threshold(row1: list, index: int) -> float | None:
    if not 0 <= index < len(row1):
        return None
    value, state = to_number(row1[index])
    return value if state == OK else None


def _is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _open_workbook(source: Source) -> CalamineWorkbook:
    try:
        if isinstance(source, (str, Path)):
            return CalamineWorkbook.from_path(str(source))
        source.seek(0)
        return CalamineWorkbook.from_filelike(source)
    except (CalamineError, OSError, ValueError) as exc:
        raise FormatError(
            [f"Не удалось открыть файл как книгу Excel (.xlsx). Тип ошибки: {type(exc).__name__}."]
        ) from exc


# ---------- просмотр XML листа ----------

_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
# Ячейка с ошибкой: <c r="Q5" s="3" t="e"><v>#VALUE!</v></c>
_ERROR_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?\bt="e"')
# Формула без сохранённого значения: <c r="B5"><f>...</f></c>, тега <v> нет или он пустой
_FORMULA_NO_VALUE = re.compile(
    r'<c r="([A-Z]+)\d+"[^>]*>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>)(?:<v\s*/>|<v></v>)?</c>'
)
_CHUNK = 8 * 1024 * 1024


def scan_sheet_xml(source: Source, sheet_name: str) -> SheetScan:
    """Потоком просматривает XML листа: скрытые колонки, ошибки Excel, формулы без значений.

    Учитываются только колонки из схемы (B–AA), служебный текст правее игнорируется.
    """
    letters = {c.letter for c in COLUMNS}
    hidden: set[str] | None = None  # станет известно, когда дойдём до начала данных
    error_rows: dict[str, set[int]] = defaultdict(set)
    no_value = 0

    if not isinstance(source, (str, Path)):
        source.seek(0)
    with zipfile.ZipFile(source) as zf:
        path = _sheet_xml_path(zf, sheet_name)
        decoder = codecs.getincrementaldecoder("utf-8")(errors="replace")
        tail = ""
        with zf.open(path) as f:
            while True:
                chunk = f.read(_CHUNK)
                text = tail + decoder.decode(chunk, final=not chunk)
                if hidden is None:
                    # Описание колонок (<cols>) стоит в XML перед данными (<sheetData>).
                    start = text.find("<sheetData")
                    if start >= 0 or not chunk:
                        hidden = _hidden_letters(text[:start] if start >= 0 else text)
                        letters -= hidden
                    elif len(text) < 2 * _CHUNK:
                        tail = text
                        continue
                # Обрабатываем только целые строки листа, хвост переносим в следующий кусок.
                cut = len(text) if not chunk else text.rfind("</row>") + len("</row>")
                if cut < len("</row>"):
                    tail = text
                    continue
                body, tail = text[:cut], text[cut:]
                for m in _ERROR_CELL.finditer(body):
                    if m.group(1) in letters:
                        error_rows[m.group(1)].add(int(m.group(2)))
                for m in _FORMULA_NO_VALUE.finditer(body):
                    if m.group(1) in letters:
                        no_value += 1
                if not chunk:
                    break
    return SheetScan(
        hidden_letters=hidden or set(),
        error_rows=dict(error_rows),
        formulas_without_values=no_value,
    )


# <col min="22" max="24" width="0" hidden="1" customWidth="1"/>
_COL_TAG = re.compile(r"<col\b([^>]*)>")
_ATTR = re.compile(r'(\w+)="([^"]*)"')


def _hidden_letters(head: str) -> set[str]:
    """Буквы скрытых колонок по описанию <cols> из начала XML листа."""
    hidden = set()
    for tag in _COL_TAG.finditer(head):
        attrs = dict(_ATTR.findall(tag.group(1)))
        if attrs.get("hidden") in ("1", "true"):
            first, last = int(attrs["min"]), int(attrs["max"])
            # Excel нумерует колонки с 1, а col_letter — с 0. Правее AA нам ничего не нужно.
            last = min(last, col_index(LAST_LETTER) + 1)
            hidden.update(col_letter(n - 1) for n in range(first, last + 1))
    return hidden


def _sheet_xml_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """По имени листа находит его XML-файл внутри xlsx."""
    workbook = ElementTree.fromstring(zf.read("xl/workbook.xml"))
    rel_id = None
    for sheet in workbook.iterfind("m:sheets/m:sheet", _NS):
        if sheet.get("name") == sheet_name:
            rel_id = sheet.get(f"{{{_NS['r']}}}id")
    rels = ElementTree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.iterfind("rel:Relationship", _NS):
        if rel.get("Id") == rel_id:
            target = rel.get("Target")
            return target.lstrip("/") if target.startswith("/") else "xl/" + target
    raise FormatError([f"Не удалось найти данные листа «{sheet_name}» внутри файла."])
