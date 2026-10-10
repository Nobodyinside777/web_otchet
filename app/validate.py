"""Проверка формата листа и тексты сообщений для пользователя.

В сообщениях указываем только буквы колонок, номера строк и типы значений,
но не содержимое ячеек: в данных коммерческая тайна. Исключение — заголовки,
они не секретные и без них непонятно, что не так.
"""

import re
from dataclasses import dataclass, field

from app.schema import HEADER_ROW, SHEET_NAME, Column

# Если нераспознанных значений больше этой доли, колонка считается испорченной.
MAX_BAD_SHARE = 0.5
# ...но только если их хотя бы столько (чтобы пара опечаток в маленьком файле не была ошибкой).
MIN_BAD_FOR_ERROR = 5

TYPE_NAMES = {"int": "целые числа", "number": "числа", "text": "текст", "date": "даты"}


class FormatError(Exception):
    """Файл не подходит. problems — список понятных сообщений для пользователя."""

    def __init__(self, problems: list[str]):
        super().__init__("\n".join(problems))
        self.problems = problems


@dataclass
class ColumnStats:
    """Что встретилось в колонке при чтении."""

    filled: int = 0  # непустых ячеек
    empty_rows: list[int] = field(default_factory=list)  # пустые или из одних пробелов
    bad_rows: list[int] = field(default_factory=list)  # значение не того типа
    bad_types: set[str] = field(default_factory=set)  # какие типы встретились вместо ожидаемого
    error_rows: list[int] = field(default_factory=list)  # ошибки Excel (#VALUE!, #N/A ...)
    out_of_range_rows: list[int] = field(default_factory=list)  # например, месяц не 1–12


def normalize_header(value) -> str:
    """Убираем пробелы по краям и схлопываем повторные пробелы и переносы строк."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def check_headers(columns: list[Column], headers: list) -> list[str]:
    """Сравнивает заголовки с ожидаемыми. Возвращает список проблем.

    columns — проверяемые колонки (видимые), headers — значения их ячеек в строке заголовков.
    """
    problems = []
    for col, value in zip(columns, headers, strict=True):
        found = normalize_header(value)
        expected = normalize_header(col.header)
        if found != expected:
            shown = f"«{found}»" if found else "пустая ячейка"
            problems.append(
                f"Колонка {col.letter}: ожидался заголовок «{expected}», найдено {shown}."
            )
    if problems:
        problems.insert(0, f"Заголовки в строке {HEADER_ROW} не совпадают с ожидаемыми:")
    return problems


def missing_sheet_message(sheet_names: list[str]) -> str:
    return (
        f"В книге нет листа «{SHEET_NAME}». "
        f"Найдены листы: {', '.join(f'«{n}»' for n in sheet_names) or 'нет'}."
    )


NO_FORMULA_VALUES = (
    "В файле есть формулы без сохранённых значений (ячеек: {count}). "
    "Так бывает, если файл сохранён не из Excel. "
    "Откройте файл в Excel, сохраните и загрузите снова."
)

def no_data_rows_message(anchors: list[Column]) -> str:
    return (
        f"На листе «{SHEET_NAME}» нет строк с данными: ни в одной строке, начиная с 3-й, "
        f"не заполнены колонки {describe_columns(anchors)}."
    )


def no_anchor_columns_message(anchors: list[Column]) -> str:
    return (
        f"Все колонки, по которым определяются строки с данными, скрыты в Excel: "
        f"{describe_columns(anchors)}. Покажите хотя бы одну из них и сохраните файл."
    )


def hidden_columns_message(hidden: list[Column]) -> str:
    return f"Скрытые в Excel колонки не используются: {describe_columns(hidden)}."


def describe_columns(columns: list[Column]) -> str:
    """'B («Год»), C («Месяц выручки»)'."""
    return ", ".join(f"{c.letter} («{c.label}»)" for c in columns)


def check_types(
    columns: list[Column], stats: dict[str, ColumnStats]
) -> tuple[list[str], list[str], list[str]]:
    """По статистике колонок решает, что является ошибкой, а что предупреждением.

    Возвращает (ошибки, предупреждения, справка).
    """
    errors, warnings, info = [], [], []
    for col in columns:
        st = stats[col.code]
        bad = len(st.bad_rows)
        if bad:
            types = ", ".join(sorted(st.bad_types))
            text = (
                f"Колонка {col.letter} («{col.label}»): ожидались {TYPE_NAMES[col.dtype]}, "
                f"но в {bad} из {st.filled} заполненных ячеек другой тип: {types}. "
                f"Строки: {format_rows(st.bad_rows)}."
            )
            if col.strict and bad >= MIN_BAD_FOR_ERROR and bad > MAX_BAD_SHARE * st.filled:
                errors.append(text)
            else:
                warnings.append(text + " Эти ячейки считаются пустыми.")
        if st.error_rows:
            n = len(st.error_rows)
            warnings.append(
                f"Колонка {col.letter} («{col.label}»): {n} {plural(n, 'ячейка', 'ячейки', 'ячеек')} "
                f"с ошибкой Excel (например, #VALUE!), они считаются пустыми. "
                f"Строки: {format_rows(st.error_rows)}."
            )
        if st.out_of_range_rows:
            n = len(st.out_of_range_rows)
            warnings.append(
                f"Колонка {col.letter} («{col.label}»): {n} "
                f"{plural(n, 'значение', 'значения', 'значений')} вне диапазона 1–12. "
                f"Строки: {format_rows(st.out_of_range_rows)}."
            )
        if st.empty_rows and col.role in ("measure", "ratio"):
            info.append(
                f"Колонка {col.letter} («{col.label}»): пустых ячеек — {len(st.empty_rows)}. "
                f"Строки: {format_rows(st.empty_rows)}."
            )
    return errors, warnings, info


def plural(n: int, one: str, few: str, many: str) -> str:
    """Склонение по числу: 1 ячейка, 2 ячейки, 5 ячеек."""
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def skipped_rows_message(rows: list[int], anchors: list[Column]) -> str:
    return (
        f"Пропущено строк, где пусты колонки {describe_columns(anchors)}, "
        f"но есть другие данные: {len(rows)}. "
        f"Похоже на итоговые или черновые строки. Строки: {format_rows(rows)}."
    )


def format_rows(rows: list[int], limit: int = 15) -> str:
    """[3, 4, 5, 9, 10] -> '3–5, 9–10'. Длинные списки обрезаются."""
    if not rows:
        return ""
    ranges = []
    start = prev = rows[0]
    for r in rows[1:]:
        if r == prev + 1:
            prev = r
            continue
        ranges.append((start, prev))
        start = prev = r
    ranges.append((start, prev))
    parts = [str(a) if a == b else f"{a}–{b}" for a, b in ranges]
    if len(parts) > limit:
        rest = len(parts) - limit
        return ", ".join(parts[:limit]) + f" и ещё {rest} диапазонов"
    return ", ".join(parts)
