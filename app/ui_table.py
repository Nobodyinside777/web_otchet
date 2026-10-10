"""Таблица с фильтрами в заголовках, как в Excel.

Сама таблица нарисована на JavaScript (app/table_component/table.js) и встроена в
страницу как компонент Streamlit. Этот модуль:
- готовит для неё данные: колонки, строки, списки значений для окон фильтров;
- превращает то, что выбрал пользователь, в фильтры ops.Filter и сортировку.
Фильтрует данные ops, поэтому итоги и выгрузка совпадают с тем, что видно в таблице.
"""

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path

import polars as pl
import streamlit as st

from app import ops
from app.schema import BY_CODE, DATE, INT, MEASURE, NUMBER, RATIO, ROW_CODE, ROW_LABEL, TEXT, Column

ASSETS = Path(__file__).parent / "table_component"
COMPONENT = st.components.v2.component(
    "prognoz_table",
    css=(ASSETS / "table.css").read_text(encoding="utf-8"),
    js=(ASSETS / "table.js").read_text(encoding="utf-8"),
)

STATE_KEY = "table_view"  # здесь лежит последнее состояние таблицы: фильтры и сортировка
MAX_OPTIONS = 3000  # больше значений в окне фильтра не показываем, вместо списка — поиск по тексту
MAX_ROWS_ON_SCREEN = 100_000  # больше строк браузер показывает медленно; итоги считаются по всем


@dataclass
class TableView:
    filters: list[ops.Filter]
    sort_by: str | None
    descending: bool


def current_view() -> TableView:
    """Фильтры и сортировка, выбранные в таблице (до её отрисовки в этом запуске)."""
    state = st.session_state.get(STATE_KEY) or {}
    filters = [_to_filter(code, f) for code, f in (state.get("filters") or {}).items() if code in BY_CODE]
    sort = state.get("sort") or {}
    return TableView(
        filters=[f for f in filters if f is not None],
        sort_by=sort.get("code"),
        descending=bool(sort.get("desc")),
    )


def reset_view() -> None:
    st.session_state.pop(STATE_KEY, None)


def show_table(
    all_rows: pl.DataFrame,
    shown: pl.DataFrame,
    columns: list[Column],
    view: TableView,
    key: str,
    height: int = 600,
) -> None:
    """Рисует таблицу. all_rows — все строки файла, shown — после фильтров (без сортировки)."""
    state = st.session_state.get(STATE_KEY) or {}
    payload_columns = [_row_column()] + [
        _column_payload(all_rows, col, view.filters, i + 1) for i, col in enumerate(columns)
    ]
    rows_df = shown.head(MAX_ROWS_ON_SCREEN).select([ROW_CODE] + [c.code for c in columns])
    data = {
        "columns": payload_columns,
        "rows": _rows_json(rows_df, columns),
        "total": all_rows.height,
        "truncated": shown.height > MAX_ROWS_ON_SCREEN,
        "filters": state.get("filters") or {},
        "sort": state.get("sort"),
        "height": height,
        # Пока набор строк не изменился, таблица не перерисовывается и не сбрасывает прокрутку.
        "data_id": _data_id(key, state.get("filters")),
    }
    COMPONENT(key=key, data=data, on_view_change=lambda: _remember(key), height="content")


def _remember(key: str) -> None:
    """Колбэк: сохраняем состояние таблицы, чтобы в следующем запуске применить фильтры."""
    result = st.session_state.get(key)
    view = getattr(result, "view", None) if result is not None else None
    if view is None and isinstance(result, dict):
        view = result.get("view")
    st.session_state[STATE_KEY] = view or {}


# ---------- данные для таблицы ----------


def _row_column() -> dict:
    return {"code": ROW_CODE, "label": ROW_LABEL, "index": 0, "fmt": "int", "kind": None}


def _fmt(col: Column) -> str:
    if col.dtype == DATE:
        return "date"
    if col.percent:
        return "percent"
    if col.dtype == NUMBER and col.role in (MEASURE, RATIO):
        return "money"
    if col.dtype == INT:
        return "int"
    if col.dtype == NUMBER:
        return "money"
    return "text"


def _column_payload(df: pl.DataFrame, col: Column, filters: list[ops.Filter], index: int) -> dict:
    kind = ops.filter_kind(col)
    source = ops.rows_for_options(df, filters, col.code)
    options = None
    if kind == ops.VALUES:
        values = ops.distinct_values(source, col.code)
        if len(values) > MAX_OPTIONS:
            kind = ops.CONTAINS
        else:
            options = values
    elif kind == ops.MONTHS:
        options = {
            "years": ops.month_options(source, col.code),
            "has_empty": source.get_column(col.code).null_count() > 0,
        }
    elif kind == ops.RANGE:
        low, high = ops.value_range(source, col.code)
        options = {"min": low, "max": high}
    return {
        "code": col.code,
        "label": col.label,
        "index": index,
        "fmt": _fmt(col),
        "kind": kind,
        "options": options,
    }


def _rows_json(df: pl.DataFrame, columns: list[Column]) -> list[list]:
    """Строки списками значений; даты — текстом ГГГГ-ММ-ДД."""
    date_codes = [c.code for c in columns if c.dtype == DATE]
    if date_codes:
        df = df.with_columns(pl.col(date_codes).dt.strftime("%Y-%m-%d"))
    return [list(r) for r in df.iter_rows()]


def _data_id(key: str, filters) -> str:
    text = key + json.dumps(filters or {}, sort_keys=True, ensure_ascii=False)
    return hashlib.md5(text.encode("utf-8")).hexdigest()


# ---------- состояние таблицы -> фильтры ----------


def _to_filter(code: str, f: dict) -> ops.Filter | None:
    """Фильтр из таблицы (JSON) -> ops.Filter с правильными типами значений."""
    col = BY_CODE[code]
    kind = f.get("kind")
    if kind == ops.VALUES:
        return ops.Filter(code, kind, values=[_cast(col, v) for v in f.get("values", [])])
    if kind == ops.MONTHS:
        return ops.Filter(code, kind, values=list(f.get("values", [])))
    if kind == ops.RANGE:
        low, high = f.get("low"), f.get("high")
        return ops.Filter(
            code, kind,
            low=None if low is None else float(low),
            high=None if high is None else float(high),
        )
    if kind == ops.CONTAINS:
        return ops.Filter(code, kind, text=str(f.get("text", "")))
    return None


def _cast(col: Column, value):
    """JSON не различает 1 и 1.0, а polars различает: приводим к типу колонки."""
    if value is None:
        return None
    if col.dtype == INT:
        return int(value)
    if col.dtype == NUMBER:
        return float(value)
    if col.dtype == TEXT:
        return str(value)
    return value
