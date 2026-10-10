"""Боковая панель с фильтрами. Возвращает список ops.Filter, саму фильтрацию делает ops."""

from datetime import date

import polars as pl
import streamlit as st

from app import ops
from app.schema import BY_CODE, DATE, Column

KEY_PREFIX = "flt_"  # все виджеты фильтров; по этому префиксу их сбрасываем


def reset_filters() -> None:
    for key in [k for k in st.session_state if k.startswith(KEY_PREFIX)]:
        del st.session_state[key]


def filters_sidebar(df: pl.DataFrame, columns: list[Column]) -> list[ops.Filter]:
    st.sidebar.header("Фильтры")
    chosen = st.sidebar.multiselect(
        "Фильтровать по колонкам",
        options=[c.code for c in columns],
        format_func=lambda code: BY_CODE[code].label,
        key=KEY_PREFIX + "columns",
        placeholder="Выберите колонки",
    )
    filters = []
    for code in chosen:
        col = BY_CODE[code]
        kind = ops.filter_kind(col)
        if kind == ops.VALUES:
            f = _values_widget(df, col)
        elif kind == ops.RANGE:
            f = _range_widget(df, col)
        else:
            f = _contains_widget(col)
        if f is not None:
            filters.append(f)

    if chosen:
        st.sidebar.button("Сбросить все фильтры", on_click=reset_filters)
    return filters


def _values_widget(df: pl.DataFrame, col: Column) -> ops.Filter | None:
    options = ops.distinct_values(df, col.code)
    selected = st.sidebar.multiselect(
        col.label,
        options=options,
        format_func=lambda v: ops.format_value(col.code, v),
        key=f"{KEY_PREFIX}{col.code}",
        placeholder="Все значения",
    )
    return ops.Filter(col.code, ops.VALUES, values=selected) if selected else None


def _range_widget(df: pl.DataFrame, col: Column) -> ops.Filter | None:
    low, high = ops.value_range(df, col.code)
    if low is None:
        st.sidebar.caption(f"{col.label}: в колонке нет значений")
        return None

    if col.dtype == DATE:
        picked = st.sidebar.date_input(
            col.label,
            value=(low, high),
            min_value=date(1990, 1, 1),
            max_value=date(2100, 12, 31),
            format="DD.MM.YYYY",
            key=f"{KEY_PREFIX}{col.code}",
        )
        # Пока выбрана только одна дата диапазона, фильтр не применяем.
        if not isinstance(picked, tuple) or len(picked) != 2:
            return None
        new_low, new_high = picked
    else:
        scale = 100.0 if col.percent else 1.0  # доли вводим в процентах
        unit = ", %" if col.percent else ""
        c1, c2 = st.sidebar.columns(2)
        new_low = c1.number_input(
            f"{col.label}{unit}: от", value=float(low) * scale, key=f"{KEY_PREFIX}{col.code}_low"
        ) / scale
        new_high = c2.number_input(
            "до", value=float(high) * scale, key=f"{KEY_PREFIX}{col.code}_high"
        ) / scale

    # Границы, совпадающие с краями данных, ничего не ограничивают: не применяем их,
    # чтобы не терять строки с пустыми значениями.
    f = ops.Filter(
        col.code,
        ops.RANGE,
        low=None if _same_or_wider(new_low, low, below=True) else new_low,
        high=None if _same_or_wider(new_high, high, below=False) else new_high,
    )
    return f if f.low is not None or f.high is not None else None


def _same_or_wider(picked, edge, below: bool) -> bool:
    """Граница не уже края данных (с допуском на округление при пересчёте процентов)."""
    if isinstance(picked, date):
        return picked <= edge if below else picked >= edge
    tolerance = 1e-9 * max(1.0, abs(edge))
    return picked <= edge + tolerance if below else picked >= edge - tolerance


def _contains_widget(col: Column) -> ops.Filter | None:
    text = st.sidebar.text_input(f"{col.label} содержит", key=f"{KEY_PREFIX}{col.code}")
    return ops.Filter(col.code, ops.CONTAINS, text=text) if text.strip() else None
