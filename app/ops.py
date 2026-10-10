"""Операции над таблицей: фильтры, сортировка, итоги.

Здесь нет ничего про интерфейс. Всё, что пользователь настроил на экране, описывается
объектом ViewSettings: его можно применить к таблице (apply_view), описать словами
(describe_filters, пригодится для шапки PDF) и позже сохранить как пресет.
Новая операция = новое поле в ViewSettings + новый шаг в apply_view.
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Any

import polars as pl

from app.schema import BY_CODE, DATE, MEASURE, NUMBER, RATIO, Column

# Виды фильтров
VALUES = "values"  # галочки у значений из списка
RANGE = "range"  # от и до (числа, даты)
MONTHS = "months"  # галочки у месяцев по годам, как в Excel; значения вида "2026-01"
CONTAINS = "contains"  # текст содержит подстроку (когда разных значений слишком много)

MONTH_NAMES = (
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
)

# Колонки для расчёта маржи % по итогам: сумма МП / сумма выручки.
MARGIN_PCT, MARGIN_RUB, REVENUE = "margin_pct", "margin_rub", "revenue"


@dataclass
class Filter:
    code: str  # колонка
    kind: str  # VALUES, RANGE, MONTHS или CONTAINS
    # Для VALUES и MONTHS — отмеченные значения; None в списке = пустые ячейки.
    values: list = field(default_factory=list)
    low: Any = None  # для RANGE; None = без нижней границы
    high: Any = None  # для RANGE; None = без верхней границы
    text: str = ""  # для CONTAINS


@dataclass
class ViewSettings:
    """Всё, что настроено на экране."""

    filters: list[Filter] = field(default_factory=list)
    sort_by: str | None = None
    descending: bool = False
    totals: list[str] = field(default_factory=list)  # колонки для итогов


def filter_kind(col: Column) -> str:
    """Какой фильтр подходит колонке: даты — по месяцам, суммы — диапазон, остальное — галочки."""
    if col.dtype == DATE:
        return MONTHS
    if col.dtype == NUMBER and col.role in (MEASURE, RATIO):
        return RANGE
    return VALUES


def totals_candidates(columns: list[Column]) -> list[Column]:
    """По каким колонкам можно считать итоги: суммы и доли."""
    return [c for c in columns if c.dtype == NUMBER and c.role in (MEASURE, RATIO)]


# ---------- применение настроек ----------


def apply_view(df: pl.DataFrame, view: ViewSettings) -> pl.DataFrame:
    df = apply_filters(df, view.filters)
    return apply_sort(df, view.sort_by, view.descending)


def apply_filters(df: pl.DataFrame, filters: list[Filter]) -> pl.DataFrame:
    conditions = [c for f in filters if (c := filter_condition(f)) is not None]
    return df.filter(conditions) if conditions else df


def filter_condition(f: Filter) -> pl.Expr | None:
    """Условие polars для одного фильтра. None — фильтр ничего не ограничивает."""
    col = pl.col(f.code)
    if f.kind in (VALUES, MONTHS):
        if not f.values:
            return None
        if f.kind == MONTHS:
            col = col.dt.strftime("%Y-%m")
        chosen = [v for v in f.values if v is not None]
        cond = col.is_in(chosen) if chosen else pl.lit(False)
        if None in f.values:
            cond = cond | col.is_null()
        return cond.fill_null(False)
    if f.kind == RANGE:
        parts = []
        if f.low is not None:
            parts.append(col >= f.low)
        if f.high is not None:
            parts.append(col <= f.high)
        if not parts:
            return None
        return pl.all_horizontal(parts)
    if f.kind == CONTAINS:
        text = f.text.strip().lower()
        if not text:
            return None
        return col.str.to_lowercase().str.contains(text, literal=True).fill_null(False)
    raise ValueError(f"Неизвестный вид фильтра: {f.kind}")


def apply_sort(df: pl.DataFrame, by: str | None, descending: bool) -> pl.DataFrame:
    if not by or by not in df.columns:
        return df
    return df.sort(by, descending=descending, nulls_last=True, maintain_order=True)


# ---------- итоги ----------

STAT_NAMES = ("Сумма", "Количество", "Среднее", "Минимум", "Максимум")


def totals(df: pl.DataFrame, codes: list[str]) -> dict[str, dict[str, float | int | None]]:
    """Итоги по колонкам: {код: {«Сумма»: ..., «Количество»: ..., ...}}.

    Количество — число заполненных ячеек. Маржу % суммировать нельзя, поэтому для неё
    «Сумма» пустая, а «Среднее» — это маржа по итогам (сумма МП / сумма выручки).
    """
    codes = [c for c in codes if c in df.columns]
    if not codes:
        return {}
    exprs = []
    for code in codes:
        col = pl.col(code)
        exprs += [
            col.sum().alias(f"{code}|Сумма"),
            col.count().alias(f"{code}|Количество"),
            col.mean().alias(f"{code}|Среднее"),
            col.min().alias(f"{code}|Минимум"),
            col.max().alias(f"{code}|Максимум"),
        ]
    row = df.select(exprs).row(0, named=True)
    result = {code: {name: row[f"{code}|{name}"] for name in STAT_NAMES} for code in codes}

    for code in codes:
        if BY_CODE[code].role == RATIO:
            result[code]["Сумма"] = None
            if code == MARGIN_PCT:
                result[code]["Среднее"] = weighted_margin(df)
    return result


def weighted_margin(df: pl.DataFrame) -> float | None:
    """Маржа % по итогам: сумма МП / сумма выручки. None, если посчитать нельзя."""
    if MARGIN_RUB not in df.columns or REVENUE not in df.columns:
        return None
    sums = df.select(pl.col(MARGIN_RUB).sum(), pl.col(REVENUE).sum()).row(0)
    margin, revenue = sums
    if not revenue:
        return None
    return margin / revenue


# ---------- справочные функции для интерфейса ----------


def distinct_values(df: pl.DataFrame, code: str) -> list:
    """Уникальные значения колонки по возрастанию; пустое значение (None) — в конце."""
    values = df.get_column(code).unique().drop_nulls().sort().to_list()
    if df.get_column(code).null_count():
        values.append(None)
    return values


def value_range(df: pl.DataFrame, code: str) -> tuple[Any, Any]:
    col = df.get_column(code)
    return col.min(), col.max()


def rows_for_options(df: pl.DataFrame, filters: list[Filter], code: str) -> pl.DataFrame:
    """Строки, из которых берутся значения для окна фильтра колонки code.

    Как в Excel: учитываются фильтры всех других колонок, но не фильтр самой колонки,
    иначе снятые галочки пропали бы из списка и их нельзя было бы вернуть.
    """
    return apply_filters(df, [f for f in filters if f.code != code])


def month_options(df: pl.DataFrame, code: str) -> list[dict]:
    """Годы и месяцы, которые есть в колонке дат: [{"year": 2026, "months": [1, 2, ...]}, ...]."""
    col = pl.col(code)
    pairs = (
        df.select(col.dt.year().alias("y"), col.dt.month().alias("m"))
        .drop_nulls()
        .unique()
        .sort("y", "m")
    )
    years: dict[int, list[int]] = {}
    for y, m in pairs.iter_rows():
        years.setdefault(y, []).append(m)
    return [{"year": y, "months": ms} for y, ms in years.items()]


def describe_months(values: list) -> str:
    """["2026-01", "2026-02", None] -> 'январь 2026, февраль 2026, (пусто)'."""
    parts = []
    for v in values:
        if v is None:
            parts.append("(пусто)")
        else:
            year, month = v.split("-")
            parts.append(f"{MONTH_NAMES[int(month) - 1]} {year}")
    return ", ".join(parts)


def describe_filters(filters: list[Filter], fmt=str) -> list[str]:
    """Фильтры словами: «Статус заказа: Поставлено, Заявка». fmt(code, value) форматирует значение."""
    lines = []
    for f in filters:
        label = BY_CODE[f.code].label
        if f.kind == VALUES and f.values:
            shown = ", ".join("(пусто)" if v is None else fmt(f.code, v) for v in f.values)
            lines.append(f"{label}: {shown}")
        elif f.kind == MONTHS and f.values:
            lines.append(f"{label}: {describe_months(f.values)}")
        elif f.kind == RANGE and (f.low is not None or f.high is not None):
            parts = []
            if f.low is not None:
                parts.append(f"от {fmt(f.code, f.low)}")
            if f.high is not None:
                parts.append(f"до {fmt(f.code, f.high)}")
            lines.append(f"{label}: {' '.join(parts)}")
        elif f.kind == CONTAINS and f.text.strip():
            lines.append(f"{label} содержит «{f.text.strip()}»")
    return lines


def format_value(code: str, value) -> str:
    """Значение для показа человеку: даты ДД.ММ.ГГГГ, доли в процентах, числа с разрядами."""
    if value is None:
        return "(пусто)"
    col = BY_CODE.get(code)
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    if col is not None and col.percent and isinstance(value, (int, float)):
        return format_number(value * 100, 1) + " %"
    if isinstance(value, float):
        return format_number(value, 0 if value.is_integer() else 2)
    if isinstance(value, int) and col is not None and col.role in (MEASURE, RATIO):
        return format_number(value, 0)
    return str(value)


def format_short(code: str, value) -> str:
    """Коротко для карточек: суммы от миллиона — в миллионах, '-147 392,9 млн'."""
    col = BY_CODE.get(code)
    if isinstance(value, (int, float)) and not (col and col.percent) and abs(value) >= 1_000_000:
        return format_number(value / 1_000_000, 1) + " млн"
    return format_value(code, value)


def format_number(value: float, decimals: int = 0) -> str:
    """1234567.891 -> '1 234 567,89' (неразрывный пробел между разрядами)."""
    text = f"{value:,.{decimals}f}"
    return text.replace(",", " ").replace(".", ",")
