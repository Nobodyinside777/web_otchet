"""Фильтры, сортировка и итоги на небольшой синтетической таблице и на образце."""

from datetime import date

import polars as pl
import pytest

from app import ops
from app.loader import load
from app.ops import CONTAINS, MONTHS, RANGE, VALUES, Filter, ViewSettings
from app.schema import BY_CODE
from tests.conftest import SAMPLE


@pytest.fixture
def df():
    return pl.DataFrame(
        {
            "excel_row": [3, 4, 5, 6, 7],
            "order_status": ["Поставлено", "Заявка", None, "Поставлено", "ТКП"],
            "revenue": [100.0, 200.0, 300.0, None, 400.0],
            "margin_rub": [10.0, 20.0, 30.0, 5.0, 40.0],
            "margin_pct": [0.1, 0.1, 0.1, None, 0.1],
            "revenue_date": [date(2026, 1, 1), date(2026, 2, 1), date(2026, 3, 1), None, date(2026, 5, 1)],
            "description": ["Поставка насосов", "Работы", None, "НАСОСЫ", "Прочее"],
        }
    )


def rows(df):
    return df["excel_row"].to_list()


def test_values_filter(df):
    f = Filter("order_status", VALUES, values=["Поставлено"])
    assert rows(ops.apply_filters(df, [f])) == [3, 6]


def test_values_filter_with_empty(df):
    f = Filter("order_status", VALUES, values=["ТКП", None])
    assert rows(ops.apply_filters(df, [f])) == [5, 7]


def test_range_filter_numbers_drops_empty(df):
    f = Filter("revenue", RANGE, low=150.0, high=350.0)
    assert rows(ops.apply_filters(df, [f])) == [4, 5]


def test_range_filter_open_end(df):
    assert rows(ops.apply_filters(df, [Filter("revenue", RANGE, low=300.0)])) == [5, 7]


def test_range_filter_dates(df):
    f = Filter("revenue_date", RANGE, low=date(2026, 2, 1), high=date(2026, 3, 31))
    assert rows(ops.apply_filters(df, [f])) == [4, 5]


def test_contains_ignores_case(df):
    f = Filter("description", CONTAINS, text="насос")
    assert rows(ops.apply_filters(df, [f])) == [3, 6]


def test_filters_combine_with_and(df):
    filters = [
        Filter("order_status", VALUES, values=["Поставлено"]),
        Filter("description", CONTAINS, text="поставка"),
    ]
    assert rows(ops.apply_filters(df, filters)) == [3]


def test_empty_filters_change_nothing(df):
    filters = [Filter("order_status", VALUES), Filter("revenue", RANGE), Filter("description", CONTAINS, text=" ")]
    assert rows(ops.apply_filters(df, filters)) == [3, 4, 5, 6, 7]


def test_sort_puts_empty_last(df):
    assert rows(ops.apply_sort(df, "revenue", descending=True)) == [7, 5, 4, 3, 6]
    assert rows(ops.apply_sort(df, "revenue", descending=False)) == [3, 4, 5, 7, 6]
    assert rows(ops.apply_sort(df, None, False)) == [3, 4, 5, 6, 7]


def test_apply_view(df):
    view = ViewSettings(filters=[Filter("order_status", VALUES, values=["Поставлено", "Заявка"])], sort_by="revenue", descending=True)
    assert rows(ops.apply_view(df, view)) == [4, 3, 6]


def test_totals(df):
    t = ops.totals(df, ["revenue", "margin_pct"])
    assert t["revenue"] == {"Сумма": 1000.0, "Количество": 4, "Среднее": 250.0, "Минимум": 100.0, "Максимум": 400.0}
    # Маржа % не суммируется, «Среднее» — это сумма МП / сумма выручки.
    assert t["margin_pct"]["Сумма"] is None
    assert t["margin_pct"]["Среднее"] == pytest.approx(105.0 / 1000.0)


def test_weighted_margin_without_revenue():
    assert ops.weighted_margin(pl.DataFrame({"margin_rub": [1.0], "revenue": [0.0]})) is None
    assert ops.weighted_margin(pl.DataFrame({"margin_rub": [1.0]})) is None


def test_distinct_values_puts_empty_last(df):
    assert ops.distinct_values(df, "order_status") == ["Заявка", "Поставлено", "ТКП", None]


def test_describe_filters():
    filters = [
        Filter("order_status", VALUES, values=["Заявка", None]),
        Filter("revenue", RANGE, low=1000.0),
        Filter("margin_pct", RANGE, high=0.25),
        Filter("revenue_date", RANGE, low=date(2026, 1, 1), high=date(2026, 3, 31)),
        Filter("description", CONTAINS, text="насос"),
    ]
    assert ops.describe_filters(filters, ops.format_value) == [
        "Статус заказа: Заявка, (пусто)",
        "Выручка, руб. (без НДС): от 1 000",
        "Маржа (Gross profit), %: до 25,0 %",
        "Дата выручки: от 01.01.2026 до 31.03.2026",
        "Описание содержит «насос»",
    ]


def test_format_short():
    assert ops.format_short("revenue", -147_392_912_940.5) == "-147 392,9 млн"
    assert ops.format_short("revenue", 999_999.0) == "999 999"
    assert ops.format_short("margin_pct", 0.115) == "11,5 %"


def test_format_number():
    assert ops.format_number(1234567.891, 2) == "1 234 567,89"
    assert ops.format_number(-5.0, 0) == "-5"


def test_filter_kinds():
    assert ops.filter_kind(BY_CODE["order_status"]) == VALUES
    assert ops.filter_kind(BY_CODE["probability"]) == VALUES
    assert ops.filter_kind(BY_CODE["description"]) == VALUES
    assert ops.filter_kind(BY_CODE["revenue"]) == RANGE
    assert ops.filter_kind(BY_CODE["margin_pct"]) == RANGE
    assert ops.filter_kind(BY_CODE["revenue_date"]) == MONTHS


def test_months_filter(df):
    f = Filter("revenue_date", MONTHS, values=["2026-01", "2026-03"])
    assert rows(ops.apply_filters(df, [f])) == [3, 5]
    f = Filter("revenue_date", MONTHS, values=["2026-05", None])
    assert rows(ops.apply_filters(df, [f])) == [6, 7]


def test_month_options(df):
    assert ops.month_options(df, "revenue_date") == [{"year": 2026, "months": [1, 2, 3, 5]}]


def test_options_ignore_own_filter(df):
    """Как в Excel: в списке колонки видны значения с учётом фильтров других колонок, но не своего."""
    filters = [
        Filter("order_status", VALUES, values=["Поставлено"]),
        Filter("revenue_date", MONTHS, values=["2026-01"]),
    ]
    own = ops.rows_for_options(df, filters, "order_status")
    assert ops.distinct_values(own, "order_status") == ["Поставлено"]  # только строки января
    dates = ops.rows_for_options(df, filters, "revenue_date")
    assert rows(dates) == [3, 6]  # все «Поставлено», любые месяцы


def test_describe_months():
    assert ops.describe_months(["2026-01", "2027-12", None]) == "январь 2026, декабрь 2027, (пусто)"


def test_on_sample():
    data = load(SAMPLE).df
    delivered = ops.apply_filters(data, [Filter("order_status", VALUES, values=["Поставлено"])])
    assert delivered.height == 197
    t = ops.totals(delivered, ["revenue"])
    assert t["revenue"]["Сумма"] == pytest.approx(delivered["revenue"].sum())
