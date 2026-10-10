"""Перевод того, что выбрал пользователь в таблице, в фильтры и подготовка данных для таблицы."""

from datetime import date

import polars as pl

from app import ops, ui_table
from app.loader import load
from app.schema import BY_CODE
from tests.conftest import SAMPLE


def test_values_are_cast_to_column_type():
    # JSON превращает 1.0 в 1; для колонки с дробными числами возвращаем 1.0.
    f = ui_table._to_filter("probability", {"kind": "values", "values": [1, 0.5, None]})
    assert f.values == [1.0, 0.5, None] and isinstance(f.values[0], float)
    f = ui_table._to_filter("month", {"kind": "values", "values": [12.0]})
    assert f.values == [12] and isinstance(f.values[0], int)


def test_other_kinds():
    assert ui_table._to_filter("revenue", {"kind": "range", "low": 0, "high": None}) == ops.Filter(
        "revenue", "range", low=0.0, high=None
    )
    assert ui_table._to_filter("revenue_date", {"kind": "months", "values": ["2026-01"]}).values == ["2026-01"]
    assert ui_table._to_filter("description", {"kind": "contains", "text": "abc"}).text == "abc"
    assert ui_table._to_filter("description", {"kind": "???"}) is None


def test_payload_on_sample():
    df = load(SAMPLE).df
    status = ui_table._column_payload(df, BY_CODE["order_status"], [], 1)
    assert status["kind"] == "values" and "Поставлено" in status["options"]
    dates = ui_table._column_payload(df, BY_CODE["revenue_date"], [], 2)
    assert dates["kind"] == "months" and dates["options"]["years"][0]["year"] == 2025
    revenue = ui_table._column_payload(df, BY_CODE["revenue"], [], 3)
    assert revenue["kind"] == "range" and revenue["options"]["min"] < 0 < revenue["options"]["max"]


def test_too_many_values_become_text_search(monkeypatch):
    monkeypatch.setattr(ui_table, "MAX_OPTIONS", 3)
    df = pl.DataFrame({"description": ["a", "b", "c", "d"]})
    assert ui_table._column_payload(df, BY_CODE["description"], [], 1)["kind"] == "contains"


def test_rows_json_formats_dates():
    df = pl.DataFrame({"excel_row": [3], "revenue_date": [date(2026, 1, 20)]})
    assert ui_table._rows_json(df, [BY_CODE["revenue_date"]]) == [[3, "2026-01-20"]]
