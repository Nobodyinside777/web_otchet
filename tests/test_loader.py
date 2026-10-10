"""Чтение обезличенного образца и приведение типов."""

import io
from datetime import date, datetime

import pytest

from app import loader
from app.loader import load, to_date, to_int, to_number, to_text
from app.schema import ROW_CODE
from app.validate import format_rows, plural
from tests.conftest import SAMPLE


@pytest.fixture(scope="module")
def sample():
    return load(SAMPLE)


def test_sample_loads(sample):
    df = sample.df
    assert df.height == 429  # строки с заполненным годом
    assert df.width == 27  # номер строки + колонки B–AA
    assert sample.threshold == 0.5
    assert df[ROW_CODE].is_unique().all()
    assert sample.hidden == []  # в образце скрытых колонок нет


def test_first_row_values(sample):
    row = sample.df.row(0, named=True)
    assert row[ROW_CODE] == 3
    assert row["year"] == 2026
    assert row["order_no"] == "9382734"  # число из колонки D стало текстом без «.0»
    assert row["revenue_date"] == date(2026, 1, 20)  # текст «20.01.2026» стал датой
    assert row["revenue_date_check"] == "0"
    assert row["passes_threshold"] == 1


def test_dates_are_all_recognized(sample):
    assert sample.df["revenue_date"].null_count() == 0


def test_contract_date_is_a_date(sample):
    df = sample.df
    assert str(df.schema["contract_date"]) == "Date"
    assert df.filter(df[ROW_CODE] == 343)["contract_date"][0] == date(2026, 9, 30)
    # В образце в V местами записаны суммы: это предупреждение, а не отказ читать файл.
    assert any(w.startswith("Колонка V") and "ожидались даты" in w for w in sample.warnings)


def test_total_and_draft_rows_are_skipped_with_warning(sample):
    rows = set(sample.df[ROW_CODE])
    assert not rows & {446, 447, 448, 449, 456}
    assert any("446–449, 456" in w for w in sample.warnings)


def test_excel_errors_are_reported(sample):
    # В строке 413 в колонках Q и R стоит #VALUE!
    assert sample.df.filter(sample.df[ROW_CODE] == 413)["cost"][0] is None
    assert any(w.startswith("Колонка Q") and "ошибкой Excel" in w and "413" in w for w in sample.warnings)


def test_messages_do_not_contain_cell_values(sample):
    text = " ".join(sample.warnings + sample.info)
    for value in ["9382734", "Поставлено", "Трейдинг", "174227697"]:
        assert value not in text


def test_load_from_file_object():
    result = load(io.BytesIO(SAMPLE.read_bytes()))
    assert result.df.height == 429


def test_progress_is_reported():
    seen = []
    load(SAMPLE, progress=seen.append)
    assert seen[-1] == 1.0


@pytest.mark.parametrize(
    "value, expected",
    [
        (2026.0, (2026, "ok")),
        ("2026", (2026, "ok")),
        ("  ", (None, "empty")),
        ("", (None, "empty")),
        (1.5, (None, "дробное число")),
        ("Ноябрь", (None, "текст")),
    ],
)
def test_to_int(value, expected):
    assert to_int(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (12.5, (12.5, "ok")),
        ("1 234,5", (1234.5, "ok")),
        ("1\xa0234", (1234.0, "ok")),
        ("20 000 000", (20_000_000.0, "ok")),
        ("50%", (0.5, "ok")),
        ("10,0%", (0.1, "ok")),
        (" ", (None, "empty")),
        ("abc", (None, "текст")),
        (date(2026, 1, 1), (None, "дата")),
        (True, (None, "логическое")),
    ],
)
def test_to_number(value, expected):
    assert to_number(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (" текст ", ("текст", "ok")),
        (9382734.0, ("9382734", "ok")),
        (1.5, ("1.5", "ok")),
        (date(2026, 3, 1), ("01.03.2026", "ok")),
        ("", (None, "empty")),
    ],
)
def test_to_text(value, expected):
    assert to_text(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (date(2026, 1, 20), (date(2026, 1, 20), "ok")),
        (datetime(2026, 1, 20, 15, 0), (date(2026, 1, 20), "ok")),
        ("20.01.2026", (date(2026, 1, 20), "ok")),
        ("2026-01-20", (date(2026, 1, 20), "ok")),
        (46042.0, (date(2026, 1, 20), "ok")),  # серийный номер даты Excel
        ("Ноябрь", (None, "текст")),
        ("", (None, "empty")),
        (12.0, (None, "число")),
    ],
)
def test_to_date(value, expected):
    assert to_date(value) == expected


def test_format_rows():
    assert format_rows([3, 4, 5, 9, 11, 12]) == "3–5, 9, 11–12"
    assert format_rows(list(range(0, 100, 2)), limit=3) == "0, 2, 4 и ещё 47 диапазонов"


def test_plural():
    assert [plural(n, "ячейка", "ячейки", "ячеек") for n in (1, 2, 5, 11, 21, 22, 112)] == [
        "ячейка", "ячейки", "ячеек", "ячеек", "ячейка", "ячейки", "ячеек",
    ]


def test_scan_finds_only_errors_inside_b_to_aa():
    scan = loader.scan_sheet_xml(SAMPLE, "Прогноз выручки")
    assert "AB" not in scan.error_rows and "AC" not in scan.error_rows
    assert 413 in scan.error_rows["Q"]
    assert scan.formulas_without_values == 0
