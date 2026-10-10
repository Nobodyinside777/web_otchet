"""Испорченные файлы должны давать понятную ошибку, а не падение."""

import io
from datetime import date

import pytest

from app.loader import load
from app.validate import FormatError


def problems_of(source) -> str:
    with pytest.raises(FormatError) as info:
        load(source)
    return "\n".join(info.value.problems)


def test_missing_sheet(make_variant):
    def rename(ws):
        ws.title = "Другой лист"

    text = problems_of(make_variant(rename))
    assert "нет листа «Прогноз выручки»" in text
    assert "«Другой лист»" in text


def test_wrong_header(make_variant):
    def change(ws):
        ws["R2"] = "Маржа"

    text = problems_of(make_variant(change))
    assert "Колонка R: ожидался заголовок «МП, руб.», найдено «Маржа»." in text


def test_empty_header(make_variant):
    def change(ws):
        ws["G2"] = None

    assert "Колонка G: ожидался заголовок «Тип закза», найдено пустая ячейка." in problems_of(
        make_variant(change)
    )


def test_extra_spaces_in_headers_are_fine(make_variant):
    def change(ws):
        ws["B2"] = "  ГОД  "
        ws["S2"] = "Маржа\n\n (Gross   profit), %"

    assert load(make_variant(change)).df.height == 429


def test_text_instead_of_numbers(make_variant):
    def change(ws):
        for row in range(3, 446):
            ws.cell(row=row, column=16).value = "много"  # колонка P

    text = problems_of(make_variant(change))
    assert "Колонка P" in text and "ожидались числа" in text


def test_formulas_without_saved_values(make_variant):
    def change(ws):
        for row in range(3, 446):
            ws.cell(row=row, column=2).value = "=2026"  # openpyxl не сохраняет значения формул

    assert "Откройте файл в Excel, сохраните и загрузите снова" in problems_of(make_variant(change))


def test_only_headers(make_variant):
    def change(ws):
        ws.delete_rows(3, ws.max_row)

    assert "нет строк с данными" in problems_of(make_variant(change))


def test_not_an_excel_file():
    assert "Не удалось открыть файл" in problems_of(io.BytesIO(b"this is not a spreadsheet"))


def test_row_typed_as_displayed_text(make_variant):
    """Строка, где значения записаны так, как их видно на экране Excel, и год не заполнен."""
    values = {
        "B": None, "C": 12, "D": "203/1", "E": "Нет   ", "K": " ",
        "N": "50%", "O": "31.12.2026", "P": "20 000 000", "Q": "18 000 000",
        "R": "2 000 000", "S": "10,0%", "T": 10, "U": "5 раздел",
        "V": 46042,  # дата 20.01.2026, сохранённая числом
    }

    def change(ws):
        for letter, value in values.items():
            ws[f"{letter}3"] = value

    df = load(make_variant(change)).df
    row = df.row(0, named=True)
    assert row["excel_row"] == 3  # строка без года не потерялась
    assert row["year"] is None
    assert row["order_no"] == "203/1"
    assert row["asutp_no"] == "Нет"
    assert row["order_status"] is None
    assert row["probability"] == 0.5
    assert row["revenue_date"] == date(2026, 12, 31)
    assert row["revenue"] == 20_000_000
    assert row["margin_pct"] == pytest.approx(0.1)
    assert row["contract_date"] == date(2026, 1, 20)


def hide(ws, *letters):
    for letter in letters:
        ws.column_dimensions[letter].hidden = True


def test_hidden_columns_are_left_out(make_variant):
    def change(ws):
        ws.column_dimensions.group("V", "X", hidden=True)  # диапазон скрытых колонок
        hide(ws, "B", "Z")

    result = load(make_variant(change))
    assert [c.letter for c in result.hidden] == ["B", "V", "W", "X", "Z"]
    for code in ("year", "contract_date", "cost2", "margin_rub2", "passes_threshold"):
        assert code not in result.df.columns
    assert "revenue" in result.df.columns and "project_comment" in result.df.columns
    assert result.df.height == 429  # без колонки года строки находятся по месяцу и дате
    assert result.threshold == 0.5  # порог из Z1 читается, даже если Z скрыта
    assert not any(w.startswith("Колонка V") for w in result.warnings)


def test_hidden_column_header_is_not_checked(make_variant):
    def change(ws):
        hide(ws, "W")
        ws["W2"] = "Что угодно"
        for row in range(3, 446):
            ws[f"W{row}"] = "текст вместо суммы"

    assert "cost2" not in load(make_variant(change)).df.columns


def test_all_anchor_columns_hidden(make_variant):
    def change(ws):
        hide(ws, "B", "C", "I", "O")

    assert "скрыты в Excel" in problems_of(make_variant(change))


def test_empty_column_a_does_not_shift_columns(make_variant):
    # Если колонка A совсем пустая, calamine начинает данные с B. Колонки не должны съехать.
    def clear_a(ws):
        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=1).value = None

    result = load(make_variant(clear_a))
    assert result.df.height == 429
    assert result.df["order_no"][0] == "9382734"
