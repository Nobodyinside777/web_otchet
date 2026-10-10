"""Общие заготовки для тестов.

Работаем только с обезличенным образцом. «Испорченные» файлы делаем из его копии
во временной папке: openpyxl для этого годится, файл маленький.
"""

from pathlib import Path

import openpyxl
import pytest

SAMPLE = Path(__file__).resolve().parents[1] / "samples" / "be2026_anon.xlsx"
SHEET = "Прогноз выручки"


@pytest.fixture
def make_variant(tmp_path):
    """Возвращает функцию: change(ws) правит копию листа, результат — путь к новому файлу."""

    def build(change, name="variant.xlsx"):
        wb = openpyxl.load_workbook(SAMPLE)
        change(wb[SHEET])
        path = tmp_path / name
        wb.save(path)
        return path

    return build
