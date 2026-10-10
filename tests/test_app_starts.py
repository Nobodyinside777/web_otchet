from pathlib import Path

from streamlit.testing.v1 import AppTest

MAIN = Path(__file__).resolve().parents[1] / "app" / "main.py"


def test_app_starts_without_errors():
    at = AppTest.from_file(str(MAIN)).run()
    assert not at.exception
    assert at.title[0].value == "Прогноз выручки"
