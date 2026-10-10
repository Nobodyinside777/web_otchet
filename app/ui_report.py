"""Блок «Презентация»: дата отчёта, плановые цифры и кнопка скачивания PPTX.

Отчёт строится по строкам, оставшимся после фильтров на сайте:
общий слайд по всем этим строкам, затем слайд и детализация по каждому МВЗ.
"""

import json
from datetime import date
from pathlib import Path

import polars as pl
import streamlit as st

from app.report import Plan, ReportParams, build_report, mvz_list
from app.report_pptx import build_pptx

# Плановые цифры вводятся раз в год. Запоминаем их рядом с программой, чтобы не вводить каждый раз.
# В репозиторий файл не попадает (presets/*.json в .gitignore).
PRESETS = Path(__file__).resolve().parents[1] / "presets"
PLAN_FILE = PRESETS / "plan.json"
# Тексты титульного слайда (название отчёта, подразделение, компания) тоже хранятся только здесь.
TITLE_FILE = PRESETS / "title.json"
TITLE_FIELDS = ["Заголовок титула", "Подзаголовок титула", "Подпись внизу титула"]
MLN = 1_000_000

PLAN_COLUMNS = {
    "revenue_forecast": "Выручка, {forecast}",
    "margin_forecast": "МП, {forecast}",
    "revenue_budget": "Выручка, {budget}",
    "margin_budget": "МП, {budget}",
}


def load_json(path: Path) -> dict:
    """Сохранённые настройки. Пустой словарь, если файла нет или он испорчен."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def title_inputs() -> tuple[str, str, str]:
    """Три поля для текстов титула; введённое запоминается."""
    saved = load_json(TITLE_FILE)
    columns = st.columns(3)
    texts = tuple(
        col.text_input(label, value=saved.get(label, ""), key=f"report_title_{i}")
        for i, (col, label) in enumerate(zip(columns, TITLE_FIELDS))
    )
    current = dict(zip(TITLE_FIELDS, texts))
    if current != {k: saved.get(k, "") for k in TITLE_FIELDS}:
        save_json(TITLE_FILE, current)
    return texts


def plan_editor(mvz: list[str], forecast: str, budget: str) -> dict[str, Plan]:
    """Таблица для ввода плана по каждому МВЗ (млн руб.). Возвращает планы в рублях."""
    saved = load_json(PLAN_FILE)
    labels = {code: text.format(forecast=forecast, budget=budget) for code, text in PLAN_COLUMNS.items()}
    table = pl.DataFrame(
        {
            "МВЗ": mvz,
            **{
                labels[code]: [float(saved.get(m, {}).get(code, 0.0)) for m in mvz]
                for code in PLAN_COLUMNS
            },
        }
    )
    edited = st.data_editor(
        table,
        hide_index=True,
        disabled=["МВЗ"],
        width="stretch",
        column_config={
            label: st.column_config.NumberColumn(label, format="%.1f", step=0.1) for label in labels.values()
        },
        key="report_plan_editor",
    )

    plans, to_save = {}, dict(saved)
    for row in edited.iter_rows(named=True):
        values = {code: float(row[labels[code]] or 0.0) for code in PLAN_COLUMNS}
        to_save[row["МВЗ"]] = values
        plans[row["МВЗ"]] = Plan(**{code: v * MLN for code, v in values.items()})
    if to_save != saved:
        save_json(PLAN_FILE, to_save)
    return plans


def report_block(shown: pl.DataFrame, filters_text: list[str]) -> None:
    # key: Streamlit запоминает, раскрыт ли блок, и не сворачивает его после нажатия кнопки.
    with st.expander("Презентация PPTX", icon=":material/slideshow:", key="report_expander"):
        if shown.height == 0:
            st.info("После фильтров не осталось строк — презентацию собрать не из чего.")
            return

        mvz = mvz_list(shown)
        scope = "все строки файла" if not filters_text else "строки после фильтров: " + "; ".join(filters_text)
        st.markdown(
            f"Отчёт строится по показанным строкам ({scope}). "
            f"Слайды по МВЗ: {', '.join(mvz)}. Чтобы оставить только нужные МВЗ, отфильтруйте колонку «МВЗ»."
        )

        # Форма: изменения даты и плана применяются одним нажатием кнопки,
        # страница не перезагружается после каждой ячейки.
        with st.form("report_form", border=False):
            left, right = st.columns(2)
            report_date = left.date_input("Дата отчёта", value=date.today(), format="DD.MM.YYYY", key="report_date")
            forecast_label = right.text_input(
                "Название колонки прогноза",
                placeholder="по отчётному месяцу, в октябре — «БП 9+3»",
                help="Оставьте пустым, чтобы название считалось от месяца отчёта.",
                key="report_forecast_label",
            )
            params = ReportParams(report_date, forecast_label=forecast_label)
            params.title_texts = title_inputs()

            st.markdown("**План по МВЗ, млн руб.** Заполняется раз в год и запоминается.")
            params.plans = plan_editor(mvz, params.label_forecast(), params.label_budget())
            submitted = st.form_submit_button("Собрать презентацию", type="primary")

        # «Отпечаток» настроек: если после сборки поменяли фильтры, старый файл не отдаём.
        settings = repr((filters_text, shown.height, params))

        if submitted:
            with st.spinner("Собираю презентацию…"):
                report = build_report(shown, params)
                st.session_state["report_pptx"] = build_pptx(report)
                st.session_state["report_pptx_name"] = f"Отчёт_БЕ_АСУТП_{report_date:%Y-%m-%d}.pptx"
                st.session_state["report_pptx_settings"] = settings

        if "report_pptx" not in st.session_state:
            return
        if st.session_state.get("report_pptx_settings") != settings:
            st.caption("Фильтры или план изменились после сборки. Нажмите «Собрать презентацию» ещё раз.")
        else:
            st.download_button(
                "Скачать PPTX",
                data=st.session_state["report_pptx"],
                file_name=st.session_state["report_pptx_name"],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                icon=":material/download:",
            )
