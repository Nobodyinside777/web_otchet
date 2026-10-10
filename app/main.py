"""Интерфейс Streamlit. Здесь только экран; логика живёт в соседних модулях."""

import sys
from pathlib import Path

import polars as pl
import streamlit as st

# Streamlit запускает этот файл как скрипт, поэтому папку проекта добавляем в путь сами.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import ops  # noqa: E402
from app.loader import LoadResult, load  # noqa: E402
from app.schema import BY_CODE, DATE, LABELS, MEASURE, TEXT, Column  # noqa: E402
from app.ui_filters import filters_button, reset_filters  # noqa: E402
from app.validate import FormatError, hidden_columns_message  # noqa: E402

# Больше строк браузер показывает медленно. Итоги всё равно считаются по всем строкам.
MAX_ROWS_ON_SCREEN = 100_000
TABLE_HEIGHT = 600  # высота таблицы на странице, пикселей
DEFAULT_TOTALS = ["revenue", "cost", "margin_rub", "margin_pct"]

st.set_page_config(page_title="Прогноз выручки", layout="wide")


# ---------- загрузка ----------


def load_uploaded(uploaded) -> LoadResult | None:
    """Читает загруженный файл один раз и запоминает результат до загрузки другого файла."""
    cache = st.session_state.get("loaded")
    if cache and cache["file_id"] == uploaded.file_id:
        return cache["result"]

    bar = st.progress(0.0, text="Читаю файл…")
    try:
        result = load(uploaded, progress=lambda share: bar.progress(share, text="Читаю файл…"))
    except FormatError as exc:
        bar.empty()
        st.error("Файл не подходит по формату:\n\n" + "\n".join(f"- {p}" for p in exc.problems))
        st.session_state.pop("loaded", None)
        return None
    bar.empty()
    st.session_state["loaded"] = {"file_id": uploaded.file_id, "result": result}
    # Фильтры и сортировка от прошлого файла к новому могут не подходить.
    reset_filters()
    for key in [k for k in st.session_state if k.startswith("view_")]:
        del st.session_state[key]
    return result


def show_file_check(result: LoadResult) -> None:
    """Итог проверки файла; подробности свёрнуты, чтобы не занимать экран."""
    problems = len(result.warnings)
    title = f"Файл подходит: строк с данными — {result.df.height}"
    if problems:
        title += f". Предупреждений: {problems}"
    with st.expander(title, expanded=False, icon="✅" if not problems else "⚠️"):
        threshold = result.threshold if result.threshold is not None else "не задан"
        st.markdown(f"Порог вероятности (Z1): {threshold}.")
        if result.hidden:
            st.markdown(hidden_columns_message(result.hidden))
        for w in result.warnings:
            st.markdown(f"- {w}")
        if result.info:
            st.markdown("**Пустые ячейки в суммах:**")
            st.markdown("\n".join(f"- {i}" for i in result.info))


# ---------- таблица ----------


def column_config(columns: list[Column]) -> dict:
    """Как показывать колонки: даты без времени, доли в процентах, суммы с разрядами."""
    config = {}
    for col in columns:
        if col.dtype == DATE:
            config[col.label] = st.column_config.DateColumn(col.label, format="DD.MM.YYYY")
        elif col.percent:
            config[col.label] = st.column_config.NumberColumn(col.label, format="percent")
        elif col.role == MEASURE:
            config[col.label] = st.column_config.NumberColumn(col.label, format="localized")
    return config


def for_screen(df: pl.DataFrame, columns: list[Column]) -> pl.DataFrame:
    """Таблица для показа: пустой текст вместо None и русские заголовки."""
    text_codes = [c.code for c in columns if c.dtype == TEXT]
    if text_codes:
        df = df.with_columns(pl.col(text_codes).fill_null(""))
    return df.rename({code: LABELS[code] for code in df.columns})


# ---------- итоги ----------


def totals_block(df: pl.DataFrame, columns: list[Column]) -> list[str]:
    candidates = [c.code for c in ops.totals_candidates(columns)]
    if not candidates:
        return []
    chosen = st.multiselect(
        "Итоги по колонкам",
        options=candidates,
        default=[c for c in DEFAULT_TOTALS if c in candidates],
        format_func=lambda code: BY_CODE[code].label,
        key="view_totals",
    )
    stats = ops.totals(df, chosen)
    if not stats:
        return chosen

    # Карточки с главным числом: сумма, а для маржи % — маржа по итогам.
    cards = st.columns(len(stats))
    for card, (code, s) in zip(cards, stats.items()):
        main = s["Среднее"] if BY_CODE[code].role == "ratio" else s["Сумма"]
        if main is None:
            card.metric(BY_CODE[code].label, "—")
            continue
        card.metric(
            BY_CODE[code].label,
            ops.format_short(code, main),
            help="Точно: " + ops.format_value(code, float(main)),
        )

    with st.expander("Подробно: сумма, количество, среднее, минимум, максимум"):
        table = {"Показатель": list(ops.STAT_NAMES)}
        for code, s in stats.items():
            table[BY_CODE[code].label] = [_stat_text(code, name, s[name]) for name in ops.STAT_NAMES]
        st.dataframe(pl.DataFrame(table), hide_index=True, width="stretch")
        if ops.MARGIN_PCT in stats:
            st.caption(
                "Маржу % нельзя складывать, поэтому для неё «Среднее» — это маржа по итогам: "
                "сумма МП / сумма выручки по показанным строкам."
            )
    return chosen


def _stat_text(code: str, name: str, value) -> str:
    if value is None:
        return "—"
    if name == "Количество":
        return ops.format_number(value, 0)
    return ops.format_value(code, float(value))


# ---------- страница ----------

st.title("Прогноз выручки")
uploaded = st.file_uploader(
    "Excel-файл с листом «Прогноз выручки»",
    type=["xlsx", "xlsm"],
    help="Перетащите файл сюда или нажмите, чтобы выбрать.",
)

if uploaded is None:
    st.info("Загрузите файл, чтобы начать.")
    st.stop()

result = load_uploaded(uploaded)
if result is None:
    st.stop()

show_file_check(result)

# Строка над итогами: слева — сколько строк показано, справа — кнопка «Фильтры».
info_col, button_col = st.columns([8, 2], vertical_alignment="center")
with button_col:
    with st.container(horizontal_alignment="right"):
        view = ops.ViewSettings(filters=filters_button(result.df, result.columns))
shown = ops.apply_view(result.df, view)

if view.filters:
    info_col.markdown(
        f"**Показано строк: {shown.height} из {result.df.height}.** Фильтры: "
        + "; ".join(ops.describe_filters(view.filters, ops.format_value))
    )
else:
    info_col.markdown(f"**Показаны все строки: {shown.height}.**")

view.totals = totals_block(shown, result.columns)

if shown.height > MAX_ROWS_ON_SCREEN:
    st.caption(
        f"В таблице показаны первые {MAX_ROWS_ON_SCREEN:,} строк из {shown.height:,}. "
        "Итоги посчитаны по всем строкам.".replace(",", " ")
    )
st.dataframe(
    for_screen(shown.head(MAX_ROWS_ON_SCREEN), result.columns),
    column_config=column_config(result.columns),
    width="stretch",
    hide_index=True,
    height="stretch",
)
st.caption(
    "Сортировка — щелчок по заголовку колонки. "
    "Кнопка ⛶ появляется над таблицей при наведении мыши и открывает её на весь экран."
)

# Текущие настройки пригодятся для выгрузки и пресетов на следующих этапах.
st.session_state["view"] = view
