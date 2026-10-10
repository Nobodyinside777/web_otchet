"""Интерфейс Streamlit. Здесь только экран; логика живёт в соседних модулях."""

import sys
from pathlib import Path

import polars as pl
import streamlit as st

# Streamlit запускает этот файл как скрипт, поэтому папку проекта добавляем в путь сами.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import ops  # noqa: E402
from app.loader import LoadResult, load  # noqa: E402
from app.schema import BY_CODE, Column  # noqa: E402
from app.ui_table import current_view, reset_view, show_table  # noqa: E402
from app.validate import FormatError, hidden_columns_message  # noqa: E402

DEFAULT_TOTALS = ["revenue", "cost", "margin_rub", "margin_pct"]
TABLE_HEIGHT = 600  # высота таблицы на странице, пикселей; кнопка «На весь экран» — над ней

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
    reset_view()
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

# Фильтры и сортировку пользователь задаёт в заголовках таблицы (как в Excel).
table_view = current_view()
view = ops.ViewSettings(
    filters=table_view.filters, sort_by=table_view.sort_by, descending=table_view.descending
)
filtered = ops.apply_filters(result.df, view.filters)

if view.filters:
    st.markdown(
        f"**Показано строк: {filtered.height} из {result.df.height}.** Фильтры: "
        + "; ".join(ops.describe_filters(view.filters, ops.format_value))
    )
else:
    st.markdown(f"**Показаны все строки: {filtered.height}.**")

view.totals = totals_block(filtered, result.columns)

show_table(
    result.df,
    filtered,
    result.columns,
    table_view,
    key=f"table_{uploaded.file_id}",
    height=TABLE_HEIGHT,
)
st.caption(
    "Фильтр — кнопка ▾ в заголовке колонки. Сортировка — щелчок по названию колонки "
    "(по возрастанию → по убыванию → как в файле)."
)

# Текущие настройки пригодятся для выгрузки и пресетов на следующих этапах.
st.session_state["view"] = view
