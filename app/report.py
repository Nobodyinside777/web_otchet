"""Расчёт цифр для презентации «Регулярная отчётность» (PPTX).

Здесь только числа и списки строк, без оформления. Слайды из них собирает report_pptx.py.
Все суммы считаются в рублях, в миллионы переводит уже оформление.

Какие строки идут в отчёт: те, что остались после фильтров на сайте.
- Слайд «Исполнение БП … БЕ АСУТП» — по всем этим строкам.
- Дальше по слайду на каждое МВЗ, которое есть среди этих строк.
- В конце детализация по каждому такому МВЗ.

Правила (со слов пользователя):
- «Признано» — статус заказа «Поставлено».
- «Законтракт.» — вероятность 100 % и любой статус, кроме «Поставлено».
- «Высокая вероятность» — вероятность ровно 50 %.
- «Итого» — сумма этих трёх колонок; «Отклонение» = Итого − БП'26.
- «Потенциал» — вероятность 25 %.
- Эти колонки считаются по году выручки, равному году отчёта.
- «БП 9+3» и «БП'26» не берутся из файла: их вводит пользователь при выгрузке.
- Воронка по месяцам: «<месяц> Факт» — месяц контрактации = отчётный месяц и вероятность 100 %;
  дальше отчётный месяц, +1 и +2 — сделки с вероятностью 50 % с этим месяцем контрактации.
  В месяце контрактации нет года, поэтому в воронку идут сделки с выручкой в году отчёта и позже.
- «План по выручке» — те же правила, но по месяцу выручки (и году выручки).
- Детализация — только сделки с вероятностью 50 %.
"""

from dataclasses import dataclass, field
from datetime import date

import polars as pl

DELIVERED = "Поставлено"
PROB_CONTRACTED = 1.0
PROB_HIGH = 0.5
PROB_POTENTIAL = 0.25

MONTHS = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]  # fmt: skip

OTHERS = "Другие"
NO_MVZ = "(МВЗ не указано)"
CUSTOMERS_ON_SLIDE = 7  # сколько крупнейших заказчиков показывать на слайде МВЗ, остальные — «Другие»


@dataclass
class Plan:
    """Плановые цифры одного МВЗ в рублях. Вводятся вручную раз в год."""

    revenue_forecast: float = 0.0  # выручка, «БП 9+3»
    margin_forecast: float = 0.0  # МП, «БП 9+3»
    revenue_budget: float = 0.0  # выручка, «БП'26»
    margin_budget: float = 0.0  # МП, «БП'26»


@dataclass
class ReportParams:
    report_date: date
    plans: dict[str, Plan] = field(default_factory=dict)  # МВЗ -> план
    forecast_label: str = ""  # заголовок колонки, например «БП 9+3»; пусто — по отчётному месяцу
    # Тексты титульного слайда сверху вниз: заголовок, подзаголовок, подпись внизу.
    # В шаблоне их нет: пользователь вводит их в приложении, они хранятся только у него.
    title_texts: tuple[str, str, str] = ("", "", "")

    @property
    def year(self) -> int:
        return self.report_date.year

    @property
    def month(self) -> int:
        return self.report_date.month

    def label_forecast(self) -> str:
        if self.forecast_label.strip():
            return self.forecast_label.strip()
        done = self.month - 1
        return f"БП {done}+{12 - done}"

    def label_budget(self) -> str:
        return f"БП'{self.year % 100:02d}"

    def plan_sum(self, mvz_list: list[str], attr: str) -> float:
        return sum(getattr(self.plans.get(m, Plan()), attr) for m in mvz_list)


# ---------- строки отчёта ----------


@dataclass
class KpiRow:
    """Строка верхней таблицы: Признано … Потенциал."""

    label: str
    recognized: float
    contracted: float
    high: float
    forecast: float  # «БП 9+3»
    budget: float  # «БП'26»
    potential: float
    percent: bool = False  # строка МП %: значения — доли

    @property
    def total(self) -> float:
        return self.recognized + self.contracted + self.high

    def values(self) -> list[float]:
        return [
            self.recognized, self.contracted, self.high, self.total,
            self.forecast, self.budget, self.total - self.budget, self.potential,
        ]  # fmt: skip


@dataclass
class PercentRow(KpiRow):
    """МП %: каждая колонка — МП / выручка. «Итого» пересчитывается из сумм, а не складывается."""

    total_value: float = 0.0

    @property
    def total(self) -> float:
        return self.total_value


@dataclass
class MonthRow:
    """Строка таблицы по месяцам: факт, три месяца; «Итого» считается сам."""

    label: str
    values: list[float]  # 4 числа: факт отчётного месяца, месяц, +1, +2
    level: str = "item"  # "item" — обычная, "total" — оранжевая, "margin" — чёрная

    @property
    def total(self) -> float:
        return sum(self.values)


@dataclass
class SummarySlide:
    title: str
    kpi: list[KpiRow]
    funnel: list[MonthRow]  # последняя строка — итог воронки
    plan: list[MonthRow]


@dataclass
class DetailSlide:
    title: str
    rows: list[list]  # месяц контр., заказчик, описание, выручка (руб.), статус, ответственный


DETAIL_HEADERS = ["Месяц контр.", "Заказчик", "Описание", "Выручка, млн. руб.", "Статус контрактации", "Ответственный"]
DETAIL_REVENUE = 3  # номер колонки с выручкой


@dataclass
class Report:
    params: ReportParams
    month_headers: list[str]  # «Октябрь Факт», «Октябрь», «Ноябрь», «Декабрь»
    summaries: list[SummarySlide]
    details: list[DetailSlide]


# ---------- расчёт ----------


def report_months(params: ReportParams) -> list[tuple[int, int]]:
    """(год, месяц) для отчётного месяца и двух следующих; декабрь переходит в январь."""
    result = []
    year, month = params.year, params.month
    for _ in range(3):
        result.append((year, month))
        month += 1
        if month > 12:
            year, month = year + 1, 1
    return result


def month_headers(params: ReportParams) -> list[str]:
    names = [MONTHS[m - 1] for _, m in report_months(params)]
    return [f"{names[0]} Факт", *names]


def mvz_list(df: pl.DataFrame) -> list[str]:
    """МВЗ, которые есть в строках, по алфавиту. Пустое МВЗ — в конце под именем NO_MVZ."""
    values = df.get_column("mvz").unique().drop_nulls().sort().to_list()
    if df.get_column("mvz").null_count():
        values.append(NO_MVZ)
    return values


def for_mvz(df: pl.DataFrame, mvz: str) -> pl.DataFrame:
    if mvz == NO_MVZ:
        return df.filter(pl.col("mvz").is_null())
    return df.filter(pl.col("mvz") == mvz)


def _is(prob: float) -> pl.Expr:
    """Вероятность равна prob (с допуском на дробные погрешности Excel)."""
    return (pl.col("probability") - prob).abs() < 1e-9


def _status() -> pl.Expr:
    return pl.col("order_status").fill_null("").str.strip_chars()


def _sum(df: pl.DataFrame, condition: pl.Expr, measure: str) -> float:
    return float(df.filter(condition).get_column(measure).sum() or 0.0)


def kpi_sums(df: pl.DataFrame, params: ReportParams, measure: str) -> dict[str, float]:
    """Признано / Законтракт. / Высокая вероятность / Потенциал за год отчёта."""
    year = df.filter(pl.col("year") == params.year)
    return {
        "recognized": _sum(year, _status() == DELIVERED, measure),
        "contracted": _sum(year, _is(PROB_CONTRACTED) & (_status() != DELIVERED), measure),
        "high": _sum(year, _is(PROB_HIGH), measure),
        "potential": _sum(year, _is(PROB_POTENTIAL), measure),
    }


def _ratio(margin: float, revenue: float) -> float:
    return margin / revenue if revenue else 0.0


def kpi_rows(df: pl.DataFrame, params: ReportParams, plan_mvz: list[str]) -> list[KpiRow]:
    """Верхняя таблица: выручка, МП и МП %. План — сумма планов перечисленных МВЗ."""
    revenue = KpiRow(
        "Выручка",
        **kpi_sums(df, params, "revenue"),
        forecast=params.plan_sum(plan_mvz, "revenue_forecast"),
        budget=params.plan_sum(plan_mvz, "revenue_budget"),
    )
    margin = KpiRow(
        "Маржинальная прибыль",
        **kpi_sums(df, params, "margin_rub"),
        forecast=params.plan_sum(plan_mvz, "margin_forecast"),
        budget=params.plan_sum(plan_mvz, "margin_budget"),
    )
    percent = PercentRow(
        "МП, %",
        recognized=_ratio(margin.recognized, revenue.recognized),
        contracted=_ratio(margin.contracted, revenue.contracted),
        high=_ratio(margin.high, revenue.high),
        forecast=_ratio(margin.forecast, revenue.forecast),
        budget=_ratio(margin.budget, revenue.budget),
        potential=_ratio(margin.potential, revenue.potential),
        percent=True,
        total_value=_ratio(margin.total, revenue.total),
    )
    return [revenue, margin, percent]


def month_values(df: pl.DataFrame, params: ReportParams, month_col: str, measure: str) -> list[float]:
    """Факт отчётного месяца (100 %) и три месяца с вероятностью 50 %.

    month_col — «contract_month» (воронка) или «month» (план по выручке).
    Для месяца выручки учитывается и год: декабрь → январь следующего года.
    """
    months = report_months(params)

    def in_month(year: int, month: int) -> pl.Expr:
        cond = pl.col(month_col) == month
        if month_col == "month":
            return cond & (pl.col("year") == year)
        # В месяце контрактации нет года. Чтобы не взять прошлогодние сделки,
        # берём только те, у которых выручка в году отчёта или позже.
        return cond & (pl.col("year") >= params.year)

    first_year, first_month = months[0]
    values = [_sum(df, in_month(first_year, first_month) & _is(PROB_CONTRACTED), measure)]
    for year, month in months:
        values.append(_sum(df, in_month(year, month) & _is(PROB_HIGH), measure))
    return values


def customer_rows(df: pl.DataFrame, params: ReportParams, limit: int) -> list[MonthRow]:
    """Крупнейшие заказчики воронки и строка «Другие»; заказчики с одними нулями не показываются."""
    rows = []
    for customer in df.get_column("customer").unique().to_list():
        part = df.filter(pl.col("customer").eq_missing(customer))
        values = month_values(part, params, "contract_month", "revenue")
        if any(abs(v) >= 0.5 for v in values):
            rows.append(MonthRow(customer if customer is not None else "(пусто)", values))
    rows.sort(key=lambda r: (-abs(r.total), r.label))
    if len(rows) <= limit:
        return rows
    top, rest = rows[:limit], rows[limit:]
    others = [sum(r.values[i] for r in rest) for i in range(4)]
    return [*top, MonthRow(OTHERS, others)]


def company_slide(df: pl.DataFrame, params: ReportParams, mvz: list[str]) -> SummarySlide:
    """Общий слайд: все отфильтрованные строки, разбивка по МВЗ."""
    funnel = [MonthRow(m, month_values(for_mvz(df, m), params, "contract_month", "revenue")) for m in mvz]
    funnel.append(MonthRow("Всего воронка продаж", month_values(df, params, "contract_month", "revenue"), "total"))
    plan = [MonthRow(m, month_values(for_mvz(df, m), params, "month", "revenue")) for m in mvz]
    plan.append(MonthRow("Всего выручка", month_values(df, params, "month", "revenue"), "total"))
    plan.append(MonthRow("Всего марж. прибыль", month_values(df, params, "month", "margin_rub"), "margin"))
    return SummarySlide(f"Исполнение {params.label_budget()}  БЕ АСУТП", kpi_rows(df, params, mvz), funnel, plan)


def mvz_slide(df: pl.DataFrame, params: ReportParams, mvz: str, customers: int) -> SummarySlide:
    """Слайд одного МВЗ: разбивка воронки по заказчикам."""
    part = for_mvz(df, mvz)
    funnel = customer_rows(part, params, customers)
    funnel.append(MonthRow("Всего контрактация", month_values(part, params, "contract_month", "revenue"), "total"))
    plan = [
        MonthRow("Всего выручка", month_values(part, params, "month", "revenue"), "total"),
        MonthRow("Всего марж. прибыль", month_values(part, params, "month", "margin_rub"), "margin"),
    ]
    title = f'Исполнение {params.label_budget()} по МВЗ "{mvz}"'
    return SummarySlide(title, kpi_rows(part, params, [mvz]), funnel, plan)


def detail_slide(df: pl.DataFrame, mvz: str) -> DetailSlide:
    """Сделки МВЗ с вероятностью 50 %: по месяцу контрактации, внутри — по убыванию выручки."""
    data = (
        for_mvz(df, mvz)
        .filter(_is(PROB_HIGH))
        .sort(["contract_month", "revenue"], descending=[False, True], nulls_last=True)
    )
    rows = []
    for r in data.iter_rows(named=True):
        rows.append(
            [
                r["contract_month"] if r["contract_month"] is not None else "",
                r["customer"] or "",
                r["description"] or "",
                r["revenue"] or 0.0,
                r["contract_status_short"] or "",
                r["manager"] or "",
            ]
        )
    return DetailSlide(f'Детализация контрактации по МВЗ "{mvz}"', rows)


def build_report(df: pl.DataFrame, params: ReportParams, customers: int = CUSTOMERS_ON_SLIDE) -> Report:
    """Все цифры презентации. df — строки листа после фильтров на сайте."""
    mvz = mvz_list(df)
    summaries = [company_slide(df, params, mvz)]
    summaries += [mvz_slide(df, params, m, customers) for m in mvz]
    details = [detail_slide(df, m) for m in mvz]
    return Report(params, month_headers(params), summaries, details)
