"""Описание листа «Прогноз выручки» в одном месте.

Если в исходной книге поменяются заголовки или колонки, править нужно только этот файл.
Колонки определяются по букве (позиции), а не по названию: «МП, руб.» встречается дважды.
"""

from dataclasses import dataclass

SHEET_NAME = "Прогноз выручки"
HEADER_ROW = 2  # строка с заголовками (нумерация как в Excel)
FIRST_DATA_ROW = 3
THRESHOLD_CELL = "Z1"  # порог вероятности, от которого зависит колонка Z

# Типы значений в колонке
INT = "int"  # целое число
NUMBER = "number"  # дробное число (суммы, доли, вероятности)
TEXT = "text"  # текст; числа приводятся к тексту
DATE = "date"  # дата; текст ДД.ММ.ГГГГ приводится к дате

# Роли колонок (пригодятся интерфейсу и сводным)
DIMENSION = "dimension"  # по ней группируют и фильтруют
MEASURE = "measure"  # её суммируют
RATIO = "ratio"  # долю нельзя суммировать, только пересчитывать
KEY = "key"  # идентификатор
ATTRIBUTE = "attribute"  # описание, комментарий
FLAG = "flag"  # 0/1


@dataclass(frozen=True)
class Column:
    letter: str  # буква колонки в Excel
    code: str  # уникальное латинское имя внутри кода
    header: str  # заголовок как в файле (сравнивается после нормализации пробелов)
    label: str  # название для интерфейса и выгрузок
    dtype: str
    role: str
    percent: bool = False  # доля 0..1, на экране и в выгрузке показывать как проценты
    # strict=False: если в колонке много значений не того типа, это предупреждение, а не ошибка.
    strict: bool = True


COLUMNS: tuple[Column, ...] = (
    Column("B", "year", "ГОД", "Год", INT, DIMENSION),
    Column("C", "month", "Месяц выручки", "Месяц выручки", INT, DIMENSION),
    Column("D", "order_no", "№ заказа", "№ заказа", TEXT, KEY),
    Column("E", "asutp_no", "№ АСУТП", "№ АСУТП", TEXT, DIMENSION),
    Column("F", "mvz", "МВЗ", "МВЗ", TEXT, DIMENSION),
    Column("G", "order_type", "Тип закза", "Тип заказа", TEXT, DIMENSION),
    Column("H", "description", "Описание", "Описание", TEXT, ATTRIBUTE),
    Column("I", "customer", "Заказчик", "Заказчик", TEXT, DIMENSION),
    Column("J", "supplier", "Поставщик", "Поставщик", TEXT, DIMENSION),
    Column("K", "order_status", "Статус заказа", "Статус заказа", TEXT, DIMENSION),
    Column("L", "manager", "Ответственный", "Ответственный", TEXT, DIMENSION),
    Column("M", "end_customer", "Конечный заказчик", "Конечный заказчик", TEXT, DIMENSION),
    Column("N", "probability", "Вероятность заказа", "Вероятность заказа", NUMBER, DIMENSION, percent=True),
    Column("O", "revenue_date", "Дата выручки", "Дата выручки", DATE, DIMENSION),
    Column("P", "revenue", "Выручка, руб. (без НДС)", "Выручка, руб. (без НДС)", NUMBER, MEASURE),
    Column("Q", "cost", "Себес., руб. (без НДС)", "Себес., руб. (без НДС)", NUMBER, MEASURE),
    Column("R", "margin_rub", "МП, руб.", "МП, руб.", NUMBER, MEASURE),
    Column("S", "margin_pct", "Маржа (Gross profit), %", "Маржа (Gross profit), %", NUMBER, RATIO, percent=True),
    Column("T", "contract_month", "Месяц контрактации", "Месяц контрактации", INT, DIMENSION),
    Column("U", "revenue_date_check", "Контроль даты выручки", "Контроль даты выручки", TEXT, DIMENSION),
    # V — дата контрактации (ответ пользователя на открытый вопрос 1 из SPEC.md).
    # Excel может хранить дату числом (серийный номер), loader переводит его в дату.
    # strict=False: в обезличенном образце в V записаны суммы, файл всё равно должен читаться.
    Column("V", "contract_date", "Дата коонтр.", "Дата контрактации", DATE, DIMENSION, strict=False),
    Column("W", "cost2", "Себестоимость, руб.", "Себестоимость, руб. (W)", NUMBER, MEASURE),
    Column("X", "margin_rub2", "МП, руб.", "МП, руб. (X)", NUMBER, MEASURE),
    Column("Y", "project_comment", "Комментарии по проекту", "Комментарии по проекту", TEXT, ATTRIBUTE),
    Column("Z", "passes_threshold", "тех колонка", "Проходит порог вероятности", INT, FLAG),
    Column("AA", "contract_status_short", "Статус контрактации кратко", "Статус контрактации кратко", TEXT, ATTRIBUTE),
)

# Служебная колонка с номером строки в исходном файле: нужна для сообщений и проверки.
ROW_CODE = "excel_row"
ROW_LABEL = "Строка в файле"

LAST_LETTER = COLUMNS[-1].letter


def col_index(letter: str) -> int:
    """Буква колонки -> индекс с нуля: A -> 0, Z -> 25, AA -> 26."""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def col_letter(index: int) -> str:
    """Индекс с нуля -> буква колонки: 0 -> A, 26 -> AA."""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


BY_CODE = {c.code: c for c in COLUMNS}
LABELS = {ROW_CODE: ROW_LABEL} | {c.code: c.label for c in COLUMNS}
