"""
Очистка значений перед записью в Excel.

Правила:
  * числовые столбцы — число или пусто. Текст, похожий на число, чиним
    ('1 234,5' → 1234.5). Маркеры «нет данных» ('n/a', '-', …) → пусто.
    Мусор ('уточнить', '#REF!') → пусто + предупреждение со строкой и столбцом;
  * столбцы дат — datetime или пусто. Понимаем ГГГГ-ММ-ДД, ДД.ММ.ГГГГ,
    ДД/ММ/ГГГГ и серийный номер даты; невозможные даты (31.02) — мусор;
  * текстовые столбцы — обрезаем пробелы по краям и схлопываем двойные
    пробелы внутри: 'СтальГрад ' и 'СтальГрад' — одна компания, иначе
    СУММЕСЛИ в Excel посчитает их отдельно.

Сами «грязные» исходные данные в Google Sheets не трогаем — чистится только
то, что уходит в Excel.
"""

import datetime
import math
from dataclasses import dataclass

from .table import Table

MISSING_MARKERS = {"", "-", "—", "–", "n/a", "na", "#n/a", "н/д", "null", "none", "нет данных"}
FORMULA_ERRORS = {"#ref!", "#value!", "#div/0!", "#name?", "#num!", "#null!", "#error!"}
DATE_FORMATS = ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M:%S")
# Excel и Google Sheets считают даты днями от 30.12.1899
_SERIAL_EPOCH = datetime.datetime(1899, 12, 30)

# Статусы разбора одного значения
OK, FIXED, MISSING, BAD = "ok", "fixed", "missing", "bad"


@dataclass
class Issue:
    level: str  # "error" — блокирует запись, "warning" — нет
    message: str
    row: int | None = None  # номер строки в листе Google Sheets
    column: str | None = None

    def __str__(self):
        where = []
        if self.row is not None:
            where.append(f"строка {self.row}")
        if self.column:
            where.append(f"столбец '{self.column}'")
        return (", ".join(where) + ": " if where else "") + self.message


@dataclass
class CleaningStats:
    numeric_cells: int = 0
    numbers_from_text: int = 0
    missing_values: int = 0
    bad_values: int = 0
    dates_reformatted: int = 0
    texts_trimmed: int = 0

    def notes(self) -> list[str]:
        """Что исправили сами — для отчёта, чтобы автоисправления не были «молчаливыми»."""
        parts = [
            (self.numbers_from_text, "чисел, записанных текстом"),
            (self.dates_reformatted, "дат в нестандартном формате"),
            (self.texts_trimmed, "значений с лишними пробелами"),
            (self.missing_values, "пустых значений и маркеров «нет данных»"),
        ]
        return [f"{count} {label}" for count, label in parts if count]


def parse_number(value):
    """Возвращает (число или None, статус)."""
    if value is None:
        return None, MISSING
    if isinstance(value, bool):  # True — тоже int в Python, но в таблице это не число
        return None, BAD
    if isinstance(value, (int, float)):
        return (value, OK) if math.isfinite(value) else (None, BAD)

    text = str(value).strip().lower()
    if text in MISSING_MARKERS:
        return None, MISSING

    # split() без аргументов режет по любым пробелам, включая неразрывные
    # (U+00A0) — ими русский Excel разделяет разряды: '1 234,5'
    s = "".join(text.split())
    negative = s.startswith("(") and s.endswith(")")  # бухгалтерская запись (123) = -123
    if negative:
        s = s[1:-1]
    percent = s.endswith("%")
    if percent:
        s = s[:-1]
    if "," in s and "." in s:
        # разделитель дробной части — тот, что правее: 1.234,5 и 1,234.5
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        # одна запятая — дробная часть (русская локаль): '1,234' читаем как 1.234
        s = s.replace(",", ".")

    try:
        number = float(s)
    except ValueError:
        return None, BAD
    if not math.isfinite(number):  # float() понимает 'inf' и 'nan' — нам такое не нужно
        return None, BAD
    if percent:
        number /= 100
    if negative:
        number = -number
    return number, FIXED


def parse_date(value):
    """Возвращает (datetime или None, статус)."""
    if value is None:
        return None, MISSING
    if isinstance(value, datetime.datetime):
        return value, OK
    if isinstance(value, datetime.date):
        return datetime.datetime.combine(value, datetime.time()), OK
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= value <= 2_958_465:  # 31.12.9999 — максимум в Excel
            return _SERIAL_EPOCH + datetime.timedelta(days=value), FIXED
        return None, BAD

    text = " ".join(str(value).split())
    if text.lower() in MISSING_MARKERS:
        return None, MISSING
    for i, fmt in enumerate(DATE_FORMATS):
        try:
            parsed = datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
        return parsed, OK if i == 0 else FIXED
    return None, BAD


def clean_text(value):
    """Возвращает (значение, статус). Числа в текстовых столбцах не трогаем."""
    if not isinstance(value, str):
        return ("", MISSING) if value is None else (value, OK)
    cleaned = " ".join(value.split())
    if cleaned.lower() in FORMULA_ERRORS:
        return "", BAD
    if not cleaned:  # пустая строка или одни пробелы
        return "", MISSING
    return cleaned, OK if cleaned == value else FIXED


def normalize_header(name) -> str:
    return " ".join(str(name).split()) if name is not None else ""


def clean_table(table: Table, numeric_idx: set[int], date_idx: set[int]):
    """Возвращает (новая Table, список Issue, CleaningStats)."""
    stats = CleaningStats()
    issues = []
    header = [normalize_header(h) for h in table.header]
    rows = []

    for row, row_number in zip(table.rows, table.row_numbers):
        new_row = []
        for col, value in enumerate(row):
            if col in numeric_idx:
                stats.numeric_cells += 1
                new_value, status = parse_number(value)
            elif col in date_idx:
                new_value, status = parse_date(value)
            else:
                new_value, status = clean_text(value)

            if status == BAD:
                stats.bad_values += 1
                kind = "дата" if col in date_idx else "число" if col in numeric_idx else "значение"
                issues.append(Issue(
                    "warning", f"некорректное {kind} {value!r} — ячейка оставлена пустой",
                    row=row_number, column=header[col],
                ))
                new_value = None
            elif status == MISSING and value not in ("", None):
                stats.missing_values += 1
            elif status == FIXED:
                if col in numeric_idx:
                    stats.numbers_from_text += 1
                elif col in date_idx:
                    stats.dates_reformatted += 1
                else:
                    stats.texts_trimmed += 1
            new_row.append(new_value)
        rows.append(new_row)

    return Table(header=header, rows=rows, row_numbers=list(table.row_numbers)), issues, stats
