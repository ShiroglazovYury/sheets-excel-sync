import datetime

import pytest

from sheetsync.cleaning import BAD, FIXED, MISSING, OK, clean_table, clean_text, parse_date, parse_number
from sheetsync.table import Table


@pytest.mark.parametrize("raw, expected, status", [
    (12.5, 12.5, OK),
    (0, 0, OK),
    ("1 234,56", 1234.56, FIXED),
    ("1\xa0234,56", 1234.56, FIXED),     # неразрывный пробел из русского Excel
    ("1.234,5", 1234.5, FIXED),
    ("1,234.5", 1234.5, FIXED),
    ("1,234", 1.234, FIXED),             # одна запятая — дробная часть
    ("(150)", -150, FIXED),
    ("12%", 0.12, FIXED),
    ("  42 ", 42, FIXED),
])
def test_parse_number_ok(raw, expected, status):
    value, got_status = parse_number(raw)
    assert value == pytest.approx(expected)
    assert got_status == status


@pytest.mark.parametrize("raw", ["", "n/a", "N/A", "-", "—", "#N/A", "нет данных", None])
def test_parse_number_missing(raw):
    assert parse_number(raw) == (None, MISSING)


@pytest.mark.parametrize("raw", ["уточнить", "#REF!", "1.234.567", "inf", "nan", True, float("nan")])
def test_parse_number_bad(raw):
    assert parse_number(raw) == (None, BAD)


@pytest.mark.parametrize("raw, expected, status", [
    ("2026-03-15", datetime.datetime(2026, 3, 15), OK),
    ("15.03.2026", datetime.datetime(2026, 3, 15), FIXED),
    ("15/03/2026", datetime.datetime(2026, 3, 15), FIXED),
    (46096, datetime.datetime(2026, 3, 15), FIXED),  # серийный номер даты
    (datetime.date(2026, 3, 15), datetime.datetime(2026, 3, 15), OK),
])
def test_parse_date_ok(raw, expected, status):
    assert parse_date(raw) == (expected, status)


@pytest.mark.parametrize("raw", ["31.02.2026", "2026-13-05", "вчера", "00.00.0000", -5])
def test_parse_date_bad(raw):
    assert parse_date(raw) == (None, BAD)


@pytest.mark.parametrize("raw, expected, status", [
    ("СтальГрад", "СтальГрад", OK),
    ("  СтальГрад ", "СтальГрад", FIXED),
    ("Горячекатаный   прокат", "Горячекатаный прокат", FIXED),
    ("   ", "", MISSING),
    ("#REF!", "", BAD),
    ("0274051582", "0274051582", OK),  # ИНН с ведущим нулём остаётся строкой
])
def test_clean_text(raw, expected, status):
    assert clean_text(raw) == (expected, status)


def test_clean_table_reports_bad_cells_with_sheet_row_numbers():
    table = Table(
        header=["name ", "volume", "date"],
        rows=[["A ", "1 000,5", "15.03.2026"], ["B", "уточнить", "31.02.2026"]],
        row_numbers=[5, 9],
    )
    cleaned, issues, stats = clean_table(table, numeric_idx={1}, date_idx={2})

    assert cleaned.header == ["name", "volume", "date"]
    assert cleaned.rows[0] == ["A", 1000.5, datetime.datetime(2026, 3, 15)]
    assert cleaned.rows[1] == ["B", None, None]
    assert [(i.row, i.column) for i in issues] == [(9, "volume"), (9, "date")]
    assert stats.bad_values == 2 and stats.numbers_from_text == 1 and stats.dates_reformatted == 1
    assert stats.texts_trimmed == 1
