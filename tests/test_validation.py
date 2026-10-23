from sheetsync.cleaning import CleaningStats
from sheetsync.config import RevenueCheck, ValidationRules
from sheetsync.table import Table
from sheetsync.validation import validate

HEADER = ["Metric", "Code", "Company", "2026-01", "2026-02"]


def forecast(rows):
    return Table(header=list(HEADER), rows=rows)


def good_rows():
    return [
        ["Volume", "VOL-A", "A", 10.0, 20.0],
        ["Price", "PRI-A", "A", 500.0, 600.0],
        ["Revenue", "REV-A", "A", 5.0, 12.0],
    ]


def run(table, rules, **kwargs):
    stats = kwargs.pop("stats", CleaningStats(numeric_cells=table.n_rows * 2))
    return validate(table, rules, numeric_idx={3, 4}, stats=stats, **kwargs)


def test_clean_data_passes():
    rules = ValidationRules(required_columns=("Metric", "Code"), key_columns=("Code",), min_rows=3)
    report = run(forecast(good_rows()), rules)
    assert report.ok and not report.warnings


def test_missing_column_with_case_hint():
    table = Table(header=["metric", "Code"], rows=[["Volume", "X"]])
    report = run(table, ValidationRules(required_columns=("Metric",)))
    assert not report.ok
    assert "Есть похожий: 'metric'" in report.errors[0].message


def test_key_column_is_required_even_if_not_listed():
    report = run(forecast(good_rows()), ValidationRules(key_columns=("order_id",)))
    assert not report.ok
    assert "order_id" in report.errors[0].message


def test_duplicate_and_empty_keys_are_errors():
    rows = good_rows() + [["Volume", "VOL-A", "A", 1.0, 1.0], ["Volume", "", "A", 1.0, 1.0]]
    table = Table(header=list(HEADER), rows=rows, row_numbers=[5, 6, 7, 8, 9])
    report = run(table, ValidationRules(key_columns=("Code",)))
    messages = [str(e) for e in report.errors]
    assert any("VOL-A повторяется в строках 5, 8" in m for m in messages)
    assert any(m.startswith("строка 9") for m in messages)


def test_row_drop_blocks_unless_forced():
    rules = ValidationRules(max_row_drop_pct=20)
    assert not run(forecast(good_rows()), rules, previous_rows=10).ok
    forced = run(forecast(good_rows()), rules, previous_rows=10, force=True)
    assert forced.ok and "принудительно" in forced.warnings[0].message


def test_too_many_bad_values_is_error():
    stats = CleaningStats(numeric_cells=100, bad_values=10)
    report = run(forecast(good_rows()), ValidationRules(max_bad_values_pct=5), stats=stats)
    assert not report.ok and "10.0%" in report.errors[0].message


def test_negative_values_are_warnings():
    rows = good_rows()
    rows[0][3] = -1.0
    report = run(forecast(rows), ValidationRules(all_numeric_non_negative=True))
    assert report.ok
    assert report.warnings[0].column == "2026-01"


REVENUE = RevenueCheck(metric_column="Metric", group_by=("Company",), volume="Volume", price="Price",
                       revenue="Revenue", divisor=1000, tolerance_pct=1)


def test_revenue_check_flags_mismatch_only():
    rows = good_rows()
    rows[2][4] = 14.4  # должно быть 20 × 600 / 1000 = 12
    report = run(forecast(rows), ValidationRules(revenue_check=REVENUE))
    assert report.ok
    assert len(report.warnings) == 1
    assert report.warnings[0].column == "2026-02" and "+20%" in report.warnings[0].message


def test_revenue_check_skips_empty_cells():
    rows = good_rows()
    rows[0][3] = None
    report = run(forecast(rows), ValidationRules(revenue_check=REVENUE))
    assert not report.warnings


def test_summary_is_truncated():
    rows = [["Volume", f"K{i}", "A", -1.0, -1.0] for i in range(10)]
    report = run(forecast(rows), ValidationRules(all_numeric_non_negative=True))
    assert "… и ещё 15" in report.summary(limit=5)
