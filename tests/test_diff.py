import datetime

from sheetsync.diff import diff_tables
from sheetsync.table import Table


def table(rows, header=("Code", "2026-01", "2026-02")):
    return Table(header=list(header), rows=[list(r) for r in rows])


def test_first_run():
    result = diff_tables(None, table([["A", 1, 2]]), ["Code"])
    assert result.first_run and result.has_changes


def test_no_changes():
    t = table([["A", 1.0, 2.0], ["B", 3.0, 4.0]])
    result = diff_tables(t, table([["A", 1.0, 2.0], ["B", 3.0, 4.0]]), ["Code"])
    assert not result.has_changes


def test_inserted_row_does_not_shift_comparison():
    """Ключевой случай: строку вставили в середину — изменилась только она."""
    old = table([["A", 1, 2], ["B", 3, 4]])
    new = table([["A", 1, 2], ["NEW", 9, 9], ["B", 3, 4]])
    result = diff_tables(old, new, ["Code"])
    assert (result.added, result.removed, result.changed_cells) == (1, 0, 0)


def test_changed_cells_sorted_by_size():
    old = table([["A", 100.0, 100.0], ["B", 100.0, 100.0]])
    new = table([["A", 101.0, 100.0], ["B", 100.0, 150.0]])
    result = diff_tables(old, new, ["Code"])
    assert result.changed_cells == 2
    assert result.top_changes[0].key == "B" and result.top_changes[0].column == "2026-02"
    assert "+50.0%" in str(result.top_changes[0])


def test_columns_added_and_removed():
    old = table([["A", 1, 2]])
    new = table([["A", 5, 3]], header=("Code", "2026-02", "2026-03"))  # 2026-02: 2 → 5
    result = diff_tables(old, new, ["Code"])
    assert result.columns_added == ["2026-03"] and result.columns_removed == ["2026-01"]
    assert result.changed_cells == 1  # сравниваются только общие столбцы


def test_dates_compare_with_snapshot_strings():
    old = table([["A", "2026-03-15T00:00:00", 1]], header=("Code", "date", "x"))
    new = table([["A", datetime.datetime(2026, 3, 15), 1]], header=("Code", "date", "x"))
    assert not diff_tables(old, new, ["Code"]).has_changes


def test_empty_string_equals_none():
    assert not diff_tables(table([["A", "", 1]]), table([["A", None, 1]]), ["Code"]).has_changes
