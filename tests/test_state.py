import datetime

import pytest

from sheetsync.errors import StateError
from sheetsync.state import StateStore
from sheetsync.table import Table


def test_snapshot_roundtrip(tmp_path):
    table = Table(header=["id", "date"], rows=[["A", datetime.datetime(2026, 3, 15)]], row_numbers=[7])
    with StateStore(tmp_path / "s.sqlite") as state:
        run_id = state.record_run("orders", "success", datetime.datetime.now(), table=table)
        state.save_snapshot("orders", run_id, table)
        loaded = state.load_snapshot("orders")
    assert loaded.rows == [["A", "2026-03-15T00:00:00"]]
    assert loaded.row_numbers == [7]


def test_last_success_ignores_failed_runs(tmp_path):
    table = Table(header=["id"], rows=[["A"]])
    with StateStore(tmp_path / "s.sqlite") as state:
        state.record_run("orders", "success", datetime.datetime.now(), table=table)
        state.record_run("orders", "failed", datetime.datetime.now(), message="boom")
        assert state.last_success("orders")["rows"] == 1
        assert [r["status"] for r in state.history("orders")] == ["failed", "success"]
        assert state.last_success("forecast") is None


def test_state_survives_reopen(tmp_path):
    with StateStore(tmp_path / "s.sqlite") as state:
        state.record_run("orders", "success", datetime.datetime.now())
    with StateStore(tmp_path / "s.sqlite") as state:
        assert len(state.history()) == 1


def test_corrupted_db_gives_clear_error(tmp_path):
    db = tmp_path / "s.sqlite"
    db.write_bytes(b"this is not a database at all" * 100)
    with pytest.raises(StateError, match="повреждён"):
        StateStore(db)
