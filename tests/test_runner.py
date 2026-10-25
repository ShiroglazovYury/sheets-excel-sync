"""Сценарий синка целиком — с поддельными Google и Excel, без сети."""

from pathlib import Path

import pytest

from sheetsync.config import load_config
from sheetsync.errors import ValidationFailed
from sheetsync.runner import SyncRunner
from sheetsync.state import StateStore
from sheetsync.table import Table

EXAMPLE = Path(__file__).resolve().parent.parent / "config.example.yaml"
HEADER = ["order_id", "order_date", "company", "customer_inn", "region", "product", "volume_t", "price_usd_t", "status"]


def orders(n, price="512,5"):
    rows = [[f"ORD-{i:06d}", "2026-03-15", " СтальГрад ", "0274051582", "Россия", "Арматура", 10, price, "shipped"]
            for i in range(n)]
    return Table(header=list(HEADER), rows=rows)


class FakeSource:
    def __init__(self, table):
        self.table = table

    def read_table(self, tab, cell_range=None):
        return self.table


class FakeTarget:
    def __init__(self):
        self.writes, self.logs = [], []

    def write_table(self, sheet_name, top_left, table, **kwargs):
        self.writes.append((sheet_name, table, kwargs))

    def append_log(self, values):
        self.logs.append(values)


@pytest.fixture
def setup(tmp_path):
    text = EXAMPLE.read_text(encoding="utf-8").replace("PASTE_SPREADSHEET_ID_HERE", "abc")
    (tmp_path / "config.yaml").write_text(text, encoding="utf-8")
    config = load_config(tmp_path / "config.yaml")
    source, target = FakeSource(orders(1500)), FakeTarget()
    state = StateStore(config.state_db)
    yield SyncRunner(config, source=source, target=target, state=state), source, target, state
    state.close()


def test_first_run_writes_cleaned_data_and_backup(setup):
    runner, _, target, state = setup
    result = runner.run("orders")

    assert result.status == "success"
    (sheet, table, kwargs), = target.writes
    assert sheet == "Orders"
    assert table.rows[0][2] == "СтальГрад" and table.rows[0][7] == 512.5
    assert kwargs["text_cols"] == {0, 2, 3, 4, 5, 8}  # ИНН среди текстовых — ведущий ноль не пропадёт
    assert result.backup_path.exists()
    assert state.load_snapshot("orders").n_rows == 1500


def test_unchanged_source_is_not_rewritten(setup):
    runner, _, target, _ = setup
    runner.run("orders")
    second = runner.run("orders")
    assert second.status == "unchanged"
    assert len(target.writes) == 1
    assert runner.run("orders", force=True).status == "success"


def test_unchanged_is_decided_like_the_message_says(setup):
    """Те же данные с «шумом» в 12-м знаке: окно пишет «не изменилось» — значит, и не пишем."""
    runner, source, target, _ = setup
    runner.run("orders")
    source.table = orders(1500, price=512.5 + 1e-12)
    result = runner.run("orders")
    assert not result.diff.has_changes
    assert result.status == "unchanged"
    assert len(target.writes) == 1


def test_changes_are_reported(setup):
    runner, source, _, _ = setup
    runner.run("orders")
    source.table = orders(1500, price="600")
    result = runner.run("orders")
    assert result.diff.changed_cells == 1500
    assert "512.50 → 600" in str(result.diff.top_changes[0])


def test_invalid_data_never_touches_excel(setup):
    runner, source, target, state = setup
    source.table = orders(1500)
    source.table.rows[5][0] = source.table.rows[4][0]  # дубль ключа
    with pytest.raises(ValidationFailed) as error:
        runner.run("orders")
    assert target.writes == []
    assert error.value.report_path.exists()
    assert state.history("orders")[0]["status"] == "failed"


def test_sharp_row_drop_needs_force(setup):
    runner, source, target, _ = setup
    runner.run("orders")
    source.table = orders(1200)  # −20% при допустимых −10%
    with pytest.raises(ValidationFailed, match="−20%"):
        runner.run("orders")
    assert runner.run("orders", force=True).status == "success"
    assert len(target.writes) == 2


def test_dry_run_saves_nothing(setup):
    runner, _, target, state = setup
    assert runner.run("orders", dry_run=True).status == "dry-run"
    assert target.writes == [] and state.load_snapshot("orders") is None
