from pathlib import Path

import pytest

from sheetsync.config import extract_spreadsheet_id, load_config
from sheetsync.errors import ConfigError, SourceError
from sheetsync.source import to_table

EXAMPLE = Path(__file__).resolve().parent.parent / "config.example.yaml"


def write_config(tmp_path, text):
    path = tmp_path / "config.yaml"
    path.write_text(text, encoding="utf-8")
    return path


def filled_example():
    return EXAMPLE.read_text(encoding="utf-8").replace("PASTE_SPREADSHEET_ID_HERE", "abc123")


def test_example_config_loads(tmp_path):
    config = load_config(write_config(tmp_path, filled_example()))
    assert set(config.jobs) == {"forecast", "orders"}
    assert config.jobs["orders"].source_range is None
    assert config.jobs["forecast"].numeric_pattern.match("2026-03")
    assert config.state_db == tmp_path / "state" / "sync_state.sqlite"


def test_placeholder_id_is_rejected(tmp_path):
    with pytest.raises(ConfigError, match="spreadsheet_id"):
        load_config(write_config(tmp_path, EXAMPLE.read_text(encoding="utf-8")))


def test_bad_range_is_rejected(tmp_path):
    text = filled_example().replace("range: B4:BP3000", "range: B4-BP3000")
    with pytest.raises(ConfigError, match="диапазон"):
        load_config(write_config(tmp_path, text))


def test_wrong_type_is_rejected(tmp_path):
    text = filled_example().replace("min_rows: 1000", "min_rows: много")
    with pytest.raises(ConfigError, match="min_rows"):
        load_config(write_config(tmp_path, text))


def test_missing_file(tmp_path):
    with pytest.raises(ConfigError, match="config.example.yaml"):
        load_config(tmp_path / "nope.yaml")


def test_unknown_job(tmp_path):
    config = load_config(write_config(tmp_path, filled_example()))
    with pytest.raises(ConfigError, match="Доступные: forecast, orders"):
        config.job("sales")


def test_spreadsheet_url_is_accepted():
    url = "https://docs.google.com/spreadsheets/d/1YN4Xyz_-abc/edit?gid=0#gid=0"
    assert extract_spreadsheet_id(url) == "1YN4Xyz_-abc"


def test_to_table_pads_ragged_rows_and_drops_empty_ones():
    raw = [["Code", "a", "b"], ["X", 1], [], ["", "  "], ["Y", 2, 3]]
    table = to_table(raw, first_row=4, where="тест")
    assert table.rows == [["X", 1, ""], ["Y", 2, 3]]
    assert table.row_numbers == [5, 8]


def test_to_table_skips_leading_empty_rows():
    table = to_table([[], ["Code"], ["X"]], first_row=1, where="тест")
    assert table.header == ["Code"] and table.row_numbers == [3]


def test_to_table_empty_source():
    with pytest.raises(SourceError, match="пуст"):
        to_table([[], [""]], first_row=1, where="тест")
