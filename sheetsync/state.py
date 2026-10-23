"""
Состояние между запусками — SQLite-файл state/sync_state.sqlite.

Таблица runs     — журнал всех запусков: когда, какое задание, чем закончилось,
                   сколько строк, отпечаток данных, что изменилось.
Таблица snapshots — последний успешно записанный срез данных по каждому
                   заданию (gzip-JSON). С ним сравниваем следующий запуск.

Зачем это нужно:
  * не перезаписывать Excel, если в источнике ничего не поменялось
    (сравниваем отпечаток с прошлым успешным запуском);
  * показать, что именно изменилось с прошлого раза;
  * поймать подозрительное падение числа строк (кто-то удалил пол-листа).

SQLite выбран вместо JSON-файла: запись атомарная (не останется
полузаписанного файла, если Excel закроют посреди синка), а историю удобно
фильтровать запросами.
"""

import datetime
import gzip
import json
import sqlite3
from pathlib import Path

from .errors import StateError
from .table import Table

_SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    job           TEXT NOT NULL,
    started_at    TEXT NOT NULL,
    finished_at   TEXT NOT NULL,
    status        TEXT NOT NULL,   -- success | unchanged | dry-run | failed
    rows          INTEGER,
    cols          INTEGER,
    fingerprint   TEXT,
    added         INTEGER,
    removed       INTEGER,
    changed_cells INTEGER,
    warnings      INTEGER,
    message       TEXT
);
CREATE TABLE IF NOT EXISTS snapshots (
    job      TEXT PRIMARY KEY,
    run_id   INTEGER NOT NULL,
    saved_at TEXT NOT NULL,
    data     BLOB NOT NULL
);
"""


class StateStore:
    def __init__(self, db_path: Path):
        db_path = Path(db_path)
        db_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            self._conn = sqlite3.connect(db_path)
            self._conn.row_factory = sqlite3.Row
            self._conn.executescript(_SCHEMA)
        except sqlite3.DatabaseError as e:
            raise StateError(
                f"Файл состояния {db_path} повреждён ({e}). Удалите его — история "
                "синков начнётся заново, данные в Excel и Google не пострадают."
            ) from None

    def close(self):
        self._conn.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    def record_run(self, job: str, status: str, started_at: datetime.datetime, *, table: Table | None = None,
                   diff=None, warnings: int = 0, message: str = "") -> int:
        with self._conn:  # транзакция: либо запись целиком, либо ничего
            cur = self._conn.execute(
                """INSERT INTO runs (job, started_at, finished_at, status, rows, cols, fingerprint,
                                     added, removed, changed_cells, warnings, message)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    job, _ts(started_at), _ts(datetime.datetime.now()), status,
                    table.n_rows if table else None, table.n_cols if table else None,
                    table.fingerprint() if table else None,
                    diff.added if diff else None, diff.removed if diff else None,
                    diff.changed_cells if diff else None, warnings, message,
                ),
            )
        return cur.lastrowid

    def last_success(self, job: str) -> sqlite3.Row | None:
        return self._conn.execute(
            "SELECT * FROM runs WHERE job = ? AND status = 'success' ORDER BY id DESC LIMIT 1", (job,)
        ).fetchone()

    def history(self, job: str | None = None, limit: int = 20) -> list[sqlite3.Row]:
        if job:
            query, params = "SELECT * FROM runs WHERE job = ? ORDER BY id DESC LIMIT ?", (job, limit)
        else:
            query, params = "SELECT * FROM runs ORDER BY id DESC LIMIT ?", (limit,)
        return self._conn.execute(query, params).fetchall()

    def save_snapshot(self, job: str, run_id: int, table: Table):
        data = gzip.compress(json.dumps(table.to_jsonable(), ensure_ascii=False).encode("utf-8"))
        with self._conn:
            self._conn.execute(
                "INSERT OR REPLACE INTO snapshots (job, run_id, saved_at, data) VALUES (?, ?, ?, ?)",
                (job, run_id, _ts(datetime.datetime.now()), data),
            )

    def load_snapshot(self, job: str) -> Table | None:
        row = self._conn.execute("SELECT data FROM snapshots WHERE job = ?", (job,)).fetchone()
        if row is None:
            return None
        try:
            return Table.from_jsonable(json.loads(gzip.decompress(row["data"])))
        except (OSError, ValueError, KeyError):
            # битый снимок — не повод ронять синк: просто сравнивать будет не с чем
            return None


def _ts(moment: datetime.datetime) -> str:
    return moment.isoformat(sep=" ", timespec="seconds")
