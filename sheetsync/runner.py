"""
Сценарий одного синка — собирает все модули в цепочку:

    Google Sheets → очистка → проверки → сравнение с прошлым разом
        → запись в Excel → бэкап → состояние и журнал

Каждый шаг может остановить цепочку понятной ошибкой (SyncError), и тогда
Excel остаётся нетронутым. Запуск в любом случае фиксируется в истории.
"""

import datetime
import time
from dataclasses import dataclass
from pathlib import Path

from .backup import rotate, save_backup
from .cleaning import clean_table, normalize_header
from .config import AppConfig, JobConfig
from .diff import DiffResult, diff_tables
from .errors import ConfigError, SyncError, ValidationFailed
from .excel_target import ExcelTarget
from .source import GoogleSheetSource
from .state import StateStore
from .table import Table
from .validation import ValidationReport, validate

SUCCESS, UNCHANGED, DRY_RUN, FAILED = "success", "unchanged", "dry-run", "failed"


@dataclass
class RunResult:
    job: JobConfig
    status: str
    table: Table
    report: ValidationReport
    diff: DiffResult
    seconds: float
    backup_path: Path | None = None
    report_path: Path | None = None

    def message(self) -> str:
        job, t = self.job, self.table
        head = {
            SUCCESS: f"«{job.title}»: записано {t.n_rows} строк × {t.n_cols} столбцов на лист '{job.target_sheet}'.",
            UNCHANGED: f"«{job.title}»: в Google Sheets ничего не изменилось — Excel не трогали.",
            DRY_RUN: f"«{job.title}»: проверка без записи (dry-run), строк: {t.n_rows}.",
        }[self.status]
        parts = [f"{head} [{self.seconds:.1f} с]", self.diff.summary()]
        if self.report.issues or self.report.notes:
            parts.append(self.report.summary())
        if self.report_path:
            parts.append(f"Полный отчёт: {self.report_path}")
        if self.backup_path:
            parts.append(f"Бэкап: {self.backup_path.name}")
        return "\n\n".join(parts)


def resolve_columns(header: list[str], job: JobConfig) -> tuple[set[int], set[int]]:
    """Номера числовых столбцов и столбцов дат — по именам и шаблону из конфига."""
    numeric = {
        i for i, name in enumerate(header)
        if name in job.numeric_columns or (job.numeric_pattern and job.numeric_pattern.match(name))
    }
    dates = {i for i, name in enumerate(header) if name in job.date_columns}
    if numeric & dates:
        both = ", ".join(header[i] for i in sorted(numeric & dates))
        raise ConfigError(f"Столбцы {both} указаны в задании '{job.name}' и как числа, и как даты.")
    return numeric, dates


class SyncRunner:
    def __init__(self, config: AppConfig, source=None, target=None, state=None):
        self.config = config
        self.source = source or GoogleSheetSource(config.spreadsheet_id, config.credentials_file)
        self.state = state or StateStore(config.state_db)
        self._target = target

    @property
    def target(self) -> ExcelTarget:
        # создаём при первой записи: dry-run и ошибки валидации не должны запускать Excel
        if self._target is None:
            self._target = ExcelTarget(self.config.workbook_path, self.config.log_sheet)
        return self._target

    def run(self, job_name: str, *, dry_run: bool = False, force: bool = False) -> RunResult:
        job = self.config.job(job_name)
        started = datetime.datetime.now()
        try:
            return self._run(job, started, dry_run=dry_run, force=force)
        except SyncError as e:
            self.state.record_run(job.name, FAILED, started, message=str(e)[:2000])
            raise

    def _run(self, job: JobConfig, started: datetime.datetime, *, dry_run: bool, force: bool) -> RunResult:
        t0 = time.perf_counter()

        raw = self.source.read_table(job.source_tab, job.source_range)
        numeric_idx, date_idx = resolve_columns([normalize_header(h) for h in raw.header], job)
        table, cleaning_issues, stats = clean_table(raw, numeric_idx, date_idx)

        last = self.state.last_success(job.name)
        report = validate(
            table, job.rules, numeric_idx=numeric_idx, stats=stats,
            previous_rows=last["rows"] if last else None, force=force,
        )
        report.issues[:0] = cleaning_issues
        report_path = self._save_report(job, report) if report.issues else None
        if not report.ok:
            raise ValidationFailed(report, report_path)

        diff = diff_tables(self.state.load_snapshot(job.name), table, job.rules.key_columns)
        unchanged = last is not None and last["fingerprint"] == table.fingerprint()

        backup_path = None
        if dry_run:
            status = DRY_RUN
        elif unchanged and not force:
            status = UNCHANGED
        else:
            text_cols = set(range(table.n_cols)) - numeric_idx - date_idx
            self.target.write_table(
                job.target_sheet, job.target_cell, table,
                text_cols=text_cols, timestamp_cell=job.timestamp_cell,
            )
            if job.backup:
                backup_path = save_backup(table, job.name, self.config.backup_dir, self.config.backup_keep)
            status = SUCCESS

        run_id = self.state.record_run(
            job.name, status, started, table=table, diff=diff,
            warnings=len(report.warnings), message=diff.summary().splitlines()[0],
        )
        if status == SUCCESS:
            self.state.save_snapshot(job.name, run_id, table)
            self.target.append_log([
                datetime.datetime.now(), job.name, status, table.n_rows,
                None if diff.first_run else diff.added, None if diff.first_run else diff.removed,
                None if diff.first_run else diff.changed_cells, len(report.warnings),
                diff.summary().splitlines()[0],
            ])

        return RunResult(job, status, table, report, diff, time.perf_counter() - t0, backup_path, report_path)

    def _save_report(self, job: JobConfig, report: ValidationReport) -> Path:
        directory = self.config.report_dir
        directory.mkdir(parents=True, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        path = directory / f"{job.name}_{stamp}.txt"
        title = f"Отчёт о проверке: {job.title} ({job.source_tab}), {stamp}"
        path.write_text(f"{title}\n{'=' * len(title)}\n\n{report.full_text()}\n", encoding="utf-8")
        rotate(directory, f"{job.name}_*.txt", self.config.backup_keep)
        return path
