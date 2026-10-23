"""
Командная строка: тот же синк, что по кнопке в Excel, плюс диагностика.

    python -m sheetsync run forecast            # одно задание
    python -m sheetsync run all --dry-run       # всё, но без записи в Excel
    python -m sheetsync run orders --force      # записать, даже если не изменилось
    python -m sheetsync history                 # журнал запусков
    python -m sheetsync check                   # проверить настройку на новом компьютере
"""

import argparse
import re
import sys

from .config import load_config
from .errors import SyncError, ValidationFailed
from .runner import SyncRunner
from .state import StateStore


def main(argv=None) -> int:
    # консоль Windows может не уметь в часть символов — пусть лучше «?», чем падение
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")

    parser = argparse.ArgumentParser(prog="python -m sheetsync", description="Синк Google Sheets → Excel")
    parser.add_argument("--config", help="путь к config.yaml (по умолчанию — в корне проекта)")
    commands = parser.add_subparsers(dest="command", required=True)

    run_p = commands.add_parser("run", help="выполнить синк")
    run_p.add_argument("job", help="имя задания из config.yaml или all")
    run_p.add_argument("--dry-run", action="store_true", help="прочитать и проверить, но не писать в Excel")
    run_p.add_argument("--force", action="store_true",
                       help="писать, даже если данные не изменились или строк резко стало меньше")

    hist_p = commands.add_parser("history", help="история запусков")
    hist_p.add_argument("job", nargs="?")
    hist_p.add_argument("-n", type=int, default=15, help="сколько последних запусков показать")

    commands.add_parser("check", help="проверить конфиг, ключ, доступ к таблице и книгу Excel")

    args = parser.parse_args(argv)
    try:
        config = load_config(args.config)
        if args.command == "run":
            return _run(config, args)
        if args.command == "history":
            return _history(config, args)
        return _check(config)
    except SyncError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 1


def _run(config, args) -> int:
    names = list(config.jobs) if args.job == "all" else [config.job(args.job).name]
    runner = SyncRunner(config)
    exit_code = 0
    for name in names:
        try:
            result = runner.run(name, dry_run=args.dry_run, force=args.force)
            print(result.message())
        except ValidationFailed as e:
            print(f"«{name}»: {e}", file=sys.stderr)
            exit_code = 2
        except SyncError as e:
            print(f"«{name}»: ошибка: {e}", file=sys.stderr)
            exit_code = max(exit_code, 1)
        print("-" * 60)
    return exit_code


def _history(config, args) -> int:
    with StateStore(config.state_db) as state:
        runs = state.history(args.job, args.n)
    if not runs:
        print("Запусков ещё не было.")
        return 0
    print(f"{'#':>4}  {'время':19}  {'задание':9}  {'статус':9}  {'строк':>6}  {'+':>5} {'-':>5} {'изм.':>6}  комментарий")
    for r in runs:
        nums = [r["added"], r["removed"], r["changed_cells"]]
        added, removed, changed = ("" if v is None else v for v in nums)
        comment = (r["message"] or "").splitlines()[0][:70] if r["message"] else ""
        print(f"{r['id']:>4}  {r['started_at']:19}  {r['job']:9}  {r['status']:9}  {r['rows'] or '':>6}  "
              f"{added:>5} {removed:>5} {changed:>6}  {comment}")
    return 0


def _check(config) -> int:
    """Пошаговая проверка установки: на чём именно сломалось — сразу видно."""
    from .auth import load_credentials
    from .source import GoogleSheetSource

    ok = True

    def report(passed: bool, text: str):
        nonlocal ok
        ok &= passed
        print(f"[{'OK' if passed else '!!'}] {text}")

    report(True, f"конфиг прочитан: {config.path}")
    try:
        _, email = load_credentials(config.credentials_file)
        report(True, f"ключ Google: {config.credentials_file.name}" + (f" (сервисный аккаунт {email})" if email else ""))
    except SyncError as e:
        report(False, str(e))
        return 1

    source = GoogleSheetSource(config.spreadsheet_id, config.credentials_file)
    try:
        spreadsheet = source.spreadsheet()
        report(True, f"таблица открыта: «{spreadsheet.title}»")
    except SyncError as e:
        report(False, str(e))
        return 1

    for job in config.jobs.values():
        # читаем только строку заголовков — проверка должна быть быстрой
        header_range = None
        if job.source_range:  # B4:BP3000 → B4:BP4
            start, end = job.source_range.split(":")
            header_row = re.search(r"\d+", start).group()
            end_column = re.match(r"[A-Z]+", end).group()
            header_range = f"{start}:{end_column}{header_row}"
        try:
            ws = spreadsheet.worksheet(job.source_tab)
            header = ws.get(header_range or "1:1")
            found = header[0] if header else []
            missing = [c for c in job.rules.required_columns if c not in [str(h).strip() for h in found]]
            report(not missing, f"задание {job.name}: лист '{job.source_tab}'"
                   + (f" — нет столбцов {missing}" if missing else f", столбцов: {len(found)}"))
        except Exception as e:
            report(False, f"задание {job.name}: лист '{job.source_tab}' — {e}")

    if config.workbook_path.exists():
        import openpyxl
        wb = openpyxl.load_workbook(config.workbook_path, read_only=True)
        sheets = set(wb.sheetnames)
        wb.close()
        report(True, f"книга Excel: {config.workbook_path.name}")
        for job in config.jobs.values():
            report(job.target_sheet in sheets, f"  лист '{job.target_sheet}' для задания {job.name}")
    else:
        report(False, f"нет книги Excel {config.workbook_path} — создайте: python scripts/create_demo_workbook.py")

    print("\nВсё готово к синку." if ok else "\nЕсть проблемы — см. строки с [!!].")
    return 0 if ok else 1
