"""
Точки входа для кнопок в Excel. Макрос VBA вызывает их через xlwings:

    RunPython "import sheetsync.excel_entry as e; e.run('forecast')"

Здесь нельзя «упасть»: при RunPython traceback показывается в некрасивом
окне, а человеку у кнопки нужен ответ, что делать. Поэтому все ошибки
превращаются в сообщение в окне Excel, а неожиданные (баги) ещё и
записываются с traceback в файл reports/crash_*.txt.
"""

import datetime
import traceback

import xlwings as xw

from .config import PROJECT_ROOT, load_config
from .errors import SyncError
from .runner import SyncRunner

_SEPARATOR = "\n\n" + "─" * 30 + "\n\n"


def run(job_name: str, force: bool = False):
    _run_jobs([job_name], force)


def run_all(force: bool = False):
    _run_jobs(None, force)


def _run_jobs(names, force):
    messages, failed = [], False
    try:
        config = load_config()
        runner = SyncRunner(config)
        for name in names or list(config.jobs):
            # задания независимы: ошибка в одном не отменяет остальные
            try:
                messages.append(runner.run(name, force=force).message())
            except SyncError as e:
                failed = True
                messages.append(f"«{name}»: {e}")
    except SyncError as e:
        failed = True
        messages.append(str(e))
    except Exception as e:
        failed = True
        messages.append(
            f"Неожиданная ошибка программы: {type(e).__name__}: {e}\n\n"
            f"Подробности сохранены в {_save_crash_log()}"
        )

    title = "Синк: есть проблемы" if failed else "Синк завершён"
    xw.apps.active.alert(_SEPARATOR.join(messages), title=title, mode="critical" if failed else "info")


def _save_crash_log():
    directory = PROJECT_ROOT / "reports"
    directory.mkdir(exist_ok=True)
    path = directory / f"crash_{datetime.datetime.now():%Y-%m-%d_%H-%M-%S}.txt"
    path.write_text(traceback.format_exc(), encoding="utf-8")
    return path
