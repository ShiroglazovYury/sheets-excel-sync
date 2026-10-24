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
    """Синк одного задания — эту функцию вызывает кнопка в Excel.

    Зачем нужна: макрос VBA вызывает её через xlwings:
    RunPython "import sheetsync.excel_entry as e; e.run('forecast')".
    Итог или ошибка показываются во всплывающем окне Excel.

    Аргументы:
        job_name: имя задания из config.yaml.
        force: записать, даже если данные не изменились или строк резко меньше.
    """
    _run_jobs([job_name], force)


def run_all(force: bool = False):
    """Синк всех заданий из config.yaml подряд — кнопка «Обновить всё».

    Зачем нужна: одно нажатие обновляет и прогноз, и заказы, а итог по всем
    заданиям собирается в одно окно.

    Аргументы:
        force: то же, что в run().
    """
    _run_jobs(None, force)


def _run_jobs(names, force):
    """Общий код для run() и run_all(): запускает задания и показывает окно с итогом.

    Зачем нужна: из кнопки Excel нельзя «упасть» — при RunPython traceback
    выводится в некрасивом окне, а человеку нужен ответ, что делать. Поэтому:
        * SyncError по заданию → текст ошибки в окне, остальные задания
          продолжают выполняться;
        * SyncError при загрузке конфига → окно с текстом ошибки;
        * любое другое исключение (баг программы) → окно с кратким текстом, а
          traceback — в файле reports/crash_*.txt.
    Заголовок и значок окна зависят от того, были ли проблемы.

    Аргументы:
        names: список имён заданий или None — все задания из конфига.
        force: передаётся в SyncRunner.run().
    """
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
    """Сохраняет traceback текущего исключения в reports/crash_<время>.txt.

    Зачем нужна: если случился баг, в окне Excel видно только короткое
    сообщение, а полный traceback для разбора лежит в файле. Путь берётся от
    корня проекта, а не из конфига: падение могло случиться как раз при чтении
    конфига.

    Возвращает:
        Путь к созданному файлу.
    """
    directory = PROJECT_ROOT / "reports"
    directory.mkdir(exist_ok=True)
    path = directory / f"crash_{datetime.datetime.now():%Y-%m-%d_%H-%M-%S}.txt"
    path.write_text(traceback.format_exc(), encoding="utf-8")
    return path
