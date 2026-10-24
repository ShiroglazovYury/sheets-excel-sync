"""
Иерархия ошибок проекта.

Все «ожидаемые» проблемы (нет ключа, нет доступа, сломалась структура листа)
наследуются от SyncError. Их текст пишется для человека и показывается в
Excel/консоли как есть — без traceback. Всё, что не SyncError, считается
багом программы: его ловит самый верхний уровень и пишет traceback в файл.
"""


class SyncError(Exception):
    """Базовая ошибка синка: её текст написан для человека.

    Зачем нужен: все «ожидаемые» ошибки наследуются от него, поэтому точки входа
    (excel_entry, cli) ловят их одним except SyncError и показывают текст без
    traceback. Всё, что не SyncError, считается багом программы.
    """


class ConfigError(SyncError):
    """Ошибка в config.yaml: нет файла, ключа, неверный тип или значение.

    Выбрасывается в config.py при загрузке конфига.
    """


class AuthError(SyncError):
    """Не удалось авторизоваться в Google: нет ключа, он битый или отозван.

    Выбрасывается в auth.py и source.py.
    """


class SourceError(SyncError):
    """Не удалось прочитать данные из Google Sheets.

    Например: нет доступа к таблице, нет листа, нет связи с Google, источник
    пуст. Выбрасывается в source.py.
    """


class TargetError(SyncError):
    """Не удалось записать данные в Excel: нет книги, листа или ошибка записи.

    Выбрасывается в excel_target.py.
    """


class StateError(SyncError):
    """Проблема с файлом состояния state/sync_state.sqlite (например, он повреждён).

    Выбрасывается в state.py.
    """


class ValidationFailed(SyncError):
    """Данные не прошли проверки — Excel не трогаем.

    Выбрасывается раннером после validate(), если в отчёте есть хотя бы одна
    проблема уровня error.
    """

    def __init__(self, report, report_path=None):
        """Сохраняет отчёт и собирает из него текст ошибки.

        Зачем нужна: текст исключения сразу содержит короткую сводку проблем и путь
        к полному отчёту — его можно показать в окне Excel как есть. Сам отчёт
        остаётся в атрибуте report, если нужна подробная информация.

        Аргументы:
            report: ValidationReport с найденными проблемами.
            report_path: путь к файлу с полным отчётом в reports/ (или None).
        """
        self.report = report
        self.report_path = report_path
        text = "Данные не прошли проверку, Excel не изменён.\n\n" + report.summary()
        if report_path:
            text += f"\n\nПолный отчёт: {report_path}"
        super().__init__(text)
