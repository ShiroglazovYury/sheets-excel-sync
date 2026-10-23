"""
Иерархия ошибок проекта.

Все «ожидаемые» проблемы (нет ключа, нет доступа, сломалась структура листа)
наследуются от SyncError. Их текст пишется для человека и показывается в
Excel/консоли как есть — без traceback. Всё, что не SyncError, считается
багом программы: его ловит самый верхний уровень и пишет traceback в файл.
"""


class SyncError(Exception):
    """Базовая ошибка синка: текст понятен пользователю."""


class ConfigError(SyncError):
    """Ошибка в config.yaml."""


class AuthError(SyncError):
    """Не удалось авторизоваться в Google."""


class SourceError(SyncError):
    """Не удалось прочитать данные из Google Sheets."""


class TargetError(SyncError):
    """Не удалось записать данные в Excel."""


class StateError(SyncError):
    """Проблема с файлом состояния (SQLite)."""


class ValidationFailed(SyncError):
    """Данные не прошли проверки — Excel не трогаем."""

    def __init__(self, report, report_path=None):
        self.report = report
        self.report_path = report_path
        text = "Данные не прошли проверку, Excel не изменён.\n\n" + report.summary()
        if report_path:
            text += f"\n\nПолный отчёт: {report_path}"
        super().__init__(text)
