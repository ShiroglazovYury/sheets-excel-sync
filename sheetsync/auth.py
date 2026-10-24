"""
Авторизация в Google.

Основной вариант — сервисный аккаунт: JSON-ключ из Google Cloud Console. Его не
нужно обновлять через браузер, и он не протухает через 7 дней, как
refresh-токен личного OAuth-приложения в режиме Testing. Таблицу достаточно
один раз «расшарить» на email сервисного аккаунта.

Для разработки поддержан и обычный OAuth-токен пользователя (token.json, у
которого есть refresh_token), — тип определяется по содержимому файла.
"""

import json
from pathlib import Path

from google.oauth2 import credentials as user_credentials
from google.oauth2 import service_account

from .errors import AuthError

# Доступ к Google Sheets. Сама программа таблицу только читает; запись
# в источник ограничена ролью «Читатель», выданной сервисному аккаунту.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def load_credentials(path: Path):
    """Возвращает (credentials, email сервисного аккаунта или None)."""
    if not path.exists():
        raise AuthError(
            f"Не найден файл ключа Google: {path}\n"
            "Создайте сервисный аккаунт в Google Cloud Console, скачайте JSON-ключ "
            "и положите его по этому пути."
        )
    try:
        info = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        raise AuthError(f"Файл ключа {path.name} повреждён — это не JSON. Скачайте ключ заново.") from None

    if info.get("type") == "service_account":
        try:
            creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        except (ValueError, KeyError) as e:
            raise AuthError(f"Ключ сервисного аккаунта неполный: {e}. Скачайте ключ заново.") from None
        return creds, info.get("client_email")

    if "refresh_token" in info:
        try:
            creds = user_credentials.Credentials.from_authorized_user_info(info, SCOPES)
        except ValueError as e:
            raise AuthError(f"OAuth-токен неполный: {e}") from None
        return creds, None

    raise AuthError(
        f"Не понимаю формат {path.name}: это не ключ сервисного аккаунта "
        "(нет \"type\": \"service_account\") и не OAuth-токен (нет refresh_token)."
    )
