"""
Чтение диапазона из Google Sheets.

Главное здесь — превратить сырые ошибки Google API в сообщения, по которым
понятно, что делать («таблица не расшарена на такой-то email», «на листе 1500
строк, а диапазон до 3000»), и привести ответ API к прямоугольной Table.
"""

import re

import gspread
import requests
from google.auth.exceptions import RefreshError, TransportError
from gspread.http_client import BackOffHTTPClient
from gspread.utils import DateTimeOption, ValueRenderOption

from .auth import load_credentials
from .errors import AuthError, SourceError
from .table import Table


class _RetryingHTTPClient(BackOffHTTPClient):
    """Повторяет запрос при 429/5xx (лимиты Google). Штатный максимум паузы
    в gspread — 128 секунд; из Excel так долго ждать нельзя, режем до 16."""

    _MAX_BACKOFF = 16


class GoogleSheetSource:
    def __init__(self, spreadsheet_id: str, credentials_file):
        self.spreadsheet_id = spreadsheet_id
        self.credentials_file = credentials_file
        self.account_email = None
        self._spreadsheet = None

    def spreadsheet(self) -> gspread.Spreadsheet:
        """Открывает таблицу один раз и переиспользует для всех заданий."""
        if self._spreadsheet is not None:
            return self._spreadsheet

        creds, self.account_email = load_credentials(self.credentials_file)
        client = gspread.authorize(creds, http_client=_RetryingHTTPClient)
        try:
            self._spreadsheet = client.open_by_key(self.spreadsheet_id)
        except gspread.SpreadsheetNotFound:
            raise SourceError(
                f"Таблица {self.spreadsheet_id} не найдена. Проверьте google.spreadsheet_id в config.yaml."
            ) from None
        except PermissionError:
            who = self.account_email or "аккаунт из token.json"
            raise SourceError(
                f"Нет доступа к таблице. Откройте её в браузере → «Настройки доступа» "
                f"и добавьте {who} (роли «Читатель» достаточно)."
            ) from None
        except Exception as e:
            raise _translate(e) from None
        return self._spreadsheet

    def read_table(self, tab: str, cell_range: str | None = None) -> Table:
        sheet = self.spreadsheet()
        try:
            ws = sheet.worksheet(tab)
        except gspread.WorksheetNotFound:
            tabs = ", ".join(f"'{w.title}'" for w in sheet.worksheets())
            raise SourceError(f"В таблице нет листа '{tab}'. Есть: {tabs}.") from None
        except Exception as e:
            raise _translate(e) from None

        try:
            raw = ws.get(
                cell_range,
                # UNFORMATTED — числа приходят числами. Иначе таблица с русской
                # локалью отдаст '1 234,5' строкой, и это не отличить от мусора.
                value_render_option=ValueRenderOption.unformatted,
                # даты — строкой, как их видно в таблице, а не серийным числом 45000
                date_time_render_option=DateTimeOption.formatted_string,
            )
        except gspread.exceptions.APIError as e:
            if "exceeds grid limits" in str(e):
                raise SourceError(
                    f"Диапазон {cell_range} выходит за границы листа '{tab}' "
                    f"({ws.row_count} строк × {ws.col_count} столбцов). Добавьте строк в "
                    f"сам лист Google Sheets или уменьшите range в config.yaml."
                ) from None
            raise _translate(e) from None
        except Exception as e:
            raise _translate(e) from None

        first_row = _first_row_number(cell_range)
        return to_table(raw, first_row, where=f"лист '{tab}'" + (f", диапазон {cell_range}" if cell_range else ""))


def to_table(raw: list[list], first_row: int, where: str) -> Table:
    """Сырые строки API → прямоугольная Table.

    API возвращает «рваные» строки: хвостовые пустые ячейки просто обрезаются,
    поэтому дополняем каждую строку до ширины заголовка. Полностью пустые
    строки (разделители, хвост диапазона с запасом) выкидываем, но помним
    номера оставшихся строк в исходном листе.
    """
    numbered = [(first_row + i, row) for i, row in enumerate(raw) if any(_filled(v) for v in row)]
    if not numbered:
        raise SourceError(f"Источник пуст: {where}. Проверьте, что данные на месте и диапазон верный.")

    (header_row, header), *data = numbered
    width = max(len(row) for _, row in numbered)
    # хвостовые столбцы без заголовка и без данных не нужны
    while width > len(header) and not any(len(row) >= width and _filled(row[width - 1]) for _, row in data):
        width -= 1

    header = [("" if h is None else str(h)) for h in header] + [""] * (width - len(header))
    rows = [list(row[:width]) + [""] * (width - len(row)) for _, row in data]
    return Table(header=header, rows=rows, row_numbers=[n for n, _ in data])


def _filled(value) -> bool:
    return value is not None and str(value).strip() != ""


def _first_row_number(cell_range: str | None) -> int:
    if not cell_range:
        return 1
    match = re.match(r"[A-Z]+(\d+)", cell_range)
    return int(match.group(1)) if match else 1


def _translate(error: Exception) -> Exception:
    """Сетевые и прочие ошибки Google → наши, с человеческим текстом."""
    if isinstance(error, RefreshError):
        return AuthError(f"Google отклонил ключ: {error}. Возможно, ключ отозван — выпустите новый.")
    if isinstance(error, (TransportError, requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
        return SourceError("Нет связи с Google. Проверьте интернет и повторите.")
    if isinstance(error, gspread.exceptions.APIError):
        return SourceError(f"Google Sheets API вернул ошибку {error.code}: {error.error.get('message')}")
    # неизвестная ошибка — пусть всплывёт как есть, это повод разбираться
    return error
