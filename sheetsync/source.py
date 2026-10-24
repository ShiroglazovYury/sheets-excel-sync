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
    """HTTP-клиент gspread, который повторяет запрос при лимитах Google (429/5xx).

    Зачем нужен: Google иногда отвечает «слишком много запросов» или временной
    ошибкой сервера — тогда достаточно подождать и повторить. Штатный максимум
    паузы в gspread — 128 секунд; из Excel так долго ждать нельзя, режем до 16.
    """

    _MAX_BACKOFF = 16


class GoogleSheetSource:
    """Источник данных — Google-таблица, из которой читаются диапазоны.

    Зачем нужен: хранит открытое соединение с таблицей, чтобы все задания
    одного запуска не авторизовывались и не открывали таблицу заново, и
    превращает ошибки Google API в понятные сообщения.
    """
    def __init__(self, spreadsheet_id: str, credentials_file):
        """Запоминает, какую таблицу и с каким ключом открывать; в Google пока не ходит.

        Зачем нужна: соединение откладывается до первого чтения (см. spreadsheet()),
        поэтому создать объект можно даже без интернета.

        Аргументы:
            spreadsheet_id: ID таблицы из конфига.
            credentials_file: путь к ключу Google.
        """
        self.spreadsheet_id = spreadsheet_id
        self.credentials_file = credentials_file
        self.account_email = None
        self._spreadsheet = None

    def spreadsheet(self) -> gspread.Spreadsheet:
        """Открывает Google-таблицу (один раз за запуск) и возвращает её.

        Зачем нужна: авторизация и открытие таблицы — самые медленные и «капризные»
        шаги. Здесь они делаются один раз, результат запоминается, а ошибки
        переводятся на человеческий язык: «таблица не найдена — проверьте
        spreadsheet_id», «нет доступа — добавьте такой-то email». Запросы идут через
        _RetryingHTTPClient, который повторяет их при лимитах Google.

        Возвращает:
            Объект таблицы gspread.

        Исключения:
            AuthError: проблема с ключом.
            SourceError: таблица не найдена, нет доступа, нет связи с Google.
        """
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
        """Читает лист или диапазон из Google-таблицы и возвращает Table.

        Зачем нужна: это шаг «Чтение» в сценарии синка. Значения запрашиваются
        «как есть» (UNFORMATTED_VALUE) — иначе таблица с русской локалью отдала бы
        числа строками вида '1 234,5'. Даты запрашиваются строкой, как их видно в
        таблице, а не серийным числом.

        Аргументы:
            tab: имя листа (source.tab в конфиге).
            cell_range: диапазон вида "B4:BP3000" или None — весь заполненный лист.
                Первая непустая строка диапазона считается заголовком.

        Возвращает:
            Table с номерами строк, как в Google Sheets.

        Исключения:
            SourceError: нет листа, диапазон за границами листа, ошибка API,
                нет связи, источник пуст.
        """
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
    """Превращает «сырой» ответ API в прямоугольную Table.

    Зачем нужна: API возвращает «рваные» строки — хвостовые пустые ячейки просто
    обрезает, — а весь остальной код рассчитывает, что у каждой строки столько
    же ячеек, сколько заголовков. Поэтому:
        * полностью пустые строки (разделители, хвост диапазона «с запасом»)
          выкидываются, но номера оставшихся строк в листе запоминаются;
        * первая непустая строка становится заголовком;
        * каждая строка дополняется пустыми ячейками до общей ширины;
        * хвостовые столбцы без заголовка и без данных отбрасываются.

    Аргументы:
        raw: список строк из API.
        first_row: номер строки листа, с которой начинается диапазон (для B4:… — 4).
        where: описание источника для текста ошибки.

    Возвращает:
        Table.

    Исключения:
        SourceError: в диапазоне нет ни одной заполненной ячейки.
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
    """Есть ли в ячейке что-то кроме пустоты и пробелов.

    Зачем нужна: по ней решается, пустая ли строка целиком (такие строки
    выкидываются) и есть ли данные в хвостовых столбцах.
    """
    return value is not None and str(value).strip() != ""


def _first_row_number(cell_range: str | None) -> int:
    """Номер первой строки диапазона: для "B4:BP3000" — 4, без диапазона — 1.

    Зачем нужна: API возвращает строки без номеров, а в сообщениях об ошибках
    человеку нужен номер строки, который он увидит слева в Google Sheets.
    """
    if not cell_range:
        return 1
    match = re.match(r"[A-Z]+(\d+)", cell_range)
    return int(match.group(1)) if match else 1


def _translate(error: Exception) -> Exception:
    """Переводит исключения Google и сети в ошибки проекта с понятным текстом.

    Зачем нужна: библиотеки бросают десятки разных исключений с техническим
    текстом. Здесь они сводятся к понятным случаям: ключ отклонён (AuthError),
    нет связи с Google или ошибка API (SourceError). Неизвестное исключение
    возвращается как есть — это повод разбираться, а не прятать проблему.

    Аргументы:
        error: пойманное исключение.

    Возвращает:
        Исключение, которое нужно выбросить (raise).
    """
    if isinstance(error, RefreshError):
        return AuthError(f"Google отклонил ключ: {error}. Возможно, ключ отозван — выпустите новый.")
    if isinstance(error, (TransportError, requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
        return SourceError("Нет связи с Google. Проверьте интернет и повторите.")
    if isinstance(error, gspread.exceptions.APIError):
        return SourceError(f"Google Sheets API вернул ошибку {error.code}: {error.error.get('message')}")
    # неизвестная ошибка — пусть всплывёт как есть, это повод разбираться
    return error
