"""
Table — прямоугольная таблица «заголовок + строки», которой обмениваются все
модули: источник её читает, очистка и проверки её разбирают, Excel её пишет,
состояние хранит её снимок.

Для каждой строки помним её номер в исходном листе Google Sheets
(row_numbers): после удаления пустых строк индексы уже не совпадают с
номерами в таблице, а в сообщениях об ошибках человеку нужен именно номер,
который он увидит слева в Google Sheets.
"""

import datetime
import hashlib
import json
from dataclasses import dataclass, field
from typing import Any


@dataclass
class Table:
    """Прямоугольная таблица «заголовок + строки», которой обмениваются все модули.

    Зачем нужен: источник её читает, очистка и проверки её разбирают, Excel её
    пишет, состояние хранит её снимок — всем модулям нужен один общий формат.

    Поля:
        header: названия столбцов.
        rows: строки — списки значений той же длины, что header.
        row_numbers: номер каждой строки в исходном листе Google Sheets. После
            удаления пустых строк индекс в списке уже не равен номеру в таблице,
            а в сообщениях об ошибках человеку нужен именно номер из Google.
    """
    header: list[str]
    rows: list[list[Any]]
    row_numbers: list[int] = field(default_factory=list)

    def __post_init__(self):
        """Проверяет таблицу сразу после создания.

        Зачем нужна: если номера строк не переданы (например, в тестах), они
        проставляются по порядку начиная со 2 (строка 1 — заголовок). Если переданы,
        но их количество не совпадает с числом строк, — это ошибка в коде, и лучше
        узнать о ней сразу, а не на середине синка.

        Исключения:
            ValueError: row_numbers и rows разной длины.
        """
        if not self.row_numbers:
            self.row_numbers = list(range(2, len(self.rows) + 2))
        if len(self.row_numbers) != len(self.rows):
            raise ValueError("row_numbers и rows разной длины")

    @property
    def n_rows(self) -> int:
        """Число строк данных (без строки заголовков)."""
        return len(self.rows)

    @property
    def n_cols(self) -> int:
        """Число столбцов."""
        return len(self.header)

    def column_index(self, name: str) -> int:
        """Номер столбца по его названию.

        Зачем нужна: в конфиге столбцы заданы именами, а строки таблицы — это
        списки, к которым обращаются по номеру.

        Аргументы:
            name: название столбца.

        Исключения:
            ValueError: такого столбца нет (стандартное поведение list.index).
        """
        return self.header.index(name)

    def values(self) -> list[list[Any]]:
        """Заголовок + строки одним двумерным списком — в таком виде данные пишутся в Excel.

        Зачем нужна: xlwings записывает двумерный список одним вызовом начиная с
        ячейки. Возвращаются копии списков, чтобы запись не могла случайно изменить
        саму таблицу.
        """
        return [list(self.header)] + [list(r) for r in self.rows]

    def to_jsonable(self) -> dict:
        """Превращает таблицу в словарь, который можно сохранить в JSON.

        Зачем нужна: снимок в базе состояния хранится как JSON, а JSON не умеет
        хранить даты — они превращаются в строки ISO (2026-03-15T00:00:00).
        """
        return {
            "header": self.header,
            "rows": [[_jsonable(v) for v in row] for row in self.rows],
            "row_numbers": self.row_numbers,
        }

    @classmethod
    def from_jsonable(cls, data: dict) -> "Table":
        """Собирает таблицу обратно из словаря, сохранённого to_jsonable().

        Зачем нужна: так загружается снимок из базы состояния. Даты при этом
        остаются строками — diff учитывает это при сравнении.

        Аргументы:
            data: словарь с ключами header, rows, row_numbers.
        """
        return cls(header=data["header"], rows=data["rows"], row_numbers=data["row_numbers"])

    def fingerprint(self) -> str:
        """Отпечаток содержимого таблицы — хэш SHA-256 от заголовка и строк.

        Зачем нужна: чтобы быстро понять, изменилось ли что-то с прошлого синка, не
        сравнивая десятки тысяч ячеек: одинаковые данные дают одинаковый отпечаток.
        Номера строк в отпечаток не входят — если в источнике просто добавили
        пустую строку, данные не считаются изменившимися.

        Возвращает:
            Строку из 64 шестнадцатеричных символов.
        """
        payload = json.dumps([self.header, self.to_jsonable()["rows"]], ensure_ascii=False)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _jsonable(value):
    """Значение для JSON: дату или время — строкой ISO, остальное как есть."""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.isoformat()
    return value
