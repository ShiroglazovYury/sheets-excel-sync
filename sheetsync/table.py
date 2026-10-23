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
    header: list[str]
    rows: list[list[Any]]
    row_numbers: list[int] = field(default_factory=list)

    def __post_init__(self):
        if not self.row_numbers:
            self.row_numbers = list(range(2, len(self.rows) + 2))
        if len(self.row_numbers) != len(self.rows):
            raise ValueError("row_numbers и rows разной длины")

    @property
    def n_rows(self) -> int:
        return len(self.rows)

    @property
    def n_cols(self) -> int:
        return len(self.header)

    def column_index(self, name: str) -> int:
        return self.header.index(name)

    def values(self) -> list[list[Any]]:
        """Заголовок + строки одним списком — в таком виде данные пишутся в Excel."""
        return [list(self.header)] + [list(r) for r in self.rows]

    def to_jsonable(self) -> dict:
        return {
            "header": self.header,
            "rows": [[_jsonable(v) for v in row] for row in self.rows],
            "row_numbers": self.row_numbers,
        }

    @classmethod
    def from_jsonable(cls, data: dict) -> "Table":
        return cls(header=data["header"], rows=data["rows"], row_numbers=data["row_numbers"])

    def fingerprint(self) -> str:
        """Хэш содержимого: по нему понимаем, что данные не изменились с прошлого раза."""
        payload = json.dumps([self.header, self.to_jsonable()["rows"]], ensure_ascii=False)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _jsonable(value):
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.isoformat()
    return value
