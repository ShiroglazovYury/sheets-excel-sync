"""
Сравнение нового среза данных с прошлым успешным синком.

Строки сопоставляем по ключу (key_columns), а не по номеру: если в середину
источника вставили строку, сравнение «по позиции» показало бы, что
изменилось всё ниже неё. Без ключа остаётся только сравнение по позиции.

Результат — короткая сводка для человека: сколько строк добавилось и
пропало, сколько ячеек изменилось и какие изменения самые крупные.
"""

import datetime
from dataclasses import dataclass, field

from .table import Table


@dataclass
class CellChange:
    """Одна изменившаяся ячейка между прошлым и новым синком.

    Зачем нужен: diff показывает человеку самые крупные изменения вида
    «PRI-STG-HRC-DOM-SVC [2026-10]: 627.40 → 658.80 (+5.0%)» — это одна такая
    запись.

    Поля:
        key: ключ строки (например код метрики или order_id).
        column: название столбца.
        old, new: значение до и после.
    """
    key: str
    column: str
    old: object
    new: object

    @property
    def delta(self) -> float | None:
        """Насколько изменилось число: new − old, или None, если значения не числа.

        Зачем нужна: по модулю этой разницы изменения сортируются — сначала
        показываются самые крупные. У текстовых изменений разницы нет.
        """
        if _is_number(self.old) and _is_number(self.new):
            return self.new - self.old
        return None

    def __str__(self):
        """Текст изменения для отчёта: «ключ [столбец]: старое → новое (+5.0%)».

        Зачем нужна: процент добавляется, только если оба значения числа и старое
        не ноль — иначе процент не имеет смысла.
        """
        text = f"{self.key} [{self.column}]: {_fmt(self.old)} → {_fmt(self.new)}"
        if self.delta is not None and self.old:
            text += f" ({self.delta / abs(self.old) * 100:+.1f}%)"
        return text


@dataclass
class DiffResult:
    """Итог сравнения нового среза данных с прошлым успешным синком.

    Зачем нужен: в окне Excel и в истории запусков видно, что именно поменялось
    в источнике с прошлого раза, — а не просто «данные записаны».

    Поля:
        first_run: прошлого снимка нет, сравнивать не с чем.
        added, removed: сколько строк появилось и пропало (по ключу).
        changed_cells: сколько ячеек изменилось в строках, которые есть в обоих
            срезах.
        columns_added, columns_removed: новые и пропавшие столбцы.
        top_changes: самые крупные изменения (CellChange).
    """
    first_run: bool = False
    added: int = 0
    removed: int = 0
    changed_cells: int = 0
    columns_added: list[str] = field(default_factory=list)
    columns_removed: list[str] = field(default_factory=list)
    top_changes: list[CellChange] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        """Есть ли хоть какие-то отличия от прошлого синка.

        Зачем нужна: короткий ответ «да/нет» для сводки. Первый синк считается
        изменением — до него в Excel данных не было.
        """
        return self.first_run or bool(
            self.added or self.removed or self.changed_cells or self.columns_added or self.columns_removed
        )

    def summary(self) -> str:
        """Короткий текст о том, что изменилось, — для окна Excel и консоли.

        Зачем нужна: первую строку этого текста раннер ещё и сохраняет в журнал
        запусков и на лист «Sync Log».

        Возвращает:
            Несколько строк: счётчики строк и ячеек, новые и пропавшие столбцы,
            список крупнейших изменений.
        """
        if self.first_run:
            return "Первый синк — сравнивать не с чем."
        if not self.has_changes:
            return "Данные не изменились с прошлого синка."
        lines = [
            f"Строк добавлено: {self.added}, удалено: {self.removed}; "
            f"изменено ячеек: {self.changed_cells}."
        ]
        if self.columns_added:
            lines.append("Новые столбцы: " + ", ".join(self.columns_added))
        if self.columns_removed:
            lines.append("Пропали столбцы: " + ", ".join(self.columns_removed))
        if self.top_changes:
            lines.append("Крупнейшие изменения:")
            lines += [f"  • {c}" for c in self.top_changes]
        return "\n".join(lines)


def diff_tables(old: Table | None, new: Table, key_columns=(), top_n: int = 5) -> DiffResult:
    """Сравнивает две таблицы и возвращает, что изменилось.

    Зачем нужна: показать человеку, что поменялось в источнике с прошлого
    синка. Строки сопоставляются по ключу (key_columns), а не по номеру: если в
    середину источника вставили строку, сравнение по номеру показало бы, что
    изменилось всё ниже неё. Сравниваются только столбцы, которые есть в обеих
    таблицах; появившиеся и пропавшие столбцы перечисляются отдельно.

    Аргументы:
        old: прошлый снимок из state.load_snapshot() или None (первый синк).
        new: только что очищенная таблица.
        key_columns: столбцы-ключ; если пусто или их нет — сравнение по номеру
            строки.
        top_n: сколько крупнейших изменений показать.

    Возвращает:
        DiffResult.
    """
    if old is None:
        return DiffResult(first_run=True)

    result = DiffResult(
        columns_added=[h for h in new.header if h and h not in old.header],
        columns_removed=[h for h in old.header if h and h not in new.header],
    )
    common = [h for h in new.header if h and h in old.header]

    old_rows = _index_rows(old, key_columns)
    new_rows = _index_rows(new, key_columns)
    result.added = len(new_rows.keys() - old_rows.keys())
    result.removed = len(old_rows.keys() - new_rows.keys())

    old_pos = {h: old.column_index(h) for h in common}
    new_pos = {h: new.column_index(h) for h in common}
    changes = []
    for key in new_rows.keys() & old_rows.keys():
        old_row, new_row = old_rows[key], new_rows[key]
        for h in common:
            before, after = _norm(old_row[old_pos[h]]), _norm(new_row[new_pos[h]])
            if _same(before, after):
                continue
            result.changed_cells += 1
            changes.append(CellChange(" / ".join(map(str, key)), h, before, after))

    # сначала самые большие по модулю числовые изменения, потом текстовые
    changes.sort(key=lambda c: -abs(c.delta) if c.delta is not None else 0)
    result.top_changes = changes[:top_n]
    return result


def _index_rows(table: Table, key_columns) -> dict:
    """Строит словарь «ключ строки → строка» для быстрого сопоставления.

    Зачем нужна: по словарю та же строка в другом срезе находится сразу, даже
    если строки переставили. Если ключевых столбцов нет, ключом становится
    номер строки в листе.

    Аргументы:
        table: таблица.
        key_columns: столбцы-ключ.

    Возвращает:
        dict: ключ (кортеж значений) → строка.
    """
    if key_columns and all(c in table.header for c in key_columns):
        idx = [table.column_index(c) for c in key_columns]
        return {tuple(row[i] for i in idx): row for row in table.rows}
    return {(f"строка {n}",): row for n, row in zip(table.row_numbers, table.rows)}


def _norm(value):
    """Приводит значение к виду, в котором оно хранится в снимке.

    Зачем нужна: снимок хранится в JSON, поэтому даты в нём — строки
    '2026-03-15T00:00:00'. Новую таблицу приводим к тому же виду, иначе каждая
    дата считалась бы изменённой. Пустая строка и None тоже считаются
    одинаковыми — это одна и та же пустая ячейка.
    """
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.isoformat()
    return None if value == "" else value


def _same(a, b) -> bool:
    """Равны ли два значения с точки зрения diff.

    Зачем нужна: числа сравниваются с крошечным допуском — в Python 0.1 + 0.2 не
    равно 0.3 ровно, и без допуска такие «изменения» засоряли бы отчёт.
    Остальные значения сравниваются как есть.
    """
    if _is_number(a) and _is_number(b):
        return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
    return a == b


def _is_number(value) -> bool:
    """Число ли это (int или float, но не True/False).

    Зачем нужна: в Python bool — подкласс int, и без этой оговорки True
    считалось бы числом 1.
    """
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _fmt(value) -> str:
    """Форматирует значение для текста изменений.

    Дробные числа — с двумя знаками и пробелами между разрядами, None — словом
    «пусто», остальное — как есть.
    """
    if value is None:
        return "пусто"
    if isinstance(value, float):
        return f"{value:,.2f}".replace(",", " ")
    return str(value)
