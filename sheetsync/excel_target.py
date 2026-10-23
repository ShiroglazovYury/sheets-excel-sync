"""
Запись в Excel через xlwings — то есть через сам запущенный Excel (COM).

Почему не openpyxl: openpyxl пересобирает файл целиком и теряет
закэшированные результаты формул, сводные, макросы, связи. А xlwings
просто вставляет значения в уже открытую книгу — как если бы человек сделал
«вставить значения», только за секунды и без зависаний.
"""

import datetime
from pathlib import Path

import xlwings as xw

from .errors import TargetError
from .table import Table

LOG_HEADER = ["Время", "Задание", "Статус", "Строк", "Добавлено", "Удалено",
              "Изменено ячеек", "Предупреждений", "Комментарий"]

# Строк за один COM-вызов: на десятках тысяч строк одна гигантская запись
# иногда падает по памяти/таймауту, кусками надёжнее.
CHUNK_ROWS = 5000


class ExcelTarget:
    def __init__(self, workbook_path: Path, log_sheet: str = "Sync Log"):
        self.workbook_path = Path(workbook_path)
        self.log_sheet = log_sheet
        self._book = None

    def book(self) -> xw.Book:
        if self._book is None:
            self._book = self._find_open_book() or self._open_book()
        return self._book

    def _find_open_book(self):
        """Ищем книгу среди уже открытых по имени файла.

        xw.Book(путь) сравнивает полный путь, а у файлов в папке, которую
        синхронизирует OneDrive, Excel считает путём https://…-адрес. Тогда
        xw.Book не узнаёт открытую книгу и пытается открыть её второй раз —
        с ошибкой. По имени файла узнаём надёжно.
        """
        name = self.workbook_path.name.lower()
        for app in xw.apps:
            for book in app.books:
                if book.name.lower() == name:
                    return book
        return None

    def _open_book(self):
        if not self.workbook_path.exists():
            raise TargetError(
                f"Не найдена книга Excel: {self.workbook_path}\n"
                "Создайте её: python scripts/create_demo_workbook.py — или поправьте excel.workbook в config.yaml."
            )
        try:
            return xw.Book(str(self.workbook_path))
        except Exception as e:
            raise TargetError(f"Excel не смог открыть {self.workbook_path.name}: {e}") from None

    def sheet(self, name: str) -> xw.Sheet:
        book = self.book()
        try:
            return book.sheets[name]
        except Exception:
            names = ", ".join(f"'{s.name}'" for s in book.sheets)
            raise TargetError(f"В книге {book.name} нет листа '{name}'. Есть: {names}.") from None

    def write_table(self, sheet_name: str, top_left: str, table: Table, *, text_cols: set[int],
                    timestamp_cell: str | None = None):
        sheet = self.sheet(sheet_name)
        app = sheet.book.app
        old_calculation = app.calculation
        app.screen_updating = False
        # без этого Excel пересчитывает формулы после каждого куска записи
        app.calculation = "manual"
        try:
            start = sheet.range(top_left)
            self._clear_old_data(sheet, start)

            n_rows, n_cols = table.n_rows + 1, table.n_cols  # +1 — строка заголовков
            first_data_row, last_row = start.row + 1, start.row + n_rows - 1

            # Формат «Текст» для заголовков и текстовых столбцов. Иначе Excel сам
            # «распознает» значения: заголовок '2026-03' станет датой 01.03.2026,
            # а ИНН '0274051582' — числом 274051582 без ведущего нуля.
            #
            # Даты формат не задаём специально: через COM Excel читает коды формата
            # в локали Windows, и 'dd.mm.yyyy' в русской локали сохраняется как
            # буквальный текст. Зато дату, записанную в ячейку «Общий», Excel сам
            # показывает коротким форматом даты текущей локали — так работает везде.
            sheet.range(start, (start.row, start.column + n_cols - 1)).number_format = "@"
            if table.n_rows:
                for col in sorted(text_cols):
                    sheet.range((first_data_row, start.column + col),
                                (last_row, start.column + col)).number_format = "@"

            start.options(chunksize=CHUNK_ROWS).value = table.values()

            if timestamp_cell:
                stamp = sheet.range(timestamp_cell)
                stamp.clear_formats()  # вернуть «Общий» формат, чтобы Excel сам показал дату
                stamp.value = datetime.datetime.now()
        except TargetError:
            raise
        except Exception as e:
            raise TargetError(f"Ошибка записи в лист '{sheet_name}': {e}") from None
        finally:
            app.calculation = old_calculation
            app.screen_updating = True

    @staticmethod
    def _clear_old_data(sheet: xw.Sheet, start: xw.Range):
        """Стираем всё от start вправо-вниз до края заполненной области.

        Без этого при уменьшении числа строк в источнике снизу останется хвост
        старых данных, и формулы СУММЕСЛИ посчитают его вместе с новыми.
        Форматы тоже сбрасываем (clear, а не clear_contents): число, записанное
        в ячейку с форматом «Текст», Excel сохранит как текст. Строкой
        number_format = "General" сбросить нельзя — в русской локали COM
        ждёт «Основной» и падает.
        """
        last = sheet.used_range.last_cell
        if last.row >= start.row and last.column >= start.column:
            sheet.range(start, (last.row, last.column)).clear()

    def append_log(self, values: list):
        """Дописывает строку в лист журнала (создаёт лист при первом синке)."""
        book = self.book()
        try:
            sheet = book.sheets[self.log_sheet]
        except Exception:
            sheet = book.sheets.add(self.log_sheet, after=book.sheets[-1])
            sheet.range("A1").value = LOG_HEADER
            sheet.range("A1").resize(1, len(LOG_HEADER)).font.bold = True
        next_row = sheet.range((sheet.cells.last_cell.row, 1)).end("up").row + 1
        sheet.range((next_row, 1)).value = values
