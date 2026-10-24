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
    """Запись таблиц в книгу Excel через запущенный Excel (xlwings, COM).

    Зачем нужен: держит найденную книгу, чтобы все задания одного запуска писали
    в один и тот же открытый файл, и прячет всё, что связано с Excel: поиск
    книги, очистку старых данных, форматы, запись кусками, журнал синков.

    Почему xlwings, а не openpyxl: openpyxl пересобирает файл целиком и теряет
    кэш формул, сводные и связи, а xlwings просто вставляет значения в открытую
    книгу.
    """
    def __init__(self, workbook_path: Path, log_sheet: str = "Sync Log"):
        """Запоминает путь к книге и имя листа журнала; сам Excel пока не трогает.

        Зачем нужна: книга ищется или открывается только при первой записи (см.
        book()), поэтому создание объекта ничего не стоит.

        Аргументы:
            workbook_path: путь к книге из config.yaml (excel.workbook).
            log_sheet: имя листа журнала синков.
        """
        self.workbook_path = Path(workbook_path)
        self.log_sheet = log_sheet
        self._book = None

    def book(self) -> xw.Book:
        """Возвращает книгу Excel: уже открытую или открывает её сама.

        Зачем нужна: найденная книга запоминается, и все следующие вызовы
        используют её без повторного поиска.

        Возвращает:
            Объект книги xlwings.
        """
        if self._book is None:
            self._book = self._find_open_book() or self._open_book()
        return self._book

    def _find_open_book(self):
        """Ищет книгу среди уже открытых в Excel по имени файла.

        Зачем нужна: обычно кнопку нажимают из открытой книги, и писать надо именно
        в неё. xw.Book(путь) сравнивает полный путь, а у файлов в папке, которую
        синхронизирует OneDrive, Excel считает путём https://…-адрес. Тогда xw.Book
        не узнаёт открытую книгу и пытается открыть её второй раз — с ошибкой. По
        имени файла книга узнаётся надёжно.

        Возвращает:
            Книгу xlwings или None, если такая не открыта.
        """
        name = self.workbook_path.name.lower()
        for app in xw.apps:
            for book in app.books:
                if book.name.lower() == name:
                    return book
        return None

    def _open_book(self):
        """Открывает книгу Excel по пути из конфига.

        Зачем нужна: если книга не открыта (например, синк запустили из консоли),
        её нужно открыть самим, а при неудаче объяснить, что не так.

        Возвращает:
            Книгу xlwings.

        Исключения:
            TargetError: файла нет или Excel не смог его открыть.
        """
        if not self.workbook_path.exists():
            raise TargetError(
                f"Не найдена книга Excel: {self.workbook_path}\n"
                "Проверьте путь excel.workbook в config.yaml."
            )
        try:
            return xw.Book(str(self.workbook_path))
        except Exception as e:
            raise TargetError(f"Excel не смог открыть {self.workbook_path.name}: {e}") from None

    def sheet(self, name: str) -> xw.Sheet:
        """Возвращает лист книги по имени.

        Зачем нужна: если листа нет (переименовали или удалили), вместо непонятной
        ошибки COM пользователь видит, какие листы в книге есть.

        Аргументы:
            name: имя листа из конфига (target.sheet).

        Исключения:
            TargetError: такого листа нет.
        """
        book = self.book()
        try:
            return book.sheets[name]
        except Exception:
            names = ", ".join(f"'{s.name}'" for s in book.sheets)
            raise TargetError(f"В книге {book.name} нет листа '{name}'. Есть: {names}.") from None

    def write_table(self, sheet_name: str, top_left: str, table: Table, *, text_cols: set[int],
                    timestamp_cell: str | None = None):
        """Записывает таблицу на лист Excel начиная с ячейки top_left.

        Зачем нужна: это шаг «Запись» в сценарии синка. По порядку:
            1. на время записи выключает пересчёт формул и обновление экрана —
               иначе Excel пересчитывает модель после каждого куска и «виснет»;
            2. стирает старые данные (_clear_old_data), чтобы не остался хвост;
            3. ставит формат «Текст» строке заголовков и текстовым столбцам — иначе
               Excel превратит заголовок '2026-03' в дату, а ИНН '0274…' в число без
               ведущего нуля;
            4. записывает заголовок и строки кусками по CHUNK_ROWS строк;
            5. ставит время обновления в timestamp_cell;
            6. в любом случае (finally) возвращает пересчёт и экран как было.

        Формат дат не задаётся специально: через COM Excel читает коды формата в
        локали Windows, и в русской локали 'dd.mm.yyyy' ломается. Дату в ячейке
        «Общего» формата Excel сам покажет в формате локали.

        Аргументы:
            sheet_name: лист для записи (target.sheet).
            top_left: левая верхняя ячейка, например "B4" (target.cell).
            table: очищенная таблица; пишется вместе с заголовком.
            text_cols: номера текстовых столбцов (все, кроме числовых и дат).
            timestamp_cell: ячейка для времени обновления или None.

        Исключения:
            TargetError: нет листа или Excel вернул ошибку при записи.
        """
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
        """Стирает всё от start вправо-вниз до края заполненной области листа.

        Зачем нужна: если в источнике стало меньше строк, без очистки снизу остался
        бы хвост старых данных, и формулы СУММЕСЛИ посчитали бы его вместе с
        новыми. Форматы тоже сбрасываются (clear, а не clear_contents): число,
        записанное в ячейку с форматом «Текст», Excel сохранит как текст. Строкой
        number_format = "General" сбросить нельзя — в русской локали COM ждёт
        «Основной» и падает.

        Аргументы:
            sheet: лист Excel.
            start: левая верхняя ячейка области данных; всё выше и левее не
                трогается (там, например, заголовок листа и время обновления).
        """
        last = sheet.used_range.last_cell
        if last.row >= start.row and last.column >= start.column:
            sheet.range(start, (last.row, last.column)).clear()

    def append_log(self, values: list):
        """Дописывает строку в лист журнала синков в книге Excel.

        Зачем нужна: история запусков видна прямо в Excel, без консоли. Лист
        создаётся при первом синке (в конце книги, с жирной строкой заголовков),
        дальше строка добавляется под последней заполненной.

        Аргументы:
            values: значения строки в порядке LOG_HEADER: время, задание, статус,
                строк, добавлено, удалено, изменено ячеек, предупреждений,
                комментарий.
        """
        book = self.book()
        try:
            sheet = book.sheets[self.log_sheet]
        except Exception:
            sheet = book.sheets.add(self.log_sheet, after=book.sheets[-1])
            sheet.range("A1").value = LOG_HEADER
            sheet.range("A1").resize(1, len(LOG_HEADER)).font.bold = True
        next_row = sheet.range((sheet.cells.last_cell.row, 1)).end("up").row + 1
        sheet.range((next_row, 1)).value = values
