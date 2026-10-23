"""
Создаёт демо-книгу «управленки» excel/Steel_Management_Report.xlsx:

  * Summary  — сводка на формулах (выручка по компаниям и годам, отгрузки
               и возвраты), которая сама пересчитывается после синка;
  * Forecast — сюда синк кладёт прогноз (с ячейки B4);
  * Orders   — сюда синк кладёт журнал заказов (с ячейки A4).

Формулы ссылаются на листы с данными так же, как в реальной модели, — на
этом видно, зачем писать через xlwings, а не пересобирать файл.

С флагом --buttons (по умолчанию включён) скрипт через xlwings рисует на
листах кнопки и привязывает их к макросам из PERSONAL.XLSB (vba/SheetSync.bas).

    python scripts/create_demo_workbook.py [--overwrite] [--no-buttons]
"""

import argparse
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from generate_demo_data import COMPANIES  # noqa: E402
from sheetsync.config import load_config  # noqa: E402
from sheetsync.errors import SyncError  # noqa: E402

YEARS = [2024, 2025, 2026, 2027]
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
HEAD_FILL = PatternFill("solid", fgColor="DDE7F0")

# Forecast: данные с B4 → B=Metric, D=Company, I..BP — 60 месяцев 2023-01…2027-12
FORECAST_FORMULA = (
    '=SUMPRODUCT((Forecast!$B$5:$B$5000="Revenue")*(Forecast!$D$5:$D$5000=$A{row})'
    '*(LEFT(Forecast!$I$4:$BP$4,4)={col}$5&"")*Forecast!$I$5:$BP$5000)'
)

# (текст на кнопке, макрос в PERSONAL.XLSB, лист, ячейка левого верхнего угла)
BUTTONS = [
    ("Обновить всё", "SyncAll", "Summary", "H2"),
    ("Обновить прогноз", "SyncForecast", "Summary", "H4"),
    ("Обновить заказы", "SyncOrders", "Summary", "H6"),
    ("Обновить прогноз", "SyncForecast", "Forecast", "F1"),
    ("Обновить заказы", "SyncOrders", "Orders", "E1"),
]


def build_workbook(path: Path):
    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary["A1"] = "Управленческий отчёт: продажи металлургов (СИНТЕТИЧЕСКИЕ ДАННЫЕ)"
    summary["A1"].font = TITLE
    summary["A2"] = "Цифры ниже — формулы по листам Forecast и Orders. Обновляются кнопками справа."

    summary["A4"] = "Выручка, млн $ (лист Forecast)"
    summary["A4"].font = BOLD
    _header(summary, 5, ["Компания", *YEARS])
    first = 6
    for i, name in enumerate(COMPANIES.values()):
        row = first + i
        summary.cell(row, 1, name)
        for j in range(len(YEARS)):
            col = openpyxl.utils.get_column_letter(2 + j)
            cell = summary.cell(row, 2 + j, FORECAST_FORMULA.format(row=row, col=col))
            cell.number_format = "#,##0.0"
    total_row = first + len(COMPANIES)
    summary.cell(total_row, 1, "Итого").font = BOLD
    for j in range(len(YEARS)):
        col = openpyxl.utils.get_column_letter(2 + j)
        cell = summary.cell(total_row, 2 + j, f"=SUM({col}{first}:{col}{total_row - 1})")
        cell.font, cell.number_format = BOLD, "#,##0.0"

    start = total_row + 3
    summary.cell(start - 1, 1, "Заказы (лист Orders)").font = BOLD
    _header(summary, start, ["Компания", "Отгружено, т", "Заказов", "Ср. цена отгрузки, $/т", "Возвраты, т"])
    for i, name in enumerate(COMPANIES.values()):
        r = start + 1 + i
        summary.cell(r, 1, name)
        summary.cell(r, 2, f'=SUMIFS(Orders!$G:$G,Orders!$C:$C,$A{r},Orders!$I:$I,"shipped")').number_format = "#,##0"
        summary.cell(r, 3, f"=COUNTIFS(Orders!$C:$C,$A{r})").number_format = "#,##0"
        summary.cell(r, 4, f'=IFERROR(AVERAGEIFS(Orders!$H:$H,Orders!$C:$C,$A{r},Orders!$I:$I,"shipped"),"")').number_format = "#,##0.0"
        summary.cell(r, 5, f'=-SUMIFS(Orders!$G:$G,Orders!$C:$C,$A{r},Orders!$I:$I,"return")').number_format = "#,##0"
    summary.column_dimensions["A"].width = 18
    for col in "BCDE":
        summary.column_dimensions[col].width = 16

    forecast = wb.create_sheet("Forecast")
    forecast["B1"] = "Прогноз продаж — данные из Google Sheets, не редактировать вручную"
    forecast["B1"].font = BOLD
    forecast["B2"] = "Обновлено:"
    forecast.freeze_panes = "I5"

    orders = wb.create_sheet("Orders")
    orders["A1"] = "Журнал заказов — данные из Google Sheets, не редактировать вручную"
    orders["A1"].font = BOLD
    orders["A2"] = "Обновлено:"
    orders.freeze_panes = "A5"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _header(ws, row, values):
    for col, value in enumerate(values, start=1):
        cell = ws.cell(row, col, value)
        cell.font, cell.fill = BOLD, HEAD_FILL


def add_buttons(path: Path):
    """Кнопки (элементы управления формы) умеет рисовать только сам Excel."""
    import xlwings as xw

    # отдельный невидимый экземпляр Excel, чтобы не трогать уже открытые книги
    app = xw.App(visible=False, add_book=False)
    # без этого при ошибке Excel может повиснуть невидимым на диалоге «Сохранить изменения?»
    app.display_alerts = False
    try:
        book = app.books.open(str(path))
        for text, macro, sheet_name, cell in BUTTONS:
            sheet = book.sheets[sheet_name]
            anchor = sheet.range(cell)
            button = sheet.api.Buttons().Add(anchor.left, anchor.top, 130, 26)
            button.OnAction = f"PERSONAL.XLSB!{macro}"
            button.Caption = text
        book.save()
    finally:
        try:
            app.quit()
        except Exception:
            app.kill()


def main():
    parser = argparse.ArgumentParser(description="Создать демо-книгу Excel")
    parser.add_argument("--config", help="путь к config.yaml")
    parser.add_argument("--overwrite", action="store_true", help="перезаписать существующую книгу")
    parser.add_argument("--no-buttons", action="store_true", help="не рисовать кнопки (если Excel не установлен)")
    args = parser.parse_args()

    try:
        path = load_config(args.config).workbook_path
    except SyncError as e:
        sys.exit(f"Ошибка: {e}")
    if path.exists() and not args.overwrite:
        sys.exit(f"Книга уже есть: {path}\nЧтобы пересоздать — добавьте --overwrite (данные в ней пропадут).")

    build_workbook(path)
    print(f"Книга создана: {path}")
    if not args.no_buttons:
        try:
            add_buttons(path)
            print("Кнопки добавлены и привязаны к макросам PERSONAL.XLSB!SyncAll / SyncForecast / SyncOrders.")
        except Exception as e:
            print(f"Кнопки добавить не удалось ({e}). Нарисуйте их вручную — см. README, шаг 6.")


if __name__ == "__main__":
    main()
