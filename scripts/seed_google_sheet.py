"""
Загружает синтетические демо-данные (generate_demo_data.py) в Google-таблицу из
config.yaml. Нужен один раз — чтобы на новом компьютере появился «источник».

Сервисному аккаунту на этой таблице нужна роль «Редактор» (для самого синка
потом достаточно «Читателя»).

    python scripts/seed_google_sheet.py

Листы Forecast Inputs и Orders перезаписываются целиком.
"""

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import gspread  # noqa: E402

from generate_demo_data import build_all  # noqa: E402
from sheetsync.config import load_config  # noqa: E402
from sheetsync.errors import SyncError  # noqa: E402
from sheetsync.source import GoogleSheetSource  # noqa: E402

# Google ограничивает размер одного запроса — пишем кусками
CHUNK_ROWS = 5000


def upload(spreadsheet, title: str, values: list[list], min_rows: int):
    cols = max(len(r) for r in values)
    values = [r + [""] * (cols - len(r)) for r in values]
    rows = max(len(values), min_rows)
    try:
        ws = spreadsheet.worksheet(title)
        ws.clear()
        ws.resize(rows, cols)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title, rows, cols)

    for start in range(0, len(values), CHUNK_ROWS):
        chunk = values[start:start + CHUNK_ROWS]
        # RAW — строки остаются строками: «грязь» вроде '1 234,56' не должна
        # превратиться в число при загрузке, иначе нечего будет чистить
        ws.update(chunk, f"A{start + 1}", value_input_option="RAW")
        print(f"  {title}: {start + len(chunk)}/{len(values)} строк")
    ws.freeze(rows=4 if title == "Forecast Inputs" else 1)


def main():
    parser = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    parser.add_argument("--config", help="путь к config.yaml")
    args = parser.parse_args()

    try:
        config = load_config(args.config)
        spreadsheet = GoogleSheetSource(config.spreadsheet_id, config.credentials_file).spreadsheet()
    except SyncError as e:
        sys.exit(f"Ошибка: {e}")

    sheets, _ = build_all()
    for title, values in sheets.items():
        # лист должен быть не короче, чем range в конфиге, иначе синк упадёт на «exceeds grid limits»
        min_rows = 1
        for job in config.jobs.values():
            if job.source_tab == title and job.source_range:
                min_rows = int(re.search(r"(\d+)$", job.source_range).group(1))
        upload(spreadsheet, title, values, min_rows)

    # пустой лист по умолчанию («Лист1»/«Sheet1») только мешает
    for ws in spreadsheet.worksheets():
        is_empty = not any(any(cell for cell in row) for row in ws.get_all_values())
        if ws.title not in sheets and len(spreadsheet.worksheets()) > 1 and is_empty:
            spreadsheet.del_worksheet(ws)

    print(f"Готово: {spreadsheet.url}")


if __name__ == "__main__":
    main()
