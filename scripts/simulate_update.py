"""
Имитирует «месячное обновление» источника — для демонстрации на защите:

  * прогноз цен на горячекатаный прокат у «СтальГрада» поднимается на 5% во
    всех прогнозных месяцах, выручка пересчитывается (объём × цена / 1000);
  * в журнал заказов добавляются 200 новых заказов за сегодня.

После этого синк покажет, что именно изменилось.

    python scripts/simulate_update.py [--price-change 5] [--new-orders 200]
"""

import argparse
import datetime
import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from gspread.utils import ValueRenderOption, rowcol_to_a1  # noqa: E402

import generate_demo_data as demo  # noqa: E402
from sheetsync.config import load_config  # noqa: E402
from sheetsync.errors import SyncError  # noqa: E402
from sheetsync.source import GoogleSheetSource  # noqa: E402

COMPANY, PRODUCT = "STG", "HRC"
HEADER_ROW, FIRST_COL = 4, 2  # данные прогноза начинаются с B4


def bump_forecast_prices(ws, change_pct: float) -> int:
    values = ws.get(value_render_option=ValueRenderOption.unformatted)
    header = values[HEADER_ROW - 1][FIRST_COL - 1:]
    last_actual = f"{demo.LAST_ACTUAL[0]}-{demo.LAST_ACTUAL[1]:02d}"
    forecast_cols = [i for i, h in enumerate(header) if isinstance(h, str) and h > last_actual and h[:2] == "20"]

    rows = {}  # код → (номер строки в листе, значения строки начиная с B)
    for number, row in enumerate(values[HEADER_ROW:], start=HEADER_ROW + 1):
        data = row[FIRST_COL - 1:]
        if len(data) > 1 and f"-{COMPANY}-{PRODUCT}-" in str(data[1]):
            rows[data[1]] = (number, data + [""] * (len(header) - len(data)))

    updated = 0
    for code, (number, price_row) in rows.items():
        if not code.startswith("PRI-"):
            continue
        suffix = code[len("PRI-"):]
        vol_number, vol_row = rows[f"VOL-{suffix}"]
        rev_number, rev_row = rows[f"REV-{suffix}"]
        for i in forecast_cols:
            price, volume = price_row[i], vol_row[i]
            if isinstance(price, (int, float)) and isinstance(volume, (int, float)):
                price_row[i] = round(price * (1 + change_pct / 100), 1)
                rev_row[i] = round(volume * price_row[i] / 1000, 3)
                updated += 1
        for number_, row in ((number, price_row), (rev_number, rev_row)):
            start = rowcol_to_a1(number_, FIRST_COL)
            ws.update([row], start, value_input_option="RAW")
    return updated


def add_orders(ws, count: int) -> int:
    last_id = ws.col_values(1)[-1]  # 'ORD-040000'
    next_number = int(last_id.split("-")[1]) + 1
    rng = random.Random()
    today = datetime.date.today().isoformat()
    rows = []
    for i in range(count):
        comp = rng.choice(list(demo.PORTFOLIO))
        prod = rng.choice(demo.PORTFOLIO[comp][0])
        market = rng.choice(list(demo.MARKETS))
        price = demo.PRODUCTS[prod][1] * demo.MARKETS[market][1] * rng.uniform(0.95, 1.05)
        rows.append([
            f"ORD-{next_number + i:06d}", today, demo.COMPANIES[comp], f"{rng.randint(10, 89)}{rng.randint(0, 10**8 - 1):08d}",
            demo.MARKETS[market][0], demo.PRODUCTS[prod][0], round(rng.lognormvariate(4.1, 0.8), 1), round(price, 1),
            "confirmed",
        ])
    ws.append_rows(rows, value_input_option="RAW")
    return count


def main():
    parser = argparse.ArgumentParser(description="Имитация обновления источника для демо")
    parser.add_argument("--config")
    parser.add_argument("--price-change", type=float, default=5.0, help="изменение цен, %%")
    parser.add_argument("--new-orders", type=int, default=200)
    args = parser.parse_args()

    try:
        config = load_config(args.config)
        sheet = GoogleSheetSource(config.spreadsheet_id, config.credentials_file).spreadsheet()
    except SyncError as e:
        sys.exit(f"Ошибка: {e}")

    cells = bump_forecast_prices(sheet.worksheet("Forecast Inputs"), args.price_change)
    print(f"Прогноз: цены HRC у {demo.COMPANIES[COMPANY]} изменены на {args.price_change:+g}% в {cells} ячейках, выручка пересчитана.")
    added = add_orders(sheet.worksheet("Orders"), args.new_orders)
    print(f"Заказы: добавлено {added} новых заказов за сегодня.")


if __name__ == "__main__":
    main()
