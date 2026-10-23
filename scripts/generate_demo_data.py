"""
Генератор синтетических демо-данных: продажи шести вымышленных металлургических
компаний.

Компании, клиенты и числа выдуманы. Реалистичными сделаны только порядки
величин, сверенные с открытыми источниками: цены горячекатаного проката и
арматуры в 2023–2025 годах держались в коридоре примерно $450–750 за тонну,
трубы большого диаметра стоят примерно в два раза дороже, а крупный
производитель отгружает от нескольких сотен тысяч до 1–1,5 млн т в месяц.

Получаются два листа, по устройству такие же, как в реальной рабочей задаче:
  * Forecast Inputs — «широкая» таблица: строка на (метрика × компания ×
    продукт × рынок × канал), столбец на месяц 2023-01…2027-12; факт до
    LAST_ACTUAL, дальше прогноз. Около 1300 строк × 60 месяцев.
  * Orders — «длинный» журнал заказов, 40 000 строк.

В данные специально подмешана «грязь», которая встречается в живых таблицах
(числа текстом с запятой, «n/a», #REF!, лишние пробелы, пустая строка-
разделитель, даты в разных форматах, ручные правки выручки). Список того, что
подмешано, печатается в конце и описан в README.

Генерация детерминирована (SEED): при каждом запуске получаются одни и те же
данные.

Запуск:
    python scripts/generate_demo_data.py          # сохранит demo_data/demo_source.xlsx
Загрузить в Google Sheets — scripts/seed_google_sheet.py.
"""

import datetime
import math
import random
from pathlib import Path

SEED = 42

COMPANIES = {
    "STG": "СтальГрад",
    "NRM": "НордМет",
    "VLP": "ВолгаПрокат",
    "SBT": "СибТрубСталь",
    "URL": "УралЛист",
    "KMM": "КамаМеталл",
}

# код: (название, базовая цена $/т, зависит ли спрос от строительного сезона)
PRODUCTS = {
    "HRC": ("Горячекатаный прокат", 600, False),
    "CRC": ("Холоднокатаный прокат", 700, False),
    "GALV": ("Оцинкованный прокат", 850, False),
    "PLATE": ("Толстый лист", 760, False),
    "REBAR": ("Арматура", 550, True),
    "WIRE": ("Катанка", 580, True),
    "BEAM": ("Фасонный прокат", 650, True),
    "LDP": ("Трубы большого диаметра", 1150, False),
    "OCTG": ("Нефтегазовые трубы", 1350, False),
    "SLAB": ("Сляб", 480, False),
}

# код: (название, коэффициент к цене) — экспорт обычно дешевле внутреннего рынка
MARKETS = {
    "DOM": ("Россия", 1.00),
    "CIS": ("СНГ", 0.97),
    "TUR": ("Турция", 0.92),
    "MENA": ("Ближний Восток и Африка", 0.90),
    "SEA": ("Юго-Восточная Азия", 0.88),
    "LATAM": ("Латинская Америка", 0.91),
}

CHANNELS = {
    "DIR": ("Прямые продажи", 1.00),
    "TRD": ("Трейдеры", 0.97),
    "SVC": ("Сервисные центры", 1.04),
}

# компания: (продукты, отгрузки всего в месяц, тыс. т)
PORTFOLIO = {
    "STG": (["HRC", "CRC", "GALV", "PLATE", "SLAB", "LDP"], 950),
    "NRM": (["HRC", "CRC", "GALV", "SLAB", "REBAR"], 1250),
    "VLP": (["REBAR", "WIRE", "BEAM", "HRC"], 600),
    "SBT": (["LDP", "OCTG", "PLATE"], 320),
    "URL": (["HRC", "CRC", "GALV", "PLATE", "REBAR", "BEAM", "WIRE"], 900),
    "KMM": (["REBAR", "WIRE", "BEAM", "SLAB"], 400),
}

METRICS = [("Volume", "kt"), ("Price", "$/t"), ("Revenue", "$m")]

FIRST_MONTH = (2023, 1)
LAST_MONTH = (2027, 12)
LAST_ACTUAL = (2026, 9)

ORDERS_COUNT = 40_000
ORDERS_START = datetime.date(2025, 1, 1)
ORDERS_END = datetime.date(2026, 10, 23)

OUTPUT = Path(__file__).resolve().parent.parent / "demo_data" / "demo_source.xlsx"


def month_list():
    months = []
    year, month = FIRST_MONTH
    while (year, month) <= LAST_MONTH:
        months.append((year, month))
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return months


def price_cycle(t: float) -> float:
    """Множитель цены от времени (t — месяцы с января 2023): мягкий спад 2023–2024
    и восстановление к 2026, как в отраслевом цикле."""
    return 1 + 0.10 * math.cos(2 * math.pi * t / 40) - 0.0015 * t


def season(month: int, construction: bool) -> float:
    """Январский провал у всех; у строительных продуктов пик летом."""
    factor = 0.86 if month == 1 else 1.0
    if construction:
        factor *= 1 + 0.18 * math.sin(2 * math.pi * (month - 4) / 12)
    return factor


# --- Forecast Inputs --------------------------------------------------------


def build_forecast(rng: random.Random):
    """Возвращает (заголовок, строки данных) без грязи."""
    months = month_list()
    header = ["Metric", "Code", "Company", "Product", "Market", "Channel", "Unit"]
    header += [f"{y}-{m:02d}" for y, m in months]

    rows = []
    for comp, (products, total_kt) in PORTFOLIO.items():
        markets = ["DOM"] + rng.sample([m for m in MARKETS if m != "DOM"], rng.randint(3, 5))
        combos = [(p, m, c) for p in products for m in markets for c in CHANNELS]
        # Доля каждой комбинации в отгрузках: внутренний рынок и прямые продажи крупнее
        weights = [
            rng.uniform(0.5, 1.5) * (3 if m == "DOM" else 1) * (2 if c == "DIR" else 1)
            for p, m, c in combos
        ]
        weight_sum = sum(weights)

        for (prod, market, channel), w in zip(combos, weights):
            _, base_price, construction = PRODUCTS[prod]
            share_kt = total_kt * w / weight_sum
            growth = rng.uniform(-0.03, 0.06)  # годовой тренд объёмов этой комбинации
            volumes, prices = [], []
            for t, (year, month) in enumerate(months):
                forecast = (year, month) > LAST_ACTUAL
                noise = 0.02 if forecast else 0.07  # прогноз ровнее факта
                vol = share_kt * season(month, construction) * (1 + growth) ** (t / 12)
                vol *= rng.uniform(1 - noise, 1 + noise)
                price = base_price * MARKETS[market][1] * CHANNELS[channel][1] * price_cycle(t)
                price *= rng.uniform(1 - noise / 2, 1 + noise / 2)
                volumes.append(round(vol, 2))
                prices.append(round(price, 1))
            revenues = [round(v * p / 1000, 3) for v, p in zip(volumes, prices)]

            dims = [COMPANIES[comp], PRODUCTS[prod][0], MARKETS[market][0], CHANNELS[channel][0]]
            for (metric, unit), values in zip(METRICS, (volumes, prices, revenues)):
                code = f"{metric[:3].upper()}-{comp}-{prod}-{market}-{channel}"
                rows.append([metric, code, *dims, unit, *values])
    return header, rows


def forecast_sheet(header, rows):
    """Раскладка листа как в рабочей таблице: данные с B4, над месяцами — ACT/FCT."""
    months = header[7:]
    last_actual = f"{LAST_ACTUAL[0]}-{LAST_ACTUAL[1]:02d}"
    labels = ["ACT" if m <= last_actual else "FCT" for m in months]
    title = "Прогноз продаж металлургических компаний — СИНТЕТИЧЕСКИЕ ДАННЫЕ"
    note = "Объём — тыс. т, цена — $/т, выручка — млн $. Факт до " + last_actual
    sheet = [
        ["", title],
        ["", note],
        [""] * 8 + labels,
        [""] + header,
    ]
    sheet += [[""] + row for row in rows]
    return sheet


def dirty_forecast(rows, rng: random.Random):
    """Портит данные «как в жизни». Возвращает список того, что сделано."""
    log = []
    first_month_col = 7
    n_months = len(rows[0]) - first_month_col

    def random_cell():
        return rng.randrange(len(rows)), first_month_col + rng.randrange(n_months)

    for _ in range(25):  # число, вставленное текстом из русского Excel
        r, c = random_cell()
        value = rows[r][c]
        text = f"{value:,.2f}".replace(",", " ").replace(".", ",")
        rows[r][c] = text
    log.append("25 чисел записаны текстом с запятой и пробелами: '1 234,56'")

    for marker, count in (("n/a", 8), ("-", 4)):
        for _ in range(count):
            r, c = random_cell()
            rows[r][c] = marker
        log.append(f"{count} ячеек со значением '{marker}' (нет данных)")

    for marker, count in (("#REF!", 3), ("уточнить", 3)):
        for _ in range(count):
            r, c = random_cell()
            rows[r][c] = marker
        log.append(f"{count} ячеек со значением '{marker}' (мусор вместо числа)")

    revenue_rows = [i for i, row in enumerate(rows) if row[0] == "Revenue"]
    for r in rng.sample(revenue_rows, 4):
        c = first_month_col + rng.randrange(n_months)
        if isinstance(rows[r][c], float):
            rows[r][c] = round(rows[r][c] * 1.2, 3)
    log.append("4 ячейки выручки «поправлены руками» на +20% — не сходятся с объём × цена")

    for r in rng.sample(range(len(rows)), 10):
        rows[r][2] = rows[r][2] + " "
    log.append("у 10 строк в названии компании лишний пробел в конце")

    middle = len(rows) // 2
    rows.insert(middle, [""] * len(rows[0]))
    log.append("в середину таблицы вставлена пустая строка-разделитель")
    return log


# --- Orders -----------------------------------------------------------------


def build_orders(rng: random.Random):
    header = [
        "order_id", "order_date", "company", "customer_inn", "region",
        "product", "volume_t", "price_usd_t", "status",
    ]
    # Пул клиентов: ИНН из 10 цифр, у компаний из регионов 01–09 он начинается
    # с нуля — Excel «съедает» такие нули, если записать ИНН как число.
    customers = []
    for _ in range(800):
        region = rng.choice([rng.randint(1, 9), rng.randint(10, 89)])
        customers.append(f"{region:02d}{rng.randint(0, 99_999_999):08d}")

    companies = list(PORTFOLIO)
    sizes = [PORTFOLIO[c][1] for c in companies]
    statuses = ["shipped", "confirmed", "cancelled", "return"]
    status_weights = [70, 18, 8, 4]
    days = (ORDERS_END - ORDERS_START).days

    dates = sorted(ORDERS_START + datetime.timedelta(days=rng.randrange(days + 1)) for _ in range(ORDERS_COUNT))
    rows = []
    for i, date in enumerate(dates, start=1):
        comp = rng.choices(companies, weights=sizes)[0]
        prod = rng.choice(PORTFOLIO[comp][0])
        market = rng.choices(list(MARKETS), weights=[5, 2, 1, 1, 1, 1])[0]
        status = rng.choices(statuses, weights=status_weights)[0]
        t = (date.year - FIRST_MONTH[0]) * 12 + date.month - 1
        price = PRODUCTS[prod][1] * MARKETS[market][1] * price_cycle(t) * rng.uniform(0.95, 1.05)
        volume = round(rng.lognormvariate(math.log(60), 0.8), 1)
        if status == "return":
            volume = -volume
        rows.append([
            f"ORD-{i:06d}", date.isoformat(), COMPANIES[comp], rng.choice(customers),
            MARKETS[market][0], PRODUCTS[prod][0], volume, round(price, 1), status,
        ])
    return header, rows


def dirty_orders(rows, rng: random.Random):
    log = []
    picks = rng.sample(range(len(rows)), 200)
    groups = iter([picks[0:40], picks[40:45], picks[45:105], picks[105:120], picks[120:126], picks[126:156], picks[156:159]])

    for r in next(groups):
        y, m, d = rows[r][1].split("-")
        rows[r][1] = f"{d}.{m}.{y}"
    log.append("40 дат в русском формате ДД.ММ.ГГГГ вместо ГГГГ-ММ-ДД")

    for r, bad in zip(next(groups), ["31.02.2026", "2026-13-05", "вчера", "2026-02-30", "00.00.0000"]):
        rows[r][1] = bad
    log.append("5 невозможных дат: 31.02.2026, 2026-13-05, 'вчера' и т.п.")

    for r in next(groups):
        rows[r][6] = str(rows[r][6]).replace(".", ",")
    log.append("60 объёмов записаны текстом с запятой: '25,5'")

    for r in next(groups):
        rows[r][7] = ""
    log.append("15 заказов без цены")

    for r in next(groups):
        rows[r][7] = "n/a"
    log.append("6 цен со значением 'n/a'")

    for r in next(groups):
        rows[r][2] = "  " + rows[r][2] + " "
    log.append("30 названий компаний с пробелами по краям")

    for r in next(groups):
        rows[r][4] = "   "
    log.append("3 региона из одних пробелов")
    return log


def build_all(seed: int = SEED):
    """Всё сразу: листы целиком (как они лягут в Google Sheets) + журнал грязи."""
    rng = random.Random(seed)
    f_header, f_rows = build_forecast(rng)
    f_log = dirty_forecast(f_rows, rng)
    o_header, o_rows = build_orders(rng)
    o_log = dirty_orders(o_rows, rng)
    return {
        "Forecast Inputs": forecast_sheet(f_header, f_rows),
        "Orders": [o_header] + o_rows,
    }, {"Forecast Inputs": f_log, "Orders": o_log}


def main():
    import openpyxl

    sheets, logs = build_all()
    OUTPUT.parent.mkdir(exist_ok=True)
    wb = openpyxl.Workbook(write_only=True)
    for name, values in sheets.items():
        ws = wb.create_sheet(name)
        for row in values:
            ws.append(row)
    wb.save(OUTPUT)

    for name, values in sheets.items():
        print(f"{name}: {len(values)} строк × {max(len(r) for r in values)} столбцов")
        for line in logs[name]:
            print("   -", line)
    print(f"\nСохранено: {OUTPUT}")


if __name__ == "__main__":
    main()
