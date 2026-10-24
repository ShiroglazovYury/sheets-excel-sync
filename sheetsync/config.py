"""
Загрузка config.yaml в типизированные объекты.

Конфиг проверяется целиком сразу при загрузке: лучше упасть до похода в
Google с понятным «в config.yaml нет ключа jobs.orders.target.sheet», чем
через минуту на середине синка с KeyError.

Все относительные пути в конфиге считаются от папки, где лежит сам конфиг, —
так программа работает одинаково, откуда бы её ни запустили (из консоли или
из Excel, у которого своя текущая папка).
"""

import os
import re
from dataclasses import dataclass
from pathlib import Path

import yaml

from .errors import ConfigError

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config.yaml"

# Переменная окружения, которой можно подменить путь к ключу, не трогая конфиг.
CREDENTIALS_ENV = "SHEETSYNC_CREDENTIALS"

_RANGE_RE = re.compile(r"^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$")
_CELL_RE = re.compile(r"^[A-Z]+[1-9]\d*$")
_SHEET_ID_RE = re.compile(r"/d/([A-Za-z0-9_-]+)")

_REQUIRED = object()


@dataclass(frozen=True)
class RevenueCheck:
    """Правило «выручка = объём × цена / divisor» для строк одной группы.

    Зачем нужен: в прогнозе объём, цена и выручка лежат отдельными строками.
    Если выручку «поправили руками» и она перестала сходиться с объёмом и
    ценой, это почти всегда ошибка в источнике. Правило ловит такие ячейки.

    Поля:
        metric_column: столбец с названием метрики (Metric).
        group_by: столбцы, которые вместе задают комбинацию (компания, продукт,
            рынок, канал); внутри неё ищутся строки объёма, цены и выручки.
        volume, price, revenue: значения metric_column для этих трёх строк.
        divisor: делитель для перевода единиц (kt × $/t = тыс. $ → млн $: 1000).
        tolerance_pct: допустимое расхождение, %.
    """

    metric_column: str
    group_by: tuple[str, ...]
    volume: str
    price: str
    revenue: str
    divisor: float = 1000.0  # kt × $/t = тыс. $, а выручка в млн $ → делим на 1000
    tolerance_pct: float = 1.0


@dataclass(frozen=True)
class ValidationRules:
    """Правила проверки данных одного задания (раздел validation в config.yaml).

    Зачем нужен: собрать все правила в один объект, который validate() получает
    целиком. Значения по умолчанию подобраны так, что пустой раздел validation
    почти ничего не проверяет.

    Поля:
        required_columns: столбцы, без которых данные писать нельзя.
        key_columns: столбцы-ключ строки (например order_id) — по ним ищутся
            дубли и сопоставляются строки при сравнении с прошлым синком.
        non_negative_columns: столбцы, где отрицательное значение подозрительно.
        all_numeric_non_negative: то же для всех числовых столбцов сразу
            (в конфиге non_negative: all).
        min_rows: если строк меньше — ошибка.
        max_row_drop_pct: на сколько процентов число строк может упасть с
            прошлого синка; None — не проверять.
        max_bad_values_pct: допустимая доля нечитаемых чисел, %.
        revenue_check: правило «выручка = объём × цена» или None.
    """
    required_columns: tuple[str, ...] = ()
    key_columns: tuple[str, ...] = ()
    non_negative_columns: tuple[str, ...] = ()
    all_numeric_non_negative: bool = False
    min_rows: int = 1
    max_row_drop_pct: float | None = None
    max_bad_values_pct: float = 5.0
    revenue_check: RevenueCheck | None = None


@dataclass(frozen=True)
class JobConfig:
    """Одно задание синка: какой диапазон откуда взять и куда положить.

    Зачем нужен: в config.yaml может быть сколько угодно заданий (сейчас
    forecast и orders), и все они обрабатываются одним и тем же кодом. Всё, чем
    задания отличаются, собрано здесь.

    Поля:
        name, title: короткое имя (для команд и имён файлов) и название для людей.
        source_tab, source_range: лист и диапазон в Google Sheets; range = None —
            весь заполненный лист.
        target_sheet, target_cell: лист в Excel и левая верхняя ячейка записи.
        timestamp_cell: куда поставить время обновления (или None).
        numeric_columns, numeric_pattern: какие столбцы числовые — по именам
            и/или по регулярному выражению (например, все месяцы вида 2026-03).
        date_columns: какие столбцы содержат даты.
        rules: правила проверки (ValidationRules).
        backup: делать ли бэкап после успешной записи.
    """

    name: str
    title: str
    source_tab: str
    source_range: str | None  # None — весь заполненный лист
    target_sheet: str
    target_cell: str
    timestamp_cell: str | None
    numeric_columns: tuple[str, ...]
    numeric_pattern: re.Pattern | None
    date_columns: tuple[str, ...]
    rules: ValidationRules
    backup: bool


@dataclass(frozen=True)
class AppConfig:
    """Вся конфигурация программы, загруженная из config.yaml.

    Зачем нужен: модули получают настройки готовым объектом с понятными полями
    и уже проверенными типами, а не разбирают словарь из YAML каждый сам. Все
    пути здесь уже абсолютные.

    Поля:
        path: откуда загружен конфиг.
        spreadsheet_id: ID Google-таблицы-источника.
        credentials_file: путь к ключу Google.
        workbook_path, log_sheet: книга Excel и имя листа журнала синков.
        state_db, backup_dir, backup_keep, report_dir: где хранить состояние,
            бэкапы (и сколько штук) и отчёты о проблемах в данных.
        jobs: задания синка по имени.
    """
    path: Path
    spreadsheet_id: str
    credentials_file: Path
    workbook_path: Path
    log_sheet: str
    state_db: Path
    backup_dir: Path
    backup_keep: int
    report_dir: Path
    jobs: dict[str, JobConfig]

    def job(self, name: str) -> JobConfig:
        """Возвращает задание по имени.

        Зачем нужна: вместо голого KeyError при опечатке (run sales) пользователь
        видит, какие задания вообще есть.

        Аргументы:
            name: имя задания, например "orders".

        Исключения:
            ConfigError: такого задания нет в config.yaml.
        """
        try:
            return self.jobs[name]
        except KeyError:
            raise ConfigError(
                f"В config.yaml нет задания '{name}'. Доступные: {', '.join(self.jobs)}"
            ) from None


def load_config(path=None) -> AppConfig:
    """Читает config.yaml и превращает его в проверенный AppConfig.

    Зачем нужна: это первое, что делает программа при любом запуске. Конфиг
    проверяется сразу и целиком: лучше упасть до похода в Google с сообщением
    «не заполнен ключ jobs.orders.target.sheet», чем через минуту посреди синка.

    Все относительные пути считаются от папки конфига — поэтому программа
    работает одинаково, откуда бы её ни запустили (из консоли или из Excel, у
    которого своя текущая папка). Путь к ключу можно подменить переменной
    окружения SHEETSYNC_CREDENTIALS.

    Аргументы:
        path: путь к конфигу; None — config.yaml в корне проекта.

    Возвращает:
        AppConfig.

    Исключения:
        ConfigError: файла нет, он не читается как YAML, не заполнены нужные
            ключи, неверные типы или в spreadsheet_id осталась заглушка.
    """
    path = Path(path) if path else DEFAULT_CONFIG_PATH
    if not path.exists():
        raise ConfigError(
            f"Не найден конфиг {path}.\n"
            "Скопируйте config.example.yaml в config.yaml и впишите ID своей таблицы."
        )
    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    except yaml.YAMLError as e:
        raise ConfigError(f"{path.name} не читается как YAML:\n{e}") from None
    if not isinstance(raw, dict):
        raise ConfigError(f"{path.name} пуст или имеет неверную структуру.")

    base = path.parent
    google = _section(raw, "google")
    excel = _section(raw, "excel")
    storage = _section(raw, "storage")

    spreadsheet_id = _get(google, "spreadsheet_id", "google", str)
    if "PASTE" in spreadsheet_id:
        raise ConfigError(
            "В config.yaml не заполнен google.spreadsheet_id — вставьте ссылку "
            "на свою таблицу или её ID (часть ссылки между /d/ и /edit)."
        )
    credentials = os.environ.get(CREDENTIALS_ENV) or _get(google, "credentials_file", "google", str)

    jobs_raw = _section(raw, "jobs")
    if not jobs_raw:
        raise ConfigError("В config.yaml нет ни одного задания в разделе jobs.")
    jobs = {name: _parse_job(name, body) for name, body in jobs_raw.items()}

    return AppConfig(
        path=path,
        spreadsheet_id=extract_spreadsheet_id(spreadsheet_id),
        credentials_file=base / credentials,
        workbook_path=base / _get(excel, "workbook", "excel", str),
        log_sheet=_get(excel, "log_sheet", "excel", str, default="Sync Log"),
        state_db=base / _get(storage, "state_db", "storage", str),
        backup_dir=base / _get(storage, "backup_dir", "storage", str),
        backup_keep=_get(storage, "backup_keep", "storage", int, default=10),
        report_dir=base / _get(storage, "report_dir", "storage", str),
        jobs=jobs,
    )


def extract_spreadsheet_id(url_or_id: str) -> str:
    """Достаёт ID таблицы из ссылки на неё или возвращает ID как есть.

    Зачем нужна: люди обычно вставляют в конфиг ссылку целиком
    (https://docs.google.com/spreadsheets/d/<ID>/edit…), а API нужен только ID —
    часть между /d/ и следующим слешем.

    Аргументы:
        url_or_id: ссылка или ID.

    Возвращает:
        ID таблицы.
    """
    match = _SHEET_ID_RE.search(url_or_id)
    return match.group(1) if match else url_or_id.strip()


def _parse_job(name: str, body) -> JobConfig:
    """Разбирает один раздел jobs.<имя> из YAML в объект JobConfig.

    Зачем нужна: проверяет всё, что относится к заданию: разделы source и target
    на месте, диапазон похож на B4:BO3000, ячейки похожи на B4, регулярное
    выражение numeric_pattern корректно. Адреса приводятся к верхнему регистру.

    Аргументы:
        name: имя задания (ключ в разделе jobs).
        body: содержимое раздела из YAML.

    Возвращает:
        JobConfig.

    Исключения:
        ConfigError: с точным путём к ошибочному ключу, например
            jobs.forecast.source.range.
    """
    where = f"jobs.{name}"
    if not isinstance(body, dict):
        raise ConfigError(f"Раздел {where} должен быть словарём.")
    source = _section(body, "source", where)
    target = _section(body, "target", where)
    columns = body.get("columns") or {}
    validation = body.get("validation") or {}

    source_range = _get(source, "range", f"{where}.source", str, default=None)
    if source_range is not None:
        source_range = source_range.upper().replace(" ", "")
        if not _RANGE_RE.match(source_range):
            raise ConfigError(f"{where}.source.range: '{source_range}' не похож на диапазон вида B4:BO3000.")

    target_cell = _get(target, "cell", f"{where}.target", str).upper()
    timestamp_cell = _get(target, "timestamp_cell", f"{where}.target", str, default=None)
    for label, cell in (("cell", target_cell), ("timestamp_cell", timestamp_cell)):
        if cell is not None and not _CELL_RE.match(cell.upper()):
            raise ConfigError(f"{where}.target.{label}: '{cell}' не похож на адрес ячейки вида B4.")

    pattern = _get(columns, "numeric_pattern", f"{where}.columns", str, default=None)
    try:
        numeric_pattern = re.compile(pattern) if pattern else None
    except re.error as e:
        raise ConfigError(f"{where}.columns.numeric_pattern — неверное регулярное выражение: {e}") from None

    return JobConfig(
        name=name,
        title=_get(body, "title", where, str, default=name),
        source_tab=_get(source, "tab", f"{where}.source", str),
        source_range=source_range,
        target_sheet=_get(target, "sheet", f"{where}.target", str),
        target_cell=target_cell,
        timestamp_cell=timestamp_cell.upper() if timestamp_cell else None,
        numeric_columns=_get(columns, "numeric", f"{where}.columns", list, default=()),
        numeric_pattern=numeric_pattern,
        date_columns=_get(columns, "date", f"{where}.columns", list, default=()),
        rules=_parse_rules(validation, f"{where}.validation"),
        backup=_get(body, "backup", where, bool, default=True),
    )


def _parse_rules(raw: dict, where: str) -> ValidationRules:
    """Разбирает раздел validation задания в ValidationRules.

    Зачем нужна: отдельно обрабатывает два особых случая — non_negative: all
    (все числовые столбцы сразу) и вложенный раздел revenue_check.
    Отсутствующие правила получают значения по умолчанию.

    Аргументы:
        raw: содержимое раздела validation (может быть пустым словарём).
        where: путь к разделу для сообщений об ошибках, например
            "jobs.orders.validation".

    Возвращает:
        ValidationRules.
    """
    non_negative = raw.get("non_negative", ())
    all_numeric = non_negative == "all"

    revenue = None
    if raw.get("revenue_check"):
        rc = _section(raw, "revenue_check", where)
        w = f"{where}.revenue_check"
        revenue = RevenueCheck(
            metric_column=_get(rc, "metric_column", w, str),
            group_by=_get(rc, "group_by", w, list),
            volume=_get(rc, "volume", w, str),
            price=_get(rc, "price", w, str),
            revenue=_get(rc, "revenue", w, str),
            divisor=_get(rc, "divisor", w, float, default=1000.0),
            tolerance_pct=_get(rc, "tolerance_pct", w, float, default=1.0),
        )

    return ValidationRules(
        required_columns=_get(raw, "required_columns", where, list, default=()),
        key_columns=_get(raw, "key_columns", where, list, default=()),
        non_negative_columns=() if all_numeric else _get(raw, "non_negative", where, list, default=()),
        all_numeric_non_negative=all_numeric,
        min_rows=_get(raw, "min_rows", where, int, default=1),
        max_row_drop_pct=_get(raw, "max_row_drop_pct", where, float, default=None),
        max_bad_values_pct=_get(raw, "max_bad_values_pct", where, float, default=5.0),
        revenue_check=revenue,
    )


def _section(d: dict, key: str, where: str = "") -> dict:
    """Достаёт из словаря вложенный раздел и проверяет, что это словарь.

    Зачем нужна: без неё опечатка в YAML (раздела нет или вместо него строка)
    дала бы непонятный AttributeError где-то дальше. Здесь же сразу будет
    сообщение «нет раздела 'jobs.orders.target'».

    Аргументы:
        d: словарь, в котором ищем.
        key: имя раздела.
        where: путь к словарю d — для текста ошибки.

    Возвращает:
        Найденный раздел (dict).

    Исключения:
        ConfigError: раздела нет или он не словарь.
    """
    full = f"{where}.{key}" if where else key
    value = d.get(key)
    if not isinstance(value, dict):
        raise ConfigError(f"В config.yaml нет раздела '{full}' (или он заполнен неверно).")
    return value


def _get(d: dict, key: str, where: str, kind: type, default=_REQUIRED):
    """Достаёт значение из словаря и проверяет его тип.

    Зачем нужна: единое место для всех проверок типов конфига, чтобы каждая
    ошибка выглядела одинаково: «'jobs.orders.validation.min_rows' должен быть
    типа int, а сейчас: 'много'». Тонкости:
        * для float принимается и целое число (tolerance_pct: 1);
        * списки должны состоять из строк и возвращаются кортежем — объекты
          конфига неизменяемые (frozen), и кортеж тоже нельзя изменить;
        * True/False не принимаются как int, хотя в Python bool — подкласс int.

    Аргументы:
        d: словарь, в котором ищем.
        key: имя ключа.
        where: путь к словарю — для текста ошибки.
        kind: ожидаемый тип: str, int, float, bool или list.
        default: значение, если ключа нет; если не передан — ключ обязателен.

    Возвращает:
        Значение нужного типа (список — кортежем) или default.

    Исключения:
        ConfigError: обязательного ключа нет или тип не тот.
    """
    full = f"{where}.{key}"
    value = d.get(key)
    if value is None:
        if default is _REQUIRED:
            raise ConfigError(f"В config.yaml не заполнен ключ '{full}'.")
        return default

    if kind is float and isinstance(value, int) and not isinstance(value, bool):
        return float(value)
    if kind is list:
        if not isinstance(value, list) or not all(isinstance(v, str) for v in value):
            raise ConfigError(f"'{full}' должен быть списком строк, например [Company, Product].")
        return tuple(value)
    # bool — подкласс int, поэтому `True` прошёл бы проверку на int без этой оговорки
    if not isinstance(value, kind) or (kind is int and isinstance(value, bool)):
        raise ConfigError(f"'{full}' должен быть типа {kind.__name__}, а сейчас: {value!r}.")
    return value
