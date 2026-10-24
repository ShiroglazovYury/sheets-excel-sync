"""
Проверки данных перед записью в Excel.

Два уровня проблем:
  * error   — данные нельзя писать в модель: сломана структура (нет нужных
              столбцов, задвоились ключи), подозрительно мало строк, слишком
              много мусора. Excel не трогаем вообще;
  * warning — данные пишем, но человек должен об этом знать: отрицательные
              значения, выручка не сходится с объём × цена, отдельные
              нечитаемые ячейки.

Логика «лучше не записать ничего, чем записать криво»: управленческая
модель с наполовину пустым листом опаснее модели со вчерашними данными.
"""

from collections import defaultdict
from dataclasses import dataclass, field

from .cleaning import CleaningStats, Issue
from .config import RevenueCheck, ValidationRules
from .table import Table


@dataclass
class ValidationReport:
    """Отчёт проверки: все найденные проблемы и заметки об автоисправлениях.

    Зачем нужен: собрать в одном месте проблемы двух уровней — error (данные
    писать нельзя) и warning (писать можно, но человек должен знать) — и уметь
    показать их коротко (в окне Excel) и полностью (в файле отчёта).

    Поля:
        issues: список Issue (и из очистки, и из проверок).
        notes: фразы «Исправлено автоматически: …» из CleaningStats.
    """
    issues: list[Issue] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def error(self, message, row=None, column=None):
        """Добавляет блокирующую проблему: с ней данные в Excel не попадут.

        Аргументы:
            message: текст для человека.
            row: номер строки в Google Sheets (если проблема в конкретной строке).
            column: название столбца.
        """
        self.issues.append(Issue("error", message, row, column))

    def warning(self, message, row=None, column=None):
        """Добавляет предупреждение: данные запишутся, но проблема попадёт в отчёт.

        Аргументы:
            message: текст для человека.
            row: номер строки в Google Sheets (если проблема в конкретной строке).
            column: название столбца.
        """
        self.issues.append(Issue("warning", message, row, column))

    @property
    def errors(self) -> list[Issue]:
        """Только блокирующие проблемы (уровень error)."""
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self) -> list[Issue]:
        """Только предупреждения (уровень warning)."""
        return [i for i in self.issues if i.level == "warning"]

    @property
    def ok(self) -> bool:
        """True, если нет ни одной ошибки — значит, данные можно писать в Excel.

        Предупреждения на это не влияют.
        """
        return not self.errors

    def summary(self, limit: int = 5) -> str:
        """Короткая версия отчёта для окна Excel.

        Зачем нужна: в окне не поместятся сотни строк, поэтому показываются первые
        `limit` проблем каждого уровня и «… и ещё N»; полный список лежит в файле
        отчёта. В конце — строка об автоисправлениях.

        Аргументы:
            limit: сколько проблем каждого уровня показать.

        Возвращает:
            Текст или «Проблем не найдено.».
        """
        lines = []
        for title, items in (("Ошибки", self.errors), ("Предупреждения", self.warnings)):
            if not items:
                continue
            lines.append(f"{title} ({len(items)}):")
            lines += [f"  • {issue}" for issue in items[:limit]]
            if len(items) > limit:
                lines.append(f"  … и ещё {len(items) - limit}")
        if self.notes:
            lines.append("Исправлено автоматически: " + "; ".join(self.notes))
        return "\n".join(lines) if lines else "Проблем не найдено."

    def full_text(self) -> str:
        """Полный текст отчёта — все проблемы с пометкой [ERROR] или [WARNING].

        Зачем нужна: этот текст сохраняется в файл reports/*.txt, чтобы можно было
        пройтись по всем проблемам и исправить источник.
        """
        lines = [f"[{i.level.upper()}] {i}" for i in self.issues]
        if self.notes:
            lines.append("Исправлено автоматически: " + "; ".join(self.notes))
        return "\n".join(lines)


def validate(
    table: Table,
    rules: ValidationRules,
    *,
    numeric_idx: set[int],
    stats: CleaningStats,
    previous_rows: int | None = None,
    force: bool = False,
) -> ValidationReport:
    """Проверяет очищенные данные по правилам задания.

    Зачем нужна: это шаг «Проверки» и главная «своя логика» проекта: лучше не
    записать ничего, чем записать криво. По порядку:
        1. заголовок: нет обязательного или ключевого столбца, дубль заголовка —
           error; дальше при этом не проверяем, иначе отчёт зашумят следствия;
        2. строк меньше min_rows — error;
        3. строк стало меньше, чем в прошлый раз, больше чем на max_row_drop_pct —
           error (при force — warning);
        4. доля нечитаемых чисел больше max_bad_values_pct — error: похоже,
           столбцы съехали;
        5. пустые или повторяющиеся ключи строк — error;
        6. отрицательные значения — warning;
        7. выручка не сходится с объём × цена — warning.

    Аргументы:
        table: очищенная таблица.
        rules: правила задания из конфига.
        numeric_idx: номера числовых столбцов.
        stats: статистика очистки (оттуда доля мусора и заметки).
        previous_rows: сколько строк было в прошлый успешный синк (или None).
        force: ослабить проверку падения числа строк до предупреждения.

    Возвращает:
        ValidationReport; report.ok — можно ли писать в Excel.
    """
    report = ValidationReport(notes=stats.notes())

    _check_header(table, rules, report)
    if not report.ok:
        # без нужных столбцов остальные проверки бессмысленны и только зашумят отчёт
        return report

    if table.n_rows < rules.min_rows:
        report.error(
            f"в источнике {table.n_rows} строк, а ожидается не меньше {rules.min_rows}. "
            "Похоже, диапазон указан неверно или данные удалили."
        )

    if previous_rows and rules.max_row_drop_pct is not None:
        drop_pct = (previous_rows - table.n_rows) / previous_rows * 100
        if drop_pct > rules.max_row_drop_pct:
            message = (
                f"строк стало {table.n_rows} вместо {previous_rows} в прошлый раз "
                f"(−{drop_pct:.0f}%, допустимо −{rules.max_row_drop_pct:.0f}%)."
            )
            if force:
                report.warning(message + " Записано принудительно (--force).")
            else:
                report.error(message + " Если так и должно быть — запустите синк с --force.")

    if stats.numeric_cells:
        bad_pct = stats.bad_values / stats.numeric_cells * 100
        if bad_pct > rules.max_bad_values_pct:
            report.error(
                f"{bad_pct:.1f}% числовых ячеек не читаются как числа "
                f"(допустимо {rules.max_bad_values_pct:g}%). Похоже, столбцы съехали."
            )

    _check_keys(table, rules.key_columns, report)
    _check_non_negative(table, rules, numeric_idx, report)
    if rules.revenue_check:
        _check_revenue(table, rules.revenue_check, numeric_idx, report)
    return report


def _check_header(table: Table, rules: ValidationRules, report: ValidationReport):
    """Проверяет заголовок: все нужные столбцы на месте и нет дублей.

    Зачем нужна: если столбец переименовали или удалили, данные уже нельзя
    писать в модель. Если такой столбец есть, но с другим регистром ('Company'
    вместо 'company'), в сообщении будет подсказка «есть похожий». Ключевые
    столбцы проверяются, даже если их забыли перечислить в required_columns.

    Аргументы:
        table: таблица.
        rules: правила (required_columns и key_columns).
        report: отчёт, в который добавляются ошибки.
    """
    present = set(table.header)
    by_lower = {h.lower(): h for h in table.header}
    # ключевые столбцы обязательны, даже если их забыли перечислить в required_columns
    needed = dict.fromkeys(rules.required_columns + rules.key_columns)
    for name in needed:
        if name in present:
            continue
        hint = f" Есть похожий: '{by_lower[name.lower()]}' — переименуйте." if name.lower() in by_lower else ""
        report.error(f"нет обязательного столбца '{name}'.{hint}")

    seen = set()
    for name in table.header:
        if name and name in seen:
            report.error(f"столбец '{name}' встречается в заголовке дважды.")
        seen.add(name)


def _check_keys(table: Table, key_columns: tuple[str, ...], report: ValidationReport):
    """Проверяет ключи строк: они не пустые и не повторяются.

    Зачем нужна: по ключу (например order_id) строки сопоставляются при
    сравнении с прошлым синком. Дубль ключа почти всегда означает ошибку в
    источнике (строку скопировали), а сопоставление по такому ключу было бы
    неоднозначным. В сообщении — номера строк с дублем (до 10).

    Аргументы:
        table: таблица.
        key_columns: столбцы-ключ; пусто — проверка не выполняется.
        report: отчёт.
    """
    if not key_columns:
        return
    idx = [table.column_index(c) for c in key_columns]
    rows_by_key = defaultdict(list)
    for row, number in zip(table.rows, table.row_numbers):
        key = tuple(row[i] for i in idx)
        if all(v in ("", None) for v in key):
            report.error(f"пустой ключ ({', '.join(key_columns)}).", row=number)
            continue
        rows_by_key[key].append(number)

    for key, numbers in rows_by_key.items():
        if len(numbers) > 1:
            shown = ", ".join(map(str, numbers[:10]))
            report.error(f"ключ {' / '.join(map(str, key))} повторяется в строках {shown}.")


def _check_non_negative(table, rules: ValidationRules, numeric_idx: set[int], report):
    """Предупреждает об отрицательных значениях в столбцах, где их быть не должно.

    Зачем нужна: отрицательная цена или объём в прогнозе — почти всегда
    опечатка. Но это не повод блокировать запись (в заказах, например, объём
    бывает отрицательным у возвратов), поэтому только предупреждение.

    Аргументы:
        table: таблица.
        rules: правила (all_numeric_non_negative или список столбцов).
        numeric_idx: номера числовых столбцов (для варианта «все числовые»).
        report: отчёт.
    """
    if rules.all_numeric_non_negative:
        columns = sorted(numeric_idx)
    else:
        columns = [table.column_index(c) for c in rules.non_negative_columns if c in table.header]
    for row, number in zip(table.rows, table.row_numbers):
        for c in columns:
            value = row[c]
            if isinstance(value, (int, float)) and value < 0:
                report.warning(f"отрицательное значение {value}.", row=number, column=table.header[c])


def _check_revenue(table: Table, rc: RevenueCheck, numeric_idx: set[int], report: ValidationReport):
    """Проверяет, что выручка в каждом месяце ≈ объём × цена / divisor.

    Зачем нужна: бизнес-правило прогноза. Если выручку «поправили руками» и она
    перестала сходиться с объёмом и ценой, правило покажет комбинацию, месяц и
    расхождение в процентах.

    Строки объёма, цены и выручки одной комбинации (компания × продукт × …)
    находятся по столбцам group_by. Сравнение — с допуском tolerance_pct, но не
    точнее 0.001: выручка в источнике округлена до 3 знаков. Если в каком-то
    месяце одно из трёх значений пустое, этот месяц пропускается.

    Аргументы:
        table: таблица.
        rc: настройки правила из конфига.
        numeric_idx: номера столбцов-месяцев.
        report: отчёт; нехватка нужных столбцов — error, расхождения — warning.
    """
    missing = [c for c in (rc.metric_column, *rc.group_by) if c not in table.header]
    if missing:
        report.error(f"для проверки выручки нет столбцов: {', '.join(missing)}.")
        return
    metric_i = table.column_index(rc.metric_column)
    group_i = [table.column_index(c) for c in rc.group_by]

    groups = defaultdict(dict)  # группа → {метрика: (строка, номер)}
    for row, number in zip(table.rows, table.row_numbers):
        groups[tuple(row[i] for i in group_i)][row[metric_i]] = (row, number)

    for key, by_metric in groups.items():
        if not all(m in by_metric for m in (rc.volume, rc.price, rc.revenue)):
            continue
        volume, price = by_metric[rc.volume][0], by_metric[rc.price][0]
        revenue, revenue_row = by_metric[rc.revenue]
        for c in sorted(numeric_idx):
            v, p, r = volume[c], price[c], revenue[c]
            if not all(isinstance(x, (int, float)) for x in (v, p, r)):
                continue
            expected = v * p / rc.divisor
            if abs(r - expected) > max(abs(expected) * rc.tolerance_pct / 100, 0.001):
                gap = (r - expected) / expected * 100 if expected else float("inf")
                report.warning(
                    f"{' / '.join(map(str, key))}: выручка {r:,.3f}, а объём × цена = {expected:,.3f} "
                    f"(расхождение {gap:+.0f}%).",
                    row=revenue_row, column=table.header[c],
                )
