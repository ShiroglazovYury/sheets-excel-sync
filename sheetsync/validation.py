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
    issues: list[Issue] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def error(self, message, row=None, column=None):
        self.issues.append(Issue("error", message, row, column))

    def warning(self, message, row=None, column=None):
        self.issues.append(Issue("warning", message, row, column))

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.level == "warning"]

    @property
    def ok(self) -> bool:
        return not self.errors

    def summary(self, limit: int = 5) -> str:
        """Короткая версия для окна в Excel: первые `limit` проблем каждого уровня."""
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
    """Выручка в каждом месяце должна быть ≈ объём × цена / divisor.

    Строки объёма, цены и выручки одной комбинации (компания × продукт × …)
    находим по столбцам group_by. Сравниваем с допуском tolerance_pct, но не
    точнее 0.001: выручка в источнике округлена до 3 знаков.
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
