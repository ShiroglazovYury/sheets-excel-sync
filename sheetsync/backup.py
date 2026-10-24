"""
Бэкапы: после каждого успешного синка сохраняем то, что ушло в Excel,
отдельным .xlsx в backups/. Если в Google что-то сломали или переписали
прогноз, всегда можно посмотреть, какие цифры стояли в модели неделю назад.

Полный экспорт гуглотаблицы через Drive не подходит: у Google жёсткий лимит
на размер экспорта («This file is too large to be exported»), и мы в него
уже упирались на рабочих данных. Поэтому бэкапим ровно тот срез, что синкнули.
"""

import datetime
from pathlib import Path

import openpyxl

from .table import Table


def save_backup(table: Table, job_name: str, backup_dir: Path, keep: int) -> Path:
    """Сохраняет записанные в Excel данные отдельным .xlsx-файлом в папку бэкапов.

    Зачем нужна: если источник в Google сломали или переписали прогноз, всегда
    можно открыть бэкап и посмотреть, какие цифры стояли в модели раньше.
    Бэкапится ровно тот срез, что ушёл в Excel, а не вся гуглотаблица: экспорт
    большой таблицы целиком Google не разрешает. Файл пишется потоково
    (write_only), поэтому 40 000 строк не загружаются в память целиком.

    Аргументы:
        table: очищенная таблица, которую только что записали в Excel.
        job_name: имя задания (forecast, orders) — попадает в имя файла и листа.
        backup_dir: папка для бэкапов (storage.backup_dir в конфиге).
        keep: сколько последних бэкапов этого задания хранить.

    Возвращает:
        Путь к созданному файлу вида backups/orders_2026-10-24_16-44-09.xlsx.
    """
    backup_dir.mkdir(parents=True, exist_ok=True)
    path = backup_dir / f"{job_name}_{_stamp()}.xlsx"

    # write_only — потоковая запись: 40 000 строк без загрузки всего в память
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(job_name[:31])  # Excel не разрешает имя листа длиннее 31 символа
    ws.append(table.header)
    for row in table.rows:
        ws.append(row)
    wb.save(path)

    rotate(backup_dir, f"{job_name}_*.xlsx", keep)
    return path


def rotate(directory: Path, pattern: str, keep: int):
    """Удаляет старые файлы, оставляя только `keep` самых свежих.

    Зачем нужна: без неё папки backups/ и reports/ росли бы с каждым синком
    бесконечно. Используется и для бэкапов, и для текстовых отчётов.

    Имена файлов содержат метку времени ГГГГ-ММ-ДД_ЧЧ-ММ-СС, поэтому
    сортировка по имени совпадает с сортировкой по времени создания.

    Аргументы:
        directory: папка, в которой чистим.
        pattern: шаблон имён, например "orders_*.xlsx", — чтобы не задеть файлы
            других заданий.
        keep: сколько файлов оставить; при keep <= 0 ничего не удаляется.
    """
    files = sorted(directory.glob(pattern))
    for old in files[:-keep] if keep > 0 else []:
        old.unlink(missing_ok=True)


def _stamp() -> str:
    """Текущее время строкой вида 2026-10-24_16-44-09 — для имён файлов.

    Зачем нужна: в имени файла нельзя использовать двоеточие, а такой формат
    ещё и сортируется по алфавиту в том же порядке, что и по времени (на этом
    построена rotate()).
    """
    return datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
