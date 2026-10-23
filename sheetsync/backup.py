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
    """Оставляет `keep` самых свежих файлов по шаблону, остальные удаляет.

    Имена содержат метку времени ГГГГ-ММ-ДД_ЧЧ-ММ-СС, поэтому сортировка по
    имени — это сортировка по времени.
    """
    files = sorted(directory.glob(pattern))
    for old in files[:-keep] if keep > 0 else []:
        old.unlink(missing_ok=True)


def _stamp() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
