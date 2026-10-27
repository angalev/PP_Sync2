"""Журнал применённых изменений (CSV, открывается Excel'ем как есть)."""

from __future__ import annotations

import csv
import datetime as _dt
import getpass
import os
from pathlib import Path

from .models import DiffRow

JOURNAL_NAME = "journal.csv"
DELIMITER = ";"
ENCODING = "utf-8-sig"  # чтобы Excel не ломал кириллицу

COLUMNS = [
    "Дата-время",
    "Пользователь",
    "Файл с изменениями",
    "Трансгаз",
    "XYZ",
    "Строка",
    "Столбец",
    "Поле",
    "Было",
    "Стало",
    "Результат",
]


def current_user() -> str:
    try:
        return os.environ.get("USERNAME") or getpass.getuser()
    except Exception:  # noqa: BLE001 — журнал не должен ронять применение
        return "неизвестен"


class Journal:
    """Дописывает строки в ``log\\journal.csv``."""

    def __init__(self, log_dir: str | Path, filename: str = JOURNAL_NAME) -> None:
        self.path = Path(log_dir) / filename
        self.user = current_user()

    def write(self, source_name: str, rows: list[tuple[DiffRow, str]]) -> Path:
        """Записать результаты применения.

        ``rows`` — пары ``(различие, результат)``, где результат это
        «применено», «не отмечено», «ошибка: …» и т.п.
        """
        self.path.parent.mkdir(parents=True, exist_ok=True)
        is_new = not self.path.exists() or self.path.stat().st_size == 0
        stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # noqa: DTZ005

        with self.path.open("a", newline="", encoding=ENCODING) as handle:
            writer = csv.writer(handle, delimiter=DELIMITER)
            if is_new:
                writer.writerow(COLUMNS)
            for diff_row, result in rows:
                writer.writerow(
                    [
                        stamp,
                        self.user,
                        source_name,
                        diff_row.transgaz,
                        diff_row.key,
                        diff_row.row or "",
                        diff_row.column,
                        diff_row.field_name,
                        diff_row.current_display,
                        diff_row.new_display,
                        result,
                    ]
                )
        return self.path

    def write_note(self, source_name: str, text: str) -> Path:
        """Служебная запись без привязки к строке (ошибка, отмена и т.п.)."""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        is_new = not self.path.exists() or self.path.stat().st_size == 0
        stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # noqa: DTZ005
        with self.path.open("a", newline="", encoding=ENCODING) as handle:
            writer = csv.writer(handle, delimiter=DELIMITER)
            if is_new:
                writer.writerow(COLUMNS)
            writer.writerow(
                [stamp, self.user, source_name, "", "", "", "", "", "", "", text]
            )
        return self.path
