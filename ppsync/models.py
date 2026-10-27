"""Модели данных: одна строка сравнения двух файлов."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any


class Status(str, Enum):
    """Статус найденного различия."""

    TO_APPLY = "к применению"  # в боевом пусто, в копии есть значение
    OVERWRITE = "перезапись"  # в боевом другое значение — требует внимания
    CLEAR = "очистка"  # в копии пусто, в боевом было значение
    FORMULA = "формула в боевом"  # трогать формулы нельзя
    SOURCE_FORMULA = "формула в копии"  # значение посчитано формулой копии
    KEY_NOT_FOUND = "строки нет в боевом"  # ключ есть в копии, но не в боевом


#: Что вообще можно применять (чекбокс активен).
#:
#: ``SOURCE_FORMULA`` здесь не по недосмотру: в боевой файл уходит **значение**,
#: посчитанное формулой копии, а не сама формула. Ячейка в боевом при этом
#: обычная — формулы в ней нет, ломать нечего.
APPLICABLE: frozenset[Status] = frozenset(
    {Status.TO_APPLY, Status.OVERWRITE, Status.CLEAR, Status.SOURCE_FORMULA}
)

#: Что показываем, но применить нельзя.
BLOCKED: frozenset[Status] = frozenset({Status.FORMULA, Status.KEY_NOT_FOUND})


@dataclass
class DiffRow:
    """Одно различие: ячейка в боевом файле против той же ячейки в копии."""

    key: str  # сводный код X/Y/Z
    column: str  # буква столбца в боевом файле
    field_name: str  # «[код] заголовок столбца» — для человека
    status: Status

    code: str = ""  # код столбца из строки кодов — по нему столбец и находится
    source_column: str = ""  # буква того же столбца в копии (может отличаться)
    row: int | None = None  # строка в боевом файле
    source_row: int | None = None  # строка в копии — чтобы найти глазами
    transgaz: str = ""
    object_name: str = ""

    current_value: Any = None  # что сейчас в боевом
    new_value: Any = None  # что в копии
    current_display: str = ""
    new_display: str = ""

    number_format: str = ""  # формат целевой ячейки
    text_format: bool = False  # формат «Текстовый»
    replaces_formula: bool = False  # в боевом стоит формула, и она исчезнет
    note: str = ""  # пояснение

    selected: bool = False

    @property
    def applicable(self) -> bool:
        return self.status in APPLICABLE

    @property
    def address(self) -> str:
        """Адрес целевой ячейки в боевом файле, например CF125."""
        return f"{self.column}{self.row}" if self.row else ""

    @property
    def source_address(self) -> str:
        """Адрес той же ячейки в файле с изменениями."""
        if not self.source_row:
            return ""
        return f"{self.source_column or self.column}{self.source_row}"
