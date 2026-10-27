"""Один скрытый экземпляр Excel: запустить, открыть книгу, закрыть.

Здесь собрано всё, чем программа защищается от «Excel молча повис»: макросы
и события выключены, связи не обновляются, на модальное окно «Конфликт имен»
отвечает фоновый сторож, а отказ ``Workbooks.Open`` превращается в сообщение,
из которого понятно, какое окно его вызвало.

Экземпляр привязан к потоку, который его создал: COM-объекты Excel нельзя
трогать из другого потока. Поэтому при сравнении каждый из двух файлов
читается своим экземпляром в своём потоке — параллельно.
"""

from __future__ import annotations

from collections.abc import Callable
from contextlib import suppress
from pathlib import Path
from typing import Any

from .config import Config
from .excel_guard import DialogWatcher, describe, find_dialogs
from .workbook_reader import WorkbookStructureError

Progress = Callable[[str], None]


class ExcelError(Exception):
    """Excel не открыл книгу. ``problems`` — что удалось узнать о причине."""

    def __init__(self, message: str, problems: list[str] | None = None) -> None:
        super().__init__(message)
        self.problems = problems or []


class ExcelInstance:
    """Скрытый Excel с одной или несколькими открытыми книгами."""

    def __init__(self, config: Config, progress: Progress | None = None) -> None:
        self.config = config
        self.progress = progress or (lambda _message: None)
        self.app: Any = None
        self.pid: int = 0  # для сторожа модальных окон (excel_guard)
        self.watcher: DialogWatcher | None = None
        self.books: list[Any] = []

    def start(self) -> None:
        import xlwings as xw  # локальный импорт: COM нужен только здесь

        self.app = xw.App(visible=False, add_book=False)
        with suppress(Exception):
            self.pid = int(self.app.pid)
        if self.config.auto_answer_name_conflict and self.pid:
            # Пока висит модальное окно, этот поток стоит внутри COM-вызова,
            # поэтому отвечать в окне может только отдельный поток.
            self.watcher = DialogWatcher(self.pid)
            self.watcher.start()
        self.app.display_alerts = False
        self.app.screen_updating = False

        api = self.app.api
        if self.config.disable_macros:
            # msoAutomationSecurityForceDisable: макросы книги не запускаются.
            # Без этого Workbook_Open выполняется и может вывалить модальную
            # ошибку VBA («Run-time error… End»), которую display_alerts не
            # гасит — невидимый Excel повиснет до ручного нажатия кнопки.
            # Сам vbaProject.bin при этом остаётся в файле нетронутым.
            with suppress(Exception):
                api.AutomationSecurity = 3
        if self.config.disable_events:
            # Обработчики Worksheet_Change не должны реагировать на нашу запись
            with suppress(Exception):
                api.EnableEvents = False
        with suppress(Exception):
            api.AskToUpdateLinks = False

    def open_sheet(
        self, path: Path, shown: Path | None = None, *, read_only: bool = True
    ) -> Any:
        """Открыть книгу и вернуть лист ПП (xlwings Sheet).

        ``shown`` — имя, под которым файл знает человек: рабочая копия лежит
        во временной папке, и её путь в сообщении об ошибке ничего не скажет.
        """
        name = (shown or path).name
        try:
            book = self.app.books.open(
                str(path), update_links=False, read_only=read_only
            )
        except Exception as exc:
            # Самая частая причина отказа Open — модальное окно Excel, которое
            # висит прямо сейчас: «Конфликт имен», вход в учётную запись,
            # ошибка макроса. Без этой строки человек видит только невнятное
            # «Метод Open из класса Workbooks завершен неверно».
            problems = [str(exc)]
            problems.extend(self.drain())
            dialogs = find_dialogs(self.pid)
            if dialogs:
                problems.append(describe(dialogs))
            else:
                problems.append(
                    "Если перед этим на экране было окно Excel «Конфликт имен» — "
                    "ответьте в нём (любое новое имя и ОК) и повторите: "
                    "пока окно висит, открыть книгу нельзя."
                )
            raise ExcelError(f"Не удалось открыть файл {name}", problems) from exc
        self.books.append(book)
        self.drain()

        if not read_only:
            # До любого save(): иначе Excel покажет модальный диалог разрешения
            # конфликтов и невидимый экземпляр повиснет навсегда.
            try:
                book.api.ConflictResolution = self.config.conflict_resolution
            except Exception as exc:  # noqa: BLE001 — книга не в общем доступе
                self.progress(f"ConflictResolution не установлен: {exc}")

        names = [sheet.name for sheet in book.sheets]
        if self.config.sheet_name not in names:
            raise WorkbookStructureError(
                f"В файле {name} нет листа «{self.config.sheet_name}»",
                [f"Есть листы: {', '.join(names)}"],
            )
        return book.sheets[self.config.sheet_name]

    def drain(self) -> list[str]:
        """Показать, что сторож окон сделал, пока мы ждали Excel."""
        if self.watcher is None:
            return []
        notes = self.watcher.drain()
        for note in notes:
            self.progress(note)
        return notes

    def close(self) -> None:
        """Закрыть книги и погасить Excel. Безопасно вызывать дважды."""
        # Здесь глушим всё: закрытие не должно маскировать исходную ошибку,
        # а COM после сбоя нередко бросает на каждом шаге.
        if self.watcher is not None:
            self.watcher.stop()
            self.drain()
            self.watcher = None
        for book in self.books:
            try:
                book.close()
            except Exception:  # noqa: BLE001, S110
                pass
        self.books = []
        if self.app is not None:
            try:
                self.app.quit()
            except Exception:  # noqa: BLE001, S110
                pass
            try:
                self.app.kill()  # добить процесс, если quit() не сработал
            except Exception:  # noqa: BLE001, S110
                pass
            self.app = None
        self.pid = 0
