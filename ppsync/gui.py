"""Интерфейс (PyQt6). Один экран: два файла -> таблица различий -> применение.

Работа с Excel вынесена в отдельный поток :class:`ExcelWorker`. Причина не в
скорости, а в COM: объекты Excel привязаны к потоку, который их создал, поэтому
и открытие книги, и запись должны идти из одного и того же потока. Заодно окно
не подвисает, пока Excel открывает файл по сети.
"""

from __future__ import annotations

import os
import queue
import signal
import sys
from contextlib import suppress
from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QAction, QBrush, QColor, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .applier import ApplyError, ApplyResult, CompareResult, SyncSession
from .config import CONFLICT_MODES, Config, DataDirError, probe_dir
from .differ import group_by_transgaz
from .excel_guard import describe, find_dialogs
from .models import BLOCKED, DiffRow, Status
from .workbook_reader import WorkbookStructureError

WORKBOOK_FILTER = "Книги Excel (*.xlsm *.xlsx)"

# Цвета строк таблицы по статусу
_ROW_COLORS = {
    Status.OVERWRITE: QColor("#fff1c9"),  # требует внимания
    Status.CLEAR: QColor("#ffe0c2"),  # стираем чужое — тем более
    Status.FORMULA: QColor("#e6e6e6"),
    Status.SOURCE_FORMULA: QColor("#dff0e4"),  # применить можно, но стоит взглянуть
    Status.KEY_NOT_FOUND: QColor("#ffd9d9"),
}

(
    COL_CHECK,
    COL_TRANSGAZ,
    COL_OBJECT,
    COL_KEY,
    COL_FIELD,
    COL_CELL,
    COL_WAS,
    COL_WILL,
    COL_STATUS,
    COL_NOTE,
) = range(10)
_HEADERS = [
    "",
    "Трансгаз",
    "Объект",
    "XYZ",
    "Столбец",
    "Ячейка",
    "В боевом",
    "В копии",
    "Статус",
    "Пояснение",
]

#: Ширина столбцов при первом показе, в пикселях. Дальше её тянет мышью
#: пользователь, поэтому это именно начальные значения, а не жёсткие рамки.
_COLUMN_WIDTHS = {
    COL_CHECK: 34,
    COL_TRANSGAZ: 140,
    COL_OBJECT: 230,
    COL_KEY: 125,
    COL_FIELD: 145,
    COL_CELL: 75,
    COL_WAS: 125,
    COL_WILL: 125,
    COL_STATUS: 120,
}


def _plural(count: int, one: str, few: str, many: str) -> str:
    """«1 изменение», «2 изменения», «358 изменений».

    Счётчик отмеченного человек читает перед записью в боевой файл, поэтому
    он должен звучать как фраза, а не как «1 изменений».
    """
    tail = count % 100
    if 11 <= tail <= 14:
        return many
    tail %= 10
    if tail == 1:
        return one
    if 2 <= tail <= 4:
        return few
    return many


def _changes(count: int) -> str:
    return f"{count} {_plural(count, 'изменение', 'изменения', 'изменений')}"


class ExcelWorker(QThread):
    """Поток-владелец сессии Excel. Команды приходят через очередь.

    Excel открыт только пока идёт сравнение или запись; между ними сессия —
    это просто найденные различия в памяти.
    """

    progress = pyqtSignal(str)
    compared = pyqtSignal(object)  # CompareResult
    applied = pyqtSignal(object)  # ApplyResult
    failed = pyqtSignal(str, str)  # заголовок, текст

    def __init__(self, config: Config) -> None:
        super().__init__()
        self.config = config
        self._commands: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.session: SyncSession | None = None

    # --- команды из GUI -----------------------------------------------------

    def compare(self) -> None:
        self._commands.put(("compare", None))

    def apply(self, rows: list[DiffRow]) -> None:
        self._commands.put(("apply", rows))

    def stop(self) -> None:
        self._commands.put(("stop", None))

    # --- сам поток ----------------------------------------------------------

    def run(self) -> None:
        import pythoncom  # из pywin32, ставится вместе с xlwings

        pythoncom.CoInitialize()
        try:
            while True:
                command, payload = self._commands.get()
                if command == "stop":
                    return
                if command == "compare":
                    self._compare()
                elif command == "apply":
                    self._apply(payload)
        finally:
            pythoncom.CoUninitialize()

    def _compare(self) -> None:
        session = SyncSession(self.config, progress=self.progress.emit)
        # Ссылку ставим до открытия книг: пока compare блокируется на COM,
        # сторож в GUI должен уже знать pid Excel, чтобы заметить модальное окно.
        self.session = session
        try:
            result = session.compare()
        except (ApplyError, WorkbookStructureError) as exc:
            self.session = None
            self.failed.emit("Не удалось сравнить файлы", exc.full_text())
            return
        except Exception as exc:  # noqa: BLE001 — любую ошибку показываем, а не роняем окно
            self.session = None
            self.failed.emit(
                "Ошибка при работе с Excel", f"{type(exc).__name__}: {exc}"
            )
            return
        self.compared.emit(result)

    def _apply(self, rows: list[DiffRow]) -> None:
        if self.session is None:
            self.failed.emit("Нечего применять", "Сначала выполните сравнение.")
            return
        try:
            result = self.session.apply(rows)
        except (ApplyError, WorkbookStructureError) as exc:
            self.failed.emit("Не удалось применить", exc.full_text())
            self.session = None
            return
        except Exception as exc:  # noqa: BLE001
            self.failed.emit("Ошибка при записи", f"{type(exc).__name__}: {exc}")
            self.session = None
            return
        self.session = None
        self.applied.emit(result)


class SettingsDialog(QDialog):
    """Редактор config.json — чтобы ничего не хардкодить."""

    def __init__(self, config: Config, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.setMinimumWidth(660)
        self.config = config

        self.sheet = QLineEdit(config.sheet_name)
        self.header_top = _spin(config.header_row_top)
        self.header_bottom = _spin(config.header_row_bottom)
        self.code_row = _spin(config.code_row)
        self.code_row.setToolTip(
            "Строка листа, где у каждого столбца стоит свой код. По этим кодам "
            "программа находит столбцы — буквы для неё ничего не значат, "
            "поэтому столбцы можно удалять и вставлять."
        )
        self.first_row = _spin(config.first_data_row)
        self.key_code = QLineEdit(config.key_code)
        self.transgaz_code = QLineEdit(config.transgaz_code)
        self.object_code = QLineEdit(config.object_code)

        self.compare_codes = QLineEdit(config.compare_codes)
        self.compare_codes.setPlaceholderText(
            "пусто — все столбцы; коды из строки кодов, пример: 259-273, 316-320"
        )
        self.ignore_codes = QLineEdit(config.ignore_codes)
        self.ignore_codes.setPlaceholderText("коды столбцов, которые не сравнивать")

        self.formula_codes = QLineEdit(config.formula_codes)
        self.formula_codes.setPlaceholderText(
            "коды расчётных столбцов: формулу в них не заменяем значением"
        )
        self.formula_codes.setToolTip(
            "Столбцы, где формула держит лист: суммы, проценты, подстановки "
            "из парного столбца КОРР. Формулу в них программа не трогает. "
            "Во всех остальных столбцах формула в боевом файле считается "
            "ленивой ссылкой вместо цифры (=DS515) и заменяется значением "
            "из копии."
        )

        self.data_dir = QLineEdit(config.data_dir)
        self.backup_keep = _spin(config.backup_keep, minimum=0, maximum=999)

        self.conflict_box = QComboBox()
        for mode, title in CONFLICT_MODES.items():
            self.conflict_box.addItem(title, mode)
        index = self.conflict_box.findData(config.conflict_resolution)
        self.conflict_box.setCurrentIndex(max(index, 0))

        self.verify = QCheckBox("Перечитывать ячейки после сохранения и сверять")
        self.verify.setChecked(config.verify_after_save)
        self.disable_macros = QCheckBox(
            "Не запускать макросы книги (иначе ошибка VBA повесит скрытый Excel)"
        )
        self.disable_macros.setChecked(config.disable_macros)
        self.disable_events = QCheckBox("Не запускать обработчики событий листа")
        self.disable_events.setChecked(config.disable_events)
        self.fix_names = QCheckBox(
            "Чинить служебные имена до открытия (убирает окно «Конфликт имен»)"
        )
        self.fix_names.setChecked(config.fix_service_names)
        self.auto_answer = QCheckBox(
            "Отвечать за меня в окне «Конфликт имен» про служебное имя"
        )
        self.auto_answer.setChecked(config.auto_answer_name_conflict)

        data_row = QHBoxLayout()
        data_row.addWidget(self.data_dir)
        browse_dir = QPushButton("Обзор…")
        browse_dir.clicked.connect(self._browse_data_dir)
        data_row.addWidget(browse_dir)

        form = QFormLayout()
        form.addRow("Лист:", self.sheet)
        form.addRow("Строка заголовков (верх):", self.header_top)
        form.addRow("Строка заголовков (низ):", self.header_bottom)
        form.addRow("Строка кодов столбцов:", self.code_row)
        form.addRow("Первая строка данных:", self.first_row)
        form.addRow("Код столбца ключа X/Y/Z:", self.key_code)
        form.addRow("Код столбца трансгаза:", self.transgaz_code)
        form.addRow("Код столбца объекта:", self.object_code)
        form.addRow("Сравнивать столбцы (коды):", self.compare_codes)
        form.addRow("Не сравнивать столбцы (коды):", self.ignore_codes)
        form.addRow("Столбцы с формулами (коды):", self.formula_codes)
        form.addRow("Папка бэкапов и журнала:", data_row)
        form.addRow("Хранить бэкапов (0 — все):", self.backup_keep)
        form.addRow("Разрешение конфликтов:", self.conflict_box)
        form.addRow("", self.verify)
        form.addRow("", self.disable_macros)
        form.addRow("", self.disable_events)
        form.addRow("", self.fix_names)
        form.addRow("", self.auto_answer)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(QLabel(f"config.json: {config.path}"))
        layout.addWidget(buttons)

    def _browse_data_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Папка бэкапов и журнала", self.data_dir.text()
        )
        if path:
            self.data_dir.setText(path)

    def _accept(self) -> None:
        config = self.config
        config.sheet_name = self.sheet.text().strip()
        config.header_row_top = self.header_top.value()
        config.header_row_bottom = self.header_bottom.value()
        config.code_row = self.code_row.value()
        config.first_data_row = self.first_row.value()
        config.key_code = self.key_code.text()
        config.transgaz_code = self.transgaz_code.text()
        config.object_code = self.object_code.text()
        config.compare_codes = self.compare_codes.text().strip()
        config.ignore_codes = self.ignore_codes.text().strip()
        config.formula_codes = self.formula_codes.text().strip()
        config.data_dir = self.data_dir.text().strip()
        config.backup_keep = self.backup_keep.value()
        config.conflict_resolution = int(self.conflict_box.currentData())
        config.verify_after_save = self.verify.isChecked()
        config.disable_macros = self.disable_macros.isChecked()
        config.disable_events = self.disable_events.isChecked()
        config.fix_service_names = self.fix_names.isChecked()
        config.auto_answer_name_conflict = self.auto_answer.isChecked()
        config.normalize()

        # Пути к файлам проверяются на главном экране, здесь их нет
        problems = [
            p
            for p in config.validate()
            if "файл" not in p.lower() or "первая строка" in p.lower()
        ]
        if problems:
            QMessageBox.warning(
                self, "Проверьте настройки", "\n".join(f"• {p}" for p in problems)
            )
            return
        config.save()
        self.accept()


def _spin(value: int, minimum: int = 1, maximum: int = 10000) -> QSpinBox:
    spin = QSpinBox()
    spin.setRange(minimum, maximum)
    spin.setValue(value)
    return spin


class MainWindow(QMainWindow):
    """Единственное окно программы."""

    def __init__(self, config: Config) -> None:
        super().__init__()
        self.config = config
        self.rows: list[DiffRow] = []
        self.compare_result: CompareResult | None = None
        self.busy = False
        self._dialog_reported = False

        self.setWindowTitle("Синхронизация Производственной программы")
        self.resize(1400, 820)

        self.worker = ExcelWorker(config)
        self.worker.progress.connect(self._on_progress)
        self.worker.compared.connect(self._on_compared)
        self.worker.applied.connect(self._on_applied)
        self.worker.failed.connect(self._on_failed)
        self.worker.start()

        # Сторож модальных окон Excel: без него зависший диалог выглядит как
        # «программа задумалась», и ждать можно бесконечно.
        self.dialog_timer = QTimer(self)
        self.dialog_timer.setInterval(2500)
        self.dialog_timer.timeout.connect(self._check_excel_dialogs)

        self._build_ui()

    # --- построение интерфейса ---------------------------------------------

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._build_top())

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._build_left())
        splitter.addWidget(self._build_table())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([300, 1100])
        layout.addWidget(splitter, stretch=1)

        layout.addWidget(self._build_bottom())

        menu = self.menuBar().addMenu("Файл")
        settings_action = QAction("Настройки…", self)
        settings_action.triggered.connect(self._open_settings)
        menu.addAction(settings_action)
        open_log = QAction("Открыть папку журнала", self)
        open_log.triggered.connect(lambda: self._open_folder(self.config.log_dir))
        menu.addAction(open_log)
        open_backup = QAction("Открыть папку бэкапов", self)
        open_backup.triggered.connect(lambda: self._open_folder(self.config.backup_dir))
        menu.addAction(open_backup)

    def _build_top(self) -> QWidget:
        box = QGroupBox("Файлы")
        layout = QVBoxLayout(box)

        self.source_edit = QLineEdit(self.config.source_file)
        self.target_edit = QLineEdit(self.config.target_file)

        layout.addLayout(
            self._file_row(
                "Файл с изменениями (ваша копия):",
                self.source_edit,
                self._choose_source,
            )
        )
        layout.addLayout(
            self._file_row(
                "Боевой файл на сервере:", self.target_edit, self._choose_target
            )
        )

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        settings_button = QPushButton("Настройки…")
        settings_button.clicked.connect(self._open_settings)
        buttons.addWidget(settings_button)
        self.compare_button = QPushButton("Сравнить")
        self.compare_button.setDefault(True)
        self.compare_button.setMinimumWidth(180)
        self.compare_button.clicked.connect(self._start_compare)
        buttons.addWidget(self.compare_button)
        layout.addLayout(buttons)

        return box

    def _file_row(self, title: str, edit: QLineEdit, handler: Any) -> QHBoxLayout:
        row = QHBoxLayout()
        label = QLabel(title)
        label.setMinimumWidth(210)
        row.addWidget(label)
        row.addWidget(edit, stretch=1)
        button = QPushButton("Выбрать…")
        button.clicked.connect(handler)
        row.addWidget(button)
        return row

    def _build_left(self) -> QWidget:
        box = QGroupBox("Трансгазы")
        layout = QVBoxLayout(box)

        self.transgaz_list = QListWidget()
        self.transgaz_list.itemChanged.connect(lambda _item: self._fill_table())
        layout.addWidget(self.transgaz_list, stretch=1)

        buttons = QHBoxLayout()
        select_all = QPushButton("Выбрать все")
        select_all.clicked.connect(lambda: self._set_all_transgazes(True))
        clear_all = QPushButton("Снять все")
        clear_all.clicked.connect(lambda: self._set_all_transgazes(False))
        buttons.addWidget(select_all)
        buttons.addWidget(clear_all)
        layout.addLayout(buttons)

        self.hide_blocked = QCheckBox("Скрывать то, что применить нельзя")
        self.hide_blocked.stateChanged.connect(lambda _state: self._fill_table())
        layout.addWidget(self.hide_blocked)

        return box

    def _build_table(self) -> QWidget:
        box = QGroupBox("Различия")
        layout = QVBoxLayout(box)

        self.table = QTableWidget(0, len(_HEADERS))
        self.table.setHorizontalHeaderLabels(_HEADERS)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.itemChanged.connect(self._on_table_item_changed)
        header = self.table.horizontalHeader()
        # Interactive для всех столбцов: только в этом режиме границу можно
        # тянуть мышью. Раньше стояли Stretch и ResizeToContents — Qt считал
        # ширину сам и мышь игнорировал. Двойной щелчок по границе по-прежнему
        # подгоняет столбец по содержимому.
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(COL_CHECK, QHeaderView.ResizeMode.Fixed)
        for column, width in _COLUMN_WIDTHS.items():
            self.table.setColumnWidth(column, width)
        # Последний столбец («Пояснение») забирает остаток ширины окна, но
        # тянуться руками при этом не перестаёт.
        header.setStretchLastSection(True)
        header.setMinimumSectionSize(30)
        layout.addWidget(self.table)

        buttons = QHBoxLayout()
        check_all = QPushButton("Отметить все в таблице")
        check_all.clicked.connect(lambda: self._set_all_rows(True))
        uncheck_all = QPushButton("Снять все в таблице")
        uncheck_all.clicked.connect(lambda: self._set_all_rows(False))
        buttons.addWidget(check_all)
        buttons.addWidget(uncheck_all)
        buttons.addStretch(1)
        self.status_summary = QLabel()
        buttons.addWidget(self.status_summary)
        layout.addLayout(buttons)

        return box

    def _build_bottom(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        row = QHBoxLayout()
        self.selected_label = QLabel("отмечено 0 изменений")
        font = QFont()
        font.setBold(True)
        self.selected_label.setFont(font)
        row.addWidget(self.selected_label)
        row.addStretch(1)

        self.apply_button = QPushButton("Применить отмеченное")
        self.apply_button.setEnabled(False)
        self.apply_button.setMinimumWidth(220)
        self.apply_button.clicked.connect(self._start_apply)
        row.addWidget(self.apply_button)
        layout.addLayout(row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumHeight(130)
        layout.addWidget(self.log_view)

        return widget

    # --- выбор файлов -------------------------------------------------------

    def _choose_source(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Файл с изменениями", self.source_edit.text(), WORKBOOK_FILTER
        )
        if path:
            self.source_edit.setText(path)

    def _choose_target(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Боевой файл", self.target_edit.text(), WORKBOOK_FILTER
        )
        if path:
            self.target_edit.setText(path)

    # --- сравнение ----------------------------------------------------------

    def _start_compare(self) -> None:
        if self.busy:
            return
        self.config.source_file = self.source_edit.text().strip()
        self.config.target_file = self.target_edit.text().strip()

        problems = self.config.validate()
        if problems:
            QMessageBox.warning(
                self, "Проверьте настройки", "\n".join(f"• {p}" for p in problems)
            )
            return
        self.config.save()

        self._set_busy(True)
        self.log_view.clear()
        self._log(f"Копия  : {self.config.source_file}")
        self._log(f"Боевой : {self.config.target_file}")
        self.worker.compare()

    def _on_compared(self, result: CompareResult) -> None:
        self._set_busy(False)
        self.compare_result = result
        self.rows = result.rows
        self._log(
            f"В боевом файле {result.target.count} строк, в копии "
            f"{result.source.count}; сравнивалось столбцов: {len(result.codes)}"
        )
        self._log(f"Найдено различий: {len(result.rows)}")
        for warning in result.warnings:
            self._log(f"! {warning}")
        if not result.rows:
            QMessageBox.information(
                self,
                "Различий нет",
                "Файлы совпадают по сравниваемым столбцам — применять нечего.",
            )

        self._fill_transgazes()
        self._fill_table()
        self._update_selected_label()

    def _fill_transgazes(self) -> None:
        self.transgaz_list.blockSignals(True)
        self.transgaz_list.clear()
        for name, rows in group_by_transgaz(self.rows).items():
            applicable = sum(1 for r in rows if r.applicable)
            item = QListWidgetItem(
                f"{name} — {len(rows)} (можно применить {applicable})"
            )
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            item.setData(Qt.ItemDataRole.UserRole, name)
            self.transgaz_list.addItem(item)
        self.transgaz_list.blockSignals(False)

    def _checked_transgazes(self) -> set[str]:
        names = set()
        for index in range(self.transgaz_list.count()):
            item = self.transgaz_list.item(index)
            if item.checkState() == Qt.CheckState.Checked:
                names.add(item.data(Qt.ItemDataRole.UserRole))
        return names

    def _visible_rows(self) -> list[DiffRow]:
        names = self._checked_transgazes()
        rows = [r for r in self.rows if (r.transgaz or "(без трансгаза)") in names]
        if self.hide_blocked.isChecked():
            rows = [r for r in rows if r.applicable]
        return rows

    def _fill_table(self) -> None:
        rows = self._visible_rows()
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        for position, row in enumerate(rows):
            check_item = QTableWidgetItem()
            check_item.setData(Qt.ItemDataRole.UserRole, row)
            if row.applicable:
                check_item.setFlags(
                    Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled
                )
                check_item.setCheckState(
                    Qt.CheckState.Checked if row.selected else Qt.CheckState.Unchecked
                )
            else:
                # Формулы и отсутствующие строки показываем, но применить нельзя
                check_item.setFlags(Qt.ItemFlag.NoItemFlags)
                check_item.setCheckState(Qt.CheckState.Unchecked)
            self.table.setItem(position, COL_CHECK, check_item)

            values = {
                # Та же подпись, что и в фильтре слева, — чтобы строку из
                # таблицы можно было глазами найти в списке трансгазов.
                COL_TRANSGAZ: row.transgaz or "(без трансгаза)",
                COL_OBJECT: row.object_name,
                COL_KEY: row.key,
                COL_FIELD: row.field_name,
                COL_CELL: row.address or f"копия {row.source_address}",
                COL_WAS: row.current_display or "—",
                COL_WILL: row.new_display or "(очистить)",
                COL_STATUS: row.status.value,
                COL_NOTE: row.note,
            }
            background = _ROW_COLORS.get(row.status)
            for column, text in values.items():
                item = QTableWidgetItem(str(text))
                if background is not None:
                    item.setBackground(QBrush(background))
                if row.status in BLOCKED:
                    item.setForeground(QBrush(QColor("#8b0000")))
                self.table.setItem(position, column, item)
            if background is not None:
                self.table.item(position, COL_CHECK).setBackground(QBrush(background))
        self.table.blockSignals(False)
        self._update_status_summary(rows)
        self._update_selected_label()

    def _update_status_summary(self, rows: list[DiffRow]) -> None:
        counts: dict[Status, int] = {}
        for row in rows:
            counts[row.status] = counts.get(row.status, 0) + 1
        parts = [f"{status.value}: {count}" for status, count in counts.items()]
        self.status_summary.setText("   ".join(parts) if parts else "нет строк")

    # --- выбор строк --------------------------------------------------------

    def _set_all_transgazes(self, checked: bool) -> None:
        self.transgaz_list.blockSignals(True)
        for index in range(self.transgaz_list.count()):
            self.transgaz_list.item(index).setCheckState(
                Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
            )
        self.transgaz_list.blockSignals(False)
        self._fill_table()

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != COL_CHECK:
            return
        row: DiffRow | None = item.data(Qt.ItemDataRole.UserRole)
        if row is None:
            return
        row.selected = item.checkState() == Qt.CheckState.Checked
        self._update_selected_label()

    def _set_all_rows(self, checked: bool) -> None:
        self.table.blockSignals(True)
        for position in range(self.table.rowCount()):
            item = self.table.item(position, COL_CHECK)
            row: DiffRow | None = item.data(Qt.ItemDataRole.UserRole)
            if row is None or not row.applicable:
                continue
            row.selected = checked
            item.setCheckState(
                Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
            )
        self.table.blockSignals(False)
        self._update_selected_label()

    def _selected_rows(self) -> list[DiffRow]:
        return [r for r in self.rows if r.selected and r.applicable]

    def _hidden_selected(self) -> list[DiffRow]:
        """Отмеченные строки, которых сейчас не видно из-за фильтра.

        Записаны они будут наравне с видимыми, поэтому молчать о них нельзя.
        """
        # Сравниваем по id: DiffRow изменяемый, в set его класть нельзя
        visible = {id(row) for row in self._visible_rows()}
        return [row for row in self._selected_rows() if id(row) not in visible]

    def _update_selected_label(self) -> None:
        selected = self._selected_rows()
        total = sum(1 for row in self.rows if row.applicable)
        text = f"отмечено {_changes(len(selected))} из {total} найденных"
        hidden = len(self._hidden_selected())
        if hidden:
            text += f"; {hidden} из отмеченных скрыто фильтром по трансгазам"
        self.selected_label.setText(text)
        self.apply_button.setEnabled(bool(selected) and not self.busy)

    # --- применение ---------------------------------------------------------

    def _start_apply(self) -> None:
        if self.busy:
            return
        selected = self._selected_rows()
        if not selected:
            return
        overwrites = sum(1 for r in selected if r.status is Status.OVERWRITE)
        clears = sum(1 for r in selected if r.status is Status.CLEAR)
        formulas = sum(1 for r in selected if r.replaces_formula)

        question = [
            f"Записать {_changes(len(selected))} в файл:",
            self.config.target_file,
            "",
        ]
        hidden = len(self._hidden_selected())
        if hidden:
            question.append(
                f"ВНИМАНИЕ: {hidden} из них не видно в таблице — они отмечены "
                "у трансгазов, снятых сейчас в фильтре слева, и будут записаны "
                "наравне с остальными."
            )
        if overwrites:
            question.append(f"Из них перезаписей заполненных ячеек: {overwrites}.")
        if clears:
            question.append(f"Из них очисток (значение будет стёрто): {clears}.")
        if formulas:
            question.append(
                f"Из них {formulas} заменят формулу на значение: в боевом файле "
                "там стоит ссылка на соседнюю ячейку вместо цифры. Расчётные "
                "столбцы (настройка «столбцы с формулами») при этом не трогаются."
            )
        question.append("Продолжить?")

        answer = QMessageBox.question(
            self,
            "Подтверждение",
            "\n".join(question),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        self._set_busy(True)
        self.worker.apply(self.rows)

    def _on_applied(self, result: ApplyResult) -> None:
        self._set_busy(False)
        self._reset_view()

        lines = [f"Готово: {result.summary()}"]
        if result.backup_path:
            lines.append(f"Бэкап: {result.backup_path}")
        if result.journal_path:
            lines.append(f"Журнал: {result.journal_path}")
        lines.extend(f"ОШИБКА: {error}" for error in result.errors)
        lines.extend(f"ВНИМАНИЕ: {warning}" for warning in result.verify_warnings)
        for line in lines:
            self._log(line)

        icon = (
            QMessageBox.Icon.Warning
            if result.errors or result.verify_warnings
            else QMessageBox.Icon.Information
        )
        QMessageBox(icon, "Применение завершено", "\n".join(lines), parent=self).exec()

    def _reset_view(self) -> None:
        self.compare_result = None
        self.rows = []
        self.table.setRowCount(0)
        self.transgaz_list.clear()
        self.apply_button.setEnabled(False)
        self._update_selected_label()

    # --- служебное ----------------------------------------------------------

    def _on_progress(self, message: str) -> None:
        self._log(message)

    def _on_failed(self, title: str, text: str) -> None:
        self._set_busy(False)
        self._log(f"{title}: {text}")
        QMessageBox.critical(self, title, text)

    def _log(self, message: str) -> None:
        self.log_view.append(message)
        self.log_view.ensureCursorVisible()

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        self.progress_bar.setVisible(busy)
        self.compare_button.setEnabled(not busy)
        if busy:
            self.apply_button.setEnabled(False)
            self._dialog_reported = False
            self.dialog_timer.start()
        else:
            self.dialog_timer.stop()
            self._update_selected_label()

    def _check_excel_dialogs(self) -> None:
        """Не висит ли на скрытом Excel модальное окно, которого нам не видно."""
        # При сравнении экземпляров Excel два — по одному на файл.
        session = self.worker.session
        pids = session.excel_pids if session else []
        dialogs = [title for pid in pids for title in find_dialogs(pid)]
        if not dialogs:
            self._dialog_reported = False
            return
        if self._dialog_reported:
            return
        self._dialog_reported = True

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle("Excel ждёт ответа")
        box.setText(describe(dialogs))
        box.setInformativeText(
            "Ответьте в окне Excel сами, либо прервите работу — тогда Excel "
            "будет закрыт. При сравнении боевой файл не меняется вовсе; при "
            "записи резервная копия уже сделана."
        )
        wait = box.addButton("Я отвечу сам", QMessageBox.ButtonRole.RejectRole)
        abort = box.addButton(
            "Прервать и закрыть Excel", QMessageBox.ButtonRole.DestructiveRole
        )
        box.setDefaultButton(wait)
        box.exec()

        if box.clickedButton() is abort:
            for pid in pids:
                self._log(f"Прерываю работу: закрываю Excel (PID {pid}).")
                with suppress(OSError):
                    os.kill(pid, signal.SIGTERM)

    def _open_settings(self) -> None:
        dialog = SettingsDialog(self.config, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                for note in self.config.ensure_dirs():
                    self._log(note)
            except DataDirError as error:
                QMessageBox.warning(self, "Папка бэкапов недоступна", str(error))
            self.worker.config = self.config
            self._log(f"Настройки сохранены: {self.config.path}")

    def _open_folder(self, path: Path) -> None:
        problem = probe_dir(path)
        if problem:
            QMessageBox.warning(self, "Папка недоступна", problem)
            return
        os.startfile(str(path))

    def closeEvent(self, event: Any) -> None:
        if self.busy:
            answer = QMessageBox.question(
                self,
                "Excel ещё работает",
                "Программа сейчас работает с файлами в Excel.\nПрервать и выйти?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.worker.stop()
        self.worker.wait(30000)
        event.accept()


def run() -> int:
    """Точка входа GUI.

    Папки готовим уже после создания QApplication: если писать некуда, человек
    должен увидеть окно с причиной, а не молчаливое падение при старте.
    """
    config = Config.load()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    notes: list[str] = []
    try:
        notes = config.ensure_dirs()
    except DataDirError as error:
        QMessageBox.warning(
            None,
            "Папка бэкапов недоступна",
            f"{error}\n\nПрограмма откроется, но применить изменения не сможет,\n"
            "пока в настройках не будет указана доступная папка.",
        )

    window = MainWindow(config)
    for note in notes + config._notes:
        window._log(note)
    window.show()
    return app.exec()
