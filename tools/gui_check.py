"""Автопрогон GUI без участия человека.

Поднимает окно в offscreen-режиме на песочнице, жмёт «Сравнить», снимает одну
галочку, жмёт «Применить отмеченное» и проверяет результат.

    python -m tools.gui_check
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import QApplication, QMessageBox

from ppsync.config import Config
from ppsync.gui import COL_CHECK, MainWindow
from ppsync.models import Status
from tools import make_test_workbook

failures: list[str] = []
steps: list[str] = []


def main() -> int:
    sandbox = Path(tempfile.mkdtemp(prefix="pp_sync_gui_"))
    target_path = sandbox / "ПП_боевой.xlsx"
    source_path = sandbox / "ПП_копия.xlsx"
    make_test_workbook.build(target_path)
    make_test_workbook.build_changed(target_path, source_path)

    config = Config()
    config.target_file = str(target_path)
    config.source_file = str(source_path)
    config.data_dir = str(sandbox / "данные")
    config.sheet_name = make_test_workbook.SHEET
    config._path = str(sandbox / "config.json")
    config.ensure_dirs()

    app = QApplication([])
    # Диалоги подтверждения в автопрогоне отвечают «да» сами
    QMessageBox.question = staticmethod(  # type: ignore[method-assign]
        lambda *a, **k: QMessageBox.StandardButton.Yes
    )
    QMessageBox.information = staticmethod(lambda *a, **k: 0)  # type: ignore[method-assign]
    QMessageBox.exec = lambda self: 0  # type: ignore[method-assign]

    window = MainWindow(config)
    window.show()

    def after_compare() -> None:
        steps.append(f"сравнение получено: {len(window.rows)} различий")
        if window.table.rowCount() != len(window.rows):
            failures.append(
                f"в таблице {window.table.rowCount()} строк, а различий {len(window.rows)}"
            )
        if not window.rows:
            failures.append("сравнение пустое")
            app.quit()
            return

        blocked = [r for r in window.rows if not r.applicable]
        if not blocked:
            failures.append("нет заблокированных строк — проверка неполная")
        for position in range(window.table.rowCount()):
            item = window.table.item(position, COL_CHECK)
            row = item.data(Qt.ItemDataRole.UserRole)
            if not row.applicable and item.flags() != Qt.ItemFlag.NoItemFlags:
                failures.append(f"строка {row.key}: чекбокс не заблокирован")

        clears = [r for r in window.rows if r.status is Status.CLEAR]
        if any(r.selected for r in clears):
            failures.append("очистка отмечена по умолчанию — так быть не должно")

        applicable = [r for r in window.rows if r.applicable]
        if len(applicable) < 2:
            failures.append("мало применимых различий")
            app.quit()
            return

        # После сравнения не должно быть отмечено ничего: запись идёт в боевой
        # файл, и «отмечено всё» человек замечает уже в окне подтверждения.
        if any(r.selected for r in window.rows):
            failures.append("после сравнения строки уже отмечены — так быть не должно")
        if "из 0 найденных" in window.selected_label.text():
            failures.append(
                f"счётчик не знает общего числа: {window.selected_label.text()}"
            )

        window._set_all_rows(True)
        applicable[-1].selected = False
        window._fill_table()
        selected = window._selected_rows()
        if len(selected) != len(applicable) - 1:
            failures.append("счётчик отмеченных не совпал")
        label = window.selected_label.text()
        if f"отмечено {len(selected)} " not in label or "найденных" not in label:
            failures.append(f"надпись счётчика неожиданная: {label}")
        steps.append(f"надпись: {label}")
        window._start_apply()

    def after_apply(result) -> None:
        steps.append(f"применение: {result.summary()}")
        if result.applied < 1:
            failures.append("ничего не применено")
        failures.extend(result.errors)
        if window.table.rowCount() != 0:
            failures.append("таблица не очищена после применения")
        app.quit()

    window.worker.compared.connect(lambda _r: QTimer.singleShot(0, after_compare))
    window.worker.applied.connect(after_apply)
    window.worker.failed.connect(lambda title, text: _fail(app, title, text))

    QTimer.singleShot(0, window._start_compare)
    QTimer.singleShot(180_000, lambda: _fail(app, "таймаут", "не уложились в 3 минуты"))
    app.exec()

    window.worker.stop()
    window.worker.wait(30_000)

    print("\nШаги:")
    for step in steps:
        print(f"  • {step}")
    if failures:
        print("\nПРОВАЛЫ:")
        for failure in failures:
            print(f"  • {failure}")
        return 1
    print("\nGUI-прогон пройден.")
    return 0


def _fail(app: QApplication, title: str, text: str) -> None:
    failures.append(f"{title}: {text}")
    app.quit()


if __name__ == "__main__":
    raise SystemExit(main())
