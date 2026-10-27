r"""Проверка автоответа на «Конфликт имен» — на подставном окне.

    python -m tools.dialog_answer_check

Настоящее окно Excel на домашней машине не воспроизводится (см. журнал ошибок),
но проверить надо именно слепую часть: умеет ли сторож прочитать текст чужого
окна, вписать имя в поле и нажать «ОК». Поэтому поднимаем стандартный диалог
Windows того же класса ``#32770`` с полем ввода и кнопкой «ОК» — им работает
``InputBox`` из VBScript, — даём ему заголовок «Конфликт имен» и текст с именем
``_FilterDatabase``, после чего натравливаем на него ``DialogWatcher``.

Проверяется то же, что решает дело на боевом файле: окно найдено по нашему
процессу, встроенное имя в тексте распознано, в поле легло новое имя, окно
закрылось именно кнопкой «ОК» (а не «Отмена» — тогда VBScript вернул бы пустую
строку).
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppsync.excel_guard import (
    DialogWatcher,
    conflicting_builtin,
    find_dialog_windows,
    is_name_conflict,
)

TITLE = "Конфликт имен"
#: Текст окна — уже как выражение VBScript: перенос строки там делается
#: сцеплением через vbCrLf, внутрь строковой константы его класть нельзя.
PROMPT = (
    '"Имя не может совпадать со встроенным именем." & vbCrLf & '
    '"Старое имя:    _FilterDatabase"'
)
EXPECTED = "FilterDatabase_1"

failures: list[str] = []


def _script(answer_path: Path) -> str:
    return (
        f"Dim answer\r\n"
        f'answer = InputBox({PROMPT}, "{TITLE}")\r\n'
        f"Dim fso, target\r\n"
        f'Set fso = CreateObject("Scripting.FileSystemObject")\r\n'
        f'Set target = fso.CreateTextFile("{answer_path}", True, True)\r\n'
        f"target.Write answer\r\n"
        f"target.Close\r\n"
    )


def _wait_for_dialog(pid: int, seconds: float = 15.0) -> object | None:
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        windows, _blocked = find_dialog_windows(pid)
        for window in windows:
            if is_name_conflict(window):
                return window
        time.sleep(0.2)
    return None


def main() -> int:
    sandbox = Path(tempfile.mkdtemp(prefix="pp_sync_dialog_"))
    answer_path = sandbox / "ответ.txt"
    script_path = sandbox / "конфликт.vbs"
    # cscript читает Unicode-скрипт по метке порядка байтов
    script_path.write_text(_script(answer_path), encoding="utf-16")

    try:
        process = subprocess.Popen(
            ["cscript.exe", "//nologo", str(script_path)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except OSError as error:
        print(f"cscript.exe недоступен ({error}) — проверку пропускаю.")
        return 0

    try:
        window = _wait_for_dialog(process.pid)
        if window is None:
            failures.append("подставное окно не появилось — проверять нечего")
            return _report(process)

        builtin = conflicting_builtin(window.hwnd)  # type: ignore[attr-defined]
        if builtin != "_FilterDatabase":
            failures.append(f"встроенное имя не распознано: {builtin!r}")

        watcher = DialogWatcher(process.pid, interval=0.2)
        watcher.start()
        deadline = time.monotonic() + 15.0
        while time.monotonic() < deadline and watcher.answered == 0:
            time.sleep(0.2)
        notes = watcher.drain()
        watcher.stop()

        if watcher.answered != 1:
            failures.append(f"сторож не ответил: {notes}")
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            failures.append("окно не закрылось")

        written = ""
        if answer_path.exists():
            written = answer_path.read_text(encoding="utf-16").strip()
        if written != EXPECTED:
            failures.append(
                f"в поле оказалось {written!r}, а ожидалось {EXPECTED!r} "
                "(пусто — значит нажали «Отмена», а не «ОК»)"
            )
        for note in notes:
            print(f"  сторож: {note}")
        return _report(process)
    finally:
        if process.poll() is None:
            process.kill()


def _report(process: subprocess.Popen[bytes]) -> int:
    if process.poll() is None:
        process.kill()
    if failures:
        print("ПРОВЕРКА НЕ ПРОЙДЕНА:")
        for line in failures:
            print(f"  • {line}")
        return 1
    print("Автоответ на «Конфликт имен»: проверка пройдена.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
