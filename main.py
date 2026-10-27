"""Запуск программы синхронизации Производственной программы.

    python main.py

Настройки берутся из config.json (по умолчанию C:\\ПП_синхронизация\\config.json,
путь можно переопределить переменной окружения PP_SYNC_CONFIG).

В собранном exe окна консоли нет, поэтому необработанное исключение некуда
напечатать: программа просто исчезла бы с экрана без следа. Поэтому здесь стоит
перехват — он пишет traceback в ``crash.log`` **рядом с самим exe** и показывает
системное окно с ошибкой. Системное, а не Qt: падать умеет и сам Qt.
"""

from __future__ import annotations

import datetime as _dt
import sys
import tempfile
import traceback
from contextlib import suppress
from pathlib import Path
from types import TracebackType

from ppsync.gui import run


def _writable(directory: Path) -> bool:
    """Проверка без импорта ppsync: сюда приходят и падения самого ppsync."""
    try:
        directory.mkdir(parents=True, exist_ok=True)
        probe = directory / ".pp_sync_report_test"
        probe.write_text("", encoding="utf-8")
        probe.unlink()
    except OSError:
        return False
    return True


def _report_dir() -> Path:
    """Куда класть ``crash.log`` и ``selftest.log``.

    Первым делом — рядом с самим exe: человек ищет журнал там, где запускал
    программу, а не в папке настроек, о существовании которой мог не знать.
    Дальше папка config.json, и уже в безнадёжном случае — ``%TEMP%``.
    """
    candidates: list[Path] = []
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent)
    with suppress(Exception):  # ppsync мог и не импортироваться
        from ppsync.config import default_config_path

        candidates.append(default_config_path().parent)
    candidates.append(Path(tempfile.gettempdir()))

    for directory in candidates:
        if _writable(directory):
            return directory
    return Path(tempfile.gettempdir())


def _crash_log_path() -> Path:
    return _report_dir() / "crash.log"


def _report_crash(text: str) -> None:
    """Записать traceback в журнал и показать его человеку."""
    stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # noqa: DTZ005
    where = "журнал записать не удалось"
    with suppress(Exception):  # некуда писать — покажем хотя бы окно
        path = _crash_log_path()
        with path.open("a", encoding="utf-8") as handle:
            handle.write(f"\n=== {stamp} ===\n{text}")
        where = str(path)

    lines = [line for line in text.strip().splitlines() if line.strip()]
    summary = lines[-1] if lines else "неизвестная ошибка"
    with suppress(Exception):  # без окна переживём: traceback уже в журнале
        import ctypes

        ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
            None,
            f"{summary}\n\nПодробности: {where}",
            "Синхронизация ПП — сбой",
            0x10,  # MB_ICONERROR
        )


def _excepthook(
    kind: type[BaseException],
    value: BaseException,
    tb: TracebackType | None,
) -> None:
    """PyQt6 при необработанном исключении в слоте гасит процесс через abort.

    До этого он зовёт ``sys.excepthook`` — единственное место, где ошибку внутри
    цикла событий ещё можно увидеть.
    """
    _report_crash("".join(traceback.format_exception(kind, value, tb)))


def _selftest() -> int:
    """``ПП_синхронизация.exe --selftest`` — проверка самой сборки, без Excel.

    В собранном exe тяжёлые модули грузятся лениво, уже в момент работы с боевым
    файлом: недостающий в сборке xlwings вылез бы только там. Ключ проверяет, что
    exe целый и настройки читаются, ничего при этом не открывая и не записывая.
    Результат — окно с отчётом и ``selftest.log`` рядом с exe, он же код возврата.
    Ключ ``--quiet`` убирает окно: оно модальное и скрипт бы на нём повис.
    """
    lines: list[str] = []
    ok = True

    # Где мы вообще работаем: пути с кириллицей и запуск с сетевого диска —
    # первое, на что смотрят, когда «на этой машине не работает».
    lines.append("--- окружение ---")
    lines.append(f"exe: {sys.executable}")
    lines.append(f"временная папка: {tempfile.gettempdir()}")
    if getattr(sys, "frozen", False):
        lines.append(f"распаковано в: {getattr(sys, '_MEIPASS', '—')}")
    lines.append("")
    lines.append("--- модули ---")

    for name in (
        "PyQt6.QtWidgets",
        "openpyxl",
        "xlwings",
        "pythoncom",
        "win32com.client",
    ):
        try:
            __import__(name)
        except Exception as error:  # noqa: BLE001 — сюда и надо ловить всё
            ok = False
            lines.append(f"НЕТ  {name}: {error}")
        else:
            lines.append(f"есть {name}")

    try:
        from ppsync.config import (
            DEFAULT_DATA_DIR,
            Config,
            default_config_path,
            explain_path_error,
            fallback_data_dir,
            probe_dir,
        )

        config = Config.load()
        lines.append(f"есть настройки: {default_config_path()}")

        # Папки: на рабочих машинах корень диска C: часто закрыт политикой,
        # а сетевой диск бывает подключён только на чтение.
        lines.append("")
        lines.append("--- папки ---")
        checks = [
            ("основная", DEFAULT_DATA_DIR),
            ("из настроек", config.data_path),
            ("бэкапы", config.backup_dir),
        ]
        trouble = False
        for title, path in checks:
            problem = probe_dir(path)
            trouble = trouble or bool(problem)
            lines.append(f"{'НЕТ ' if problem else 'есть'} {title}: {problem or path}")
        if trouble:
            # Запасную папку создаём только когда она действительно нужна.
            spare = fallback_data_dir()
            problem = probe_dir(spare)
            lines.append(
                f"{'НЕТ ' if problem else 'есть'} запасная: {problem or spare}"
            )
            ok = ok and not problem

        # Файлы: сюда попадают отказы сетевого диска.
        lines.append("")
        lines.append("--- файлы ---")
        for title, path in (
            ("боевой", config.target_path),
            ("изменения", config.source_path),
        ):
            if not str(path).strip() or str(path) == ".":
                lines.append(f"—    {title}: не указан")
                continue
            try:
                size = path.stat().st_size
            except OSError as error:
                ok = False
                lines.append(f"НЕТ  {title}: {explain_path_error(path, error)}")
            else:
                lines.append(f"есть {title}: {path} ({size // 1024} КБ)")

        # Служебные имена: из-за имени вида _FilterDatabase без префикса
        # _xlnm. Excel при открытии показывает окно «Конфликт имен».
        from ppsync import name_check

        lines.append("")
        lines.append("--- служебные имена ---")
        for title, path in (
            ("боевой", config.target_path),
            ("изменения", config.source_path),
        ):
            if not path.exists():
                continue
            try:
                problems = name_check.scan(path)
            except name_check.NameCheckError as error:
                lines.append(f"?    {title}: прочитать не удалось ({error})")
                continue
            if problems:
                for text in name_check.describe(problems):
                    lines.append(f"!    {title}: {text}")
            else:
                lines.append(f"есть {title}: служебные имена в порядке")
    except Exception as error:  # noqa: BLE001
        ok = False
        lines.append(f"НЕТ  настройки: {error}")

    lines.append("")
    lines.append("ИТОГ: всё на месте" if ok else "ИТОГ: есть проблемы, см. выше")

    saved = ""
    with suppress(Exception):
        path = _report_dir() / "selftest.log"
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        saved = str(path)
    if saved:
        lines.append(f"отчёт: {saved}")
    report = "\n".join(lines)
    print(report)

    # У собранного exe консоли нет: без окна проверка выглядела бы как
    # «ничего не произошло». Окно модальное, поэтому для скриптов есть --quiet.
    if getattr(sys, "frozen", False) and "--quiet" not in sys.argv:
        with suppress(Exception):
            import ctypes

            ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
                None,
                report,
                "Синхронизация ПП — проверка",
                0x40 if ok else 0x30,  # MB_ICONINFORMATION / MB_ICONWARNING
            )
    return 0 if ok else 1


if __name__ == "__main__":
    sys.excepthook = _excepthook
    if "--selftest" in sys.argv:
        sys.exit(_selftest())
    try:
        code = run()
    except SystemExit:
        raise
    except Exception:
        _report_crash(traceback.format_exc())
        raise
    sys.exit(code)
