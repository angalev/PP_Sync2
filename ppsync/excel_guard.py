"""Сторож модальных окон Excel.

Часть диалогов Excel не гасится ни ``DisplayAlerts = False``, ни
``AutomationSecurity``: ошибка макроса VBA («Run-time error… End»), «Конфликт
имен», «Обновить связи». Пока такое окно висит, любой COM-вызов к невидимому
Excel не возвращается — программа зависает молча и навсегда.

Модуль позволяет это увидеть: найти видимые диалоговые окна, принадлежащие
нашему экземпляру Excel, и показать их пользователю.

Одно окно — и только одно — модуль закрывает сам: «Конфликт имен» про
служебное имя вида ``_FilterDatabase``. Оно появляется на каждом открытии
боевой книги, ответ на него всегда один и тот же (любое непохожее имя), и
данных оно не касается — это имя диапазона автофильтра. Остальные окна
по-прежнему уходят человеку: на боевом файле цена случайного «Да» слишком
велика.

Автоответ обвешан проверками: окно принадлежит нашему процессу Excel,
заголовок — «Конфликт имен», а в тексте окна найдено имя из списка
встроенных (``name_check.BUILTIN_NAMES``). Не сошлось хоть что-то — не
трогаем.
"""

from __future__ import annotations

import ctypes
import re
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass

from .name_check import BUILTIN_NAMES

#: Классы окон, которые Excel использует для модальных диалогов.
#: bosa_sdm_* — «родные» диалоги Office, #32770 — стандартный диалог Windows,
#: NUIDialog — окна нового оформления Office (вход в учётную запись, лицензия).
_DIALOG_CLASSES = ("bosa_sdm", "#32770", "NUIDialog")

#: Главное окно Excel. Пока модальное окно висит, оно disabled — по этому
#: признаку ловятся диалоги неизвестного класса.
_MAIN_CLASS = "XLMAIN"

_user32 = ctypes.windll.user32
_EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


#: Что делать с конкретными диалогами. Ключ — кусок заголовка окна.
#: Без подсказки человек видит окно Excel и не знает, чем оно грозит боевому
#: файлу, поэтому жмёт наугад или закрывает программу.
_HINTS: tuple[tuple[str, str], ...] = (
    (
        "конфликт имен",
        (
            "Это служебное имя автофильтра, а не данные. Введите в поле "
            "«Новое имя» любое непохожее, например FilterDatabase_1, и нажмите "
            "ОК — Excel продолжит открывать книгу. На строки и формулы это "
            "не влияет."
        ),
    ),
    (
        "name conflict",
        (
            "Служебное имя автофильтра. Введите любое другое, например "
            "FilterDatabase_1, и нажмите ОК."
        ),
    ),
    (
        "учетные записи",
        (
            "Окно входа в учётную запись Microsoft. Закройте его крестиком — "
            "для работы с файлом вход не нужен."
        ),
    ),
    (
        "microsoft visual basic",
        "Ошибка макроса книги. Нажмите End (Завершить), а не Debug.",
    ),
)


def hint_for(title: str) -> str:
    """Подсказка по заголовку окна; пустая строка — окно нам незнакомо."""
    lowered = title.lower()
    for marker, text in _HINTS:
        if marker in lowered:
            return text
    return ""


def _window_text(hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(512)
    _user32.GetWindowTextW(hwnd, buffer, 512)
    return buffer.value


def _window_class(hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    _user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


@dataclass(frozen=True)
class DialogWindow:
    """Модальное окно нашего Excel: то, что видно снаружи процесса."""

    hwnd: int
    title: str
    class_name: str

    @property
    def name(self) -> str:
        """Как называть окно человеку."""
        return self.title or self.class_name


def find_dialog_windows(pid: int) -> tuple[list[DialogWindow], bool]:
    """Видимые диалоги процесса ``pid`` и признак «главное окно заблокировано».

    Пока модальное окно висит, XLMAIN остаётся ``disabled`` — по этому признаку
    ловятся и те диалоги, класс которых нам заранее неизвестен, иначе программа
    снова молчала бы про зависание.
    """
    if not pid:
        return [], False

    found: list[DialogWindow] = []
    blocked = False

    def callback(hwnd: int, _lparam: int) -> bool:
        nonlocal blocked
        window_pid = wintypes.DWORD()
        _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
        if window_pid.value != pid:
            return True
        class_name = _window_class(hwnd)
        if class_name == _MAIN_CLASS and not _user32.IsWindowEnabled(hwnd):
            blocked = True
            return True
        if not _user32.IsWindowVisible(hwnd):
            return True
        if not any(class_name.startswith(prefix) for prefix in _DIALOG_CLASSES):
            return True
        found.append(
            DialogWindow(hwnd=hwnd, title=_window_text(hwnd), class_name=class_name)
        )
        return True

    _user32.EnumWindows(_EnumWindowsProc(callback), 0)
    return found, blocked


def find_dialogs(pid: int) -> list[str]:
    """Заголовки видимых диалоговых окон процесса ``pid``.

    Пустой список — значит модальных окон нет и Excel просто занят работой.
    """
    windows, blocked = find_dialog_windows(pid)
    titles = [window.name for window in windows]
    if blocked and not titles:
        titles.append("неизвестное окно Excel")
    return titles


# --- автоответ на «Конфликт имен» ----------------------------------------

WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
BM_CLICK = 0x00F5
SMTO_ABORTIFHUNG = 0x0002
_SEND_TIMEOUT_MS = 3000

#: Заголовки окна «Конфликт имен» — русский и английский Excel.
_NAME_CONFLICT_TITLES = ("конфликт имен", "конфликт имён", "name conflict")

#: Тексты кнопки согласия. «ОК» пишут и латиницей, и кириллицей.
_OK_LABELS = frozenset({"ok", "ок", "да", "yes"})

#: Имя с этим префиксом Excel уже считает встроенным — про него не спрашивают.
_BUILTIN_PREFIX = "_xlnm."
_BUILTIN_LOWER = {name.lower(): name for name in BUILTIN_NAMES}
_TOKEN_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")

_EnumChildProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

# Для WM_SETTEXT/WM_GETTEXT в чужом процессе нужен SendMessage, а не
# GetWindowText: текст контролов другого приложения тот не отдаёт. Прототипы
# отличаются типом lParam (строка на запись, буфер на чтение), поэтому
# объявляем три отдельных.
_send_set_text = ctypes.WINFUNCTYPE(
    wintypes.LPARAM,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPCWSTR,
    wintypes.UINT,
    wintypes.UINT,
    ctypes.POINTER(wintypes.DWORD),
)(("SendMessageTimeoutW", _user32))

_send_get_text = ctypes.WINFUNCTYPE(
    wintypes.LPARAM,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPWSTR,
    wintypes.UINT,
    wintypes.UINT,
    ctypes.POINTER(wintypes.DWORD),
)(("SendMessageTimeoutW", _user32))

_send_plain = ctypes.WINFUNCTYPE(
    wintypes.LPARAM,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
    wintypes.UINT,
    wintypes.UINT,
    ctypes.POINTER(wintypes.DWORD),
)(("SendMessageTimeoutW", _user32))


def _children(hwnd: int) -> list[tuple[int, str]]:
    """Дочерние окна диалога парами «hwnd, класс»."""
    items: list[tuple[int, str]] = []

    def callback(child: int, _lparam: int) -> bool:
        items.append((child, _window_class(child)))
        return True

    _user32.EnumChildWindows(hwnd, _EnumChildProc(callback), 0)
    return items


def _control_text(hwnd: int) -> str:
    """Текст контрола из чужого процесса."""
    result = wintypes.DWORD()
    length = _send_plain(
        hwnd,
        WM_GETTEXTLENGTH,
        0,
        0,
        SMTO_ABORTIFHUNG,
        _SEND_TIMEOUT_MS,
        ctypes.byref(result),
    )
    size = int(result.value) if length else 0
    if size <= 0:
        return ""
    buffer = ctypes.create_unicode_buffer(size + 1)
    if not _send_get_text(
        hwnd,
        WM_GETTEXT,
        size + 1,
        buffer,
        SMTO_ABORTIFHUNG,
        _SEND_TIMEOUT_MS,
        ctypes.byref(result),
    ):
        return ""
    return buffer.value


def is_name_conflict(window: DialogWindow) -> bool:
    """Это окно «Конфликт имен»?"""
    lowered = window.title.lower()
    return any(marker in lowered for marker in _NAME_CONFLICT_TITLES)


def conflicting_builtin(hwnd: int) -> str:
    """Встроенное имя, о котором спрашивает окно; пустая строка — не нашли.

    Это главная проверка перед автоответом: без неё мы вслепую печатали бы
    текст в окно, содержания которого не понимаем.
    """
    for child, class_name in _children(hwnd):
        if class_name not in ("Static", "Edit"):
            continue
        for token in _TOKEN_RE.findall(_control_text(child)):
            if token.lower().startswith(_BUILTIN_PREFIX):
                continue  # имя уже с префиксом — окно не про него
            builtin = _BUILTIN_LOWER.get(token.lower())
            if builtin:
                return builtin
    return ""


def control_classes(hwnd: int) -> str:
    """Классы контролов окна — строкой для журнала.

    Нужны, когда ответить не вышло: по ним сразу видно, что за окно попалось
    и почему в нём не нашлось ни поля ввода, ни кнопки «ОК».
    """
    return ", ".join(sorted({class_name for _child, class_name in _children(hwnd)}))


def _find_controls(hwnd: int) -> tuple[int, int]:
    """Поле ввода и кнопка «ОК» диалога; 0 — если чего-то нет."""
    edit = ok = 0
    for child, class_name in _children(hwnd):
        if not edit and class_name == "Edit":
            edit = child
        elif not ok and class_name == "Button":
            label = _control_text(child).replace("&", "").strip().lower()
            if label in _OK_LABELS:
                ok = child
    return edit, ok


def replacement_name(builtin: str, attempt: int) -> str:
    """Имя, которое впишем вместо встроенного: FilterDatabase_1, _2, …"""
    return f"{builtin.lstrip('_')}_{attempt}"


def answer_name_conflict(hwnd: int, new_name: str) -> bool:
    """Вписать имя и нажать «ОК». ``False`` — не вышло, окно остаётся человеку."""
    edit, ok = _find_controls(hwnd)
    if not edit or not ok:
        return False

    result = wintypes.DWORD()
    _send_set_text(
        edit,
        WM_SETTEXT,
        0,
        new_name,
        SMTO_ABORTIFHUNG,
        _SEND_TIMEOUT_MS,
        ctypes.byref(result),
    )
    # Убеждаемся, что имя действительно легло в поле: жать «ОК» с пустым полем
    # нельзя — получится согласие непонятно на что.
    if _control_text(edit).strip() != new_name:
        return False

    _user32.PostMessageW(ok, BM_CLICK, 0, 0)
    deadline = time.monotonic() + 3.0
    while time.monotonic() < deadline:
        if not _user32.IsWindow(hwnd) or not _user32.IsWindowVisible(hwnd):
            return True
        time.sleep(0.05)
    return False


class DialogWatcher:
    """Фоновый поток: отвечает на «Конфликт имен», пока основной ждёт Excel.

    Отвечать обязан именно фоновый поток: пока висит модальное окно, основной
    стоит внутри COM-вызова ``Workbooks.Open`` и сделать ничего не может.

    Сделанное складывается в ``drain()``: писать в журнал из чужого потока
    сторож не будет, это дело вызывающего кода.
    """

    def __init__(self, pid: int, limit: int = 100, interval: float = 0.4) -> None:
        self.pid = pid
        self.limit = limit
        self.interval = interval
        self.answered = 0
        self._notes: list[str] = []
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._failed: set[int] = set()  # окна, с которыми не вышло

    def start(self) -> None:
        if self._thread is not None or not self.pid:
            return
        self._thread = threading.Thread(
            target=self._run, name="pp_sync_dialog_watcher", daemon=True
        )
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        thread, self._thread = self._thread, None
        if thread is not None:
            thread.join(timeout=2.0)

    def drain(self) -> list[str]:
        """Забрать накопленные сообщения (и очистить их)."""
        with self._lock:
            notes, self._notes = self._notes, []
        return notes

    def _note(self, text: str) -> None:
        with self._lock:
            self._notes.append(text)

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                self._tick()
            except Exception as error:  # noqa: BLE001 — сторож не смеет падать
                self._note(f"Сторож окон Excel споткнулся: {error}")
                return
            self._stop.wait(self.interval)

    def _tick(self) -> None:
        windows, _blocked = find_dialog_windows(self.pid)
        for window in windows:
            if window.hwnd in self._failed or not is_name_conflict(window):
                continue
            if self.answered >= self.limit:
                self._note(
                    f"Окон «Конфликт имен» больше {self.limit} — дальше отвечайте "
                    "сами: похоже, дело не в служебном имени."
                )
                self._stop.set()
                return
            builtin = conflicting_builtin(window.hwnd)
            if not builtin:
                self._failed.add(window.hwnd)
                self._note(
                    f"Окно «{window.name}» не про служебное имя — отвечайте сами."
                )
                continue
            new_name = replacement_name(builtin, self.answered + 1)
            if answer_name_conflict(window.hwnd, new_name):
                self.answered += 1
                self._note(
                    f"Excel спросил про служебное имя «{builtin}» — ответил "
                    f"«{new_name}» и продолжил открытие."
                )
            else:
                self._failed.add(window.hwnd)
                self._note(
                    f"Не удалось ответить в окне «{window.name}» — ответьте сами: "
                    "любое новое имя и ОК. "
                    f"(контролы окна: {control_classes(window.hwnd) or 'не видны'})"
                )


def describe(dialogs: list[str]) -> str:
    """Текст для пользователя по найденным окнам."""
    if not dialogs:
        return ""
    names = ", ".join(f"«{title}»" for title in dialogs)
    text = (
        f"Excel показал окно {names} и ждёт ответа. Пока оно открыто, "
        "программа не может продолжить: окно принадлежит скрытому экземпляру "
        "Excel, поэтому его не видно на панели задач."
    )
    hints = []
    for title in dialogs:
        hint = hint_for(title)
        if hint and hint not in hints:
            hints.append(hint)
    if hints:
        text += chr(10) * 2 + chr(10).join(hints)
    return text
