"""Настройки программы (config.json).

Ничего не хардкодится в логике: пути к обоим файлам, имя листа, номера строк и
буквы столбцов лежат здесь и редактируются из GUI.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

DEFAULT_DATA_DIR = Path(r"C:\ПП_синхронизация")
CONFIG_FILENAME = "config.json"
ENV_CONFIG_PATH = "PP_SYNC_CONFIG"
PROBE_NAME = ".pp_sync_write_test"

# Режимы Application.Calculation
XL_CALCULATION_AUTOMATIC = -4105
XL_CALCULATION_MANUAL = -4135

# Режимы Book.ConflictResolution (общий доступ)
XL_USER_RESOLUTION = 1  # модальный диалог — для невидимого Excel смертельно
XL_LOCAL_SESSION_CHANGES = 2  # побеждают наши изменения
XL_OTHER_SESSION_CHANGES = 3  # побеждают чужие

CONFLICT_MODES = {
    XL_USER_RESOLUTION: "1 — спрашивать (ОПАСНО: диалог повесит скрытый Excel)",
    XL_LOCAL_SESSION_CHANGES: "2 — побеждают наши изменения (рекомендуется)",
    XL_OTHER_SESSION_CHANGES: "3 — побеждают чужие изменения",
}

_RANGE_RE = re.compile(r"^(\d+)\s*(?:-|–|:)\s*(\d+)$")
_LETTERS_RE = re.compile(r"^[A-Za-z]{1,3}(?:\s*(?:-|–|:)\s*[A-Za-z]{1,3})?$")


class DataDirError(RuntimeError):
    """Некуда складывать бэкапы и журнал: обе папки недоступны."""


def fallback_data_dir() -> Path:
    """Куда переезжаем, если основная папка закрыта.

    ``%LOCALAPPDATA%`` есть у любого доменного пользователя и пишется всегда —
    в отличие от корня диска C:, где на рабочих машинах создание папок обычно
    запрещено политикой, и от сетевого диска, который может быть только на чтение.
    """
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return Path(base) / "PP_sync"


def explain_path_error(path: Path, error: OSError) -> str:
    """Человеческая причина вместо «WinError 5».

    Windows кладёт свой код в ``winerror``, но при нехватке места или
    read-only-шаре остаётся только ``errno`` — учитываем оба.
    """
    code: int | None = getattr(error, "winerror", None) or error.errno
    reasons: dict[int | None, str] = {
        2: "путь не существует",
        3: "путь не существует",
        5: "нет прав на запись (папку закрыла политика или администратор)",
        13: "нет прав на запись",
        28: "на диске нет свободного места",
        30: "диск доступен только для чтения",
        53: "сетевой путь недоступен",
        67: "сетевое имя не найдено — диск не подключён",
        1326: "не пройдена авторизация на сервере",
        1920: "система не может получить доступ к файлу",
    }
    reason = reasons.get(code, str(error))
    return f"{path} — {reason}"


def probe_dir(path: Path) -> str:
    """Создать папку и убедиться, что в неё пишется.

    Пустая строка — всё в порядке; иначе текст с причиной. Одного ``mkdir``
    мало: сетевой диск может дать создать папку и не дать записать файл.
    """
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / PROBE_NAME
        probe.write_text("", encoding="utf-8")
        probe.unlink()
    except OSError as error:
        return explain_path_error(path, error)
    return ""


def default_config_path() -> Path:
    """Где лежит config.json: переменная окружения, основная папка или запасная.

    Если ``C:\\ПП_синхронизация`` уже существует — работаем с ней, даже если она
    недоступна на запись: пусть об этом скажет проверка, а не молчаливый переезд
    на пустые настройки.
    """
    env = os.environ.get(ENV_CONFIG_PATH)
    if env:
        return Path(env)
    if DEFAULT_DATA_DIR.exists() or not probe_dir(DEFAULT_DATA_DIR):
        return DEFAULT_DATA_DIR / CONFIG_FILENAME
    return fallback_data_dir() / CONFIG_FILENAME


def parse_codes(text: str) -> set[str]:
    """Разобрать «259-267, 316, 320» в набор кодов столбцов.

    Код — то, что стоит в строке кодов листа (в ПП это строка 9). Диапазон
    через дефис раскрывается только для чисел. Пустая строка — пустой набор
    (в вызывающем коде это значит «все столбцы»). Пустые куски пропускаются:
    настройку правит человек руками, и падать из-за лишней запятой нельзя.
    """
    codes: set[str] = set()
    for chunk in re.split(r"[,;]", text or ""):
        item = " ".join(chunk.split())
        if not item:
            continue
        match = _RANGE_RE.match(item)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            codes.update(str(n) for n in range(min(start, end), max(start, end) + 1))
        else:
            codes.add(item)
    return codes


def letters_only(text: str) -> list[str]:
    """Куски настройки, похожие на буквы столбцов, а не на коды.

    До перехода на коды столбцы задавались буквами (``BT-CB, DX``). Такая
    запись теперь не найдёт ни одного столбца, и молчать об этом нельзя.
    """
    return [
        chunk.strip()
        for chunk in re.split(r"[,;]", text or "")
        if _LETTERS_RE.match(chunk.strip())
    ]


# Расчётные столбцы по умолчанию до формы от 14.10.2026. Сохранённый с ними
# config.json при чтении переходит на новый список (см. Config.load).
_OLD_FORMULA_CODES = (
    "33, 104, 130, 276, 278, 280, 282, 284, 286, 288, 290, 292, 294, 296, "
    "298, 300, 304, 311, 313, 321, 323, 324, 326, 327"
)


@dataclass
class Config:
    """Настройки. Значения по умолчанию — по замерам на боевом файле."""

    # Два файла: куда пишем и откуда берём изменения
    target_file: str = r"\\server\share\Производственная программа 2026.xlsm"
    source_file: str = ""

    # Структура листа
    sheet_name: str = "ПП 2026"
    header_row_top: int = 6
    header_row_bottom: int = 7
    code_row: int = 9  # строка с кодами столбцов — по ней находятся все столбцы
    first_data_row: int = 10

    # Столбцы задаются КОДАМИ из строки кодов, а не буквами: буквы уезжают при
    # любом удалении или вставке столбца, коды остаются при своих столбцах.
    # Поэтому боевой файл и копия могут быть разной ширины — столбцы
    # сопоставляются по коду.
    key_code: str = "39"  # сводный код X/Y/Z — по нему сопоставляются строки
    transgaz_code: str = "7"  # наименование трансгаза
    object_code: str = "14"  # наименование объекта — только для показа

    # Какие столбцы сравнивать (коды). Пусто — все. «Свой блок» снижает шум.
    compare_codes: str = ""
    ignore_codes: str = ""

    # Столбцы, где формула — часть конструкции листа, а не чья-то правка.
    # Формулу в них программа не заменяет значением никогда.
    #
    # Список составлен по форме ПП 2026 от 14.10.2026 (1279 строк данных):
    #   33 «Всего», 104 «Ожидаемое_итого», 130 «Закрытие ДП», 330 «Остаток
    #     к закрытию», 321 «Техакт подписан» — суммы и признаки по строке;
    #   311 «Кол-во (план)», 313 «% выполнения» — расчёт физического выполнения;
    #   326, 327 — фактические даты полевых работ.
    #
    # Код 323 в новой форме — «Статус камеральных работ», столбец ввода; в
    # старых форматах под ним были даты полевых работ, поэтому из списка убран.
    #
    # Всё остальное формульным столбцом не считается. Там формула — это
    # односложная ссылка на соседнюю ячейку (=DS515, =O152): работник не стал
    # впечатывать цифру. В ПП это 312 «Кол-во (факт)», даты факта 307, 308,
    # 310, 317-319. Такую формулу программа заменяет значением из копии —
    # с пометкой в таблице и в журнале.
    formula_codes: str = "33, 104, 130, 311, 313, 321, 326, 327, 330"

    # Общий доступ и поведение Excel
    conflict_resolution: int = XL_LOCAL_SESSION_CHANGES
    verify_after_save: bool = True  # перечитать ячейки после save() и сверить
    disable_macros: bool = True  # не запускать макросы книги при открытии
    disable_events: bool = True  # не запускать обработчики событий листа

    # Служебное имя _FilterDatabase без префикса _xlnm. заставляет Excel
    # спрашивать «Конфликт имен» при каждом открытии книги. Две меры против
    # этого: починить имя в файле до открытия и ответить за человека, если
    # окно всё-таки появилось (например, книгу в этот момент кто-то держит).
    fix_service_names: bool = True
    auto_answer_name_conflict: bool = True

    # Куда складывать резервные копии и журнал
    data_dir: str = str(DEFAULT_DATA_DIR)
    backup_keep: int = 30  # сколько копий держать (0 — не чистить)

    _path: str = field(default="", repr=False, compare=False)
    # Что сказать человеку после чтения старого config.json (см. load)
    _notes: list[str] = field(default_factory=list, repr=False, compare=False)

    # --- пути ---------------------------------------------------------------

    @property
    def data_path(self) -> Path:
        return Path(self.data_dir)

    @property
    def backup_dir(self) -> Path:
        return self.data_path / "backup"

    @property
    def log_dir(self) -> Path:
        return self.data_path

    @property
    def target_path(self) -> Path:
        return Path(self.target_file)

    @property
    def source_path(self) -> Path:
        return Path(self.source_file)

    @property
    def path(self) -> Path:
        return Path(self._path) if self._path else default_config_path()

    def ensure_dirs(self) -> list[str]:
        """Подготовить папки для бэкапов и журнала.

        Возвращает предупреждения. Если указанная папка недоступна (корень C:
        закрыт политикой, сетевой диск отвалился или дан только на чтение),
        программа не падает, а переезжает в ``%LOCALAPPDATA%\\PP_sync`` и
        говорит об этом вслух. Если и там нельзя — ``DataDirError``.
        """
        problem = probe_dir(self.data_path)
        if not problem:
            problem = probe_dir(self.backup_dir)
        if not problem:
            return []

        fallback = fallback_data_dir()
        if fallback == self.data_path:
            raise DataDirError(f"Некуда писать бэкапы и журнал: {problem}")
        second = probe_dir(fallback) or probe_dir(fallback / "backup")
        if second:
            raise DataDirError(
                f"Некуда писать бэкапы и журнал.\n{problem}\n{second}\n"
                "Укажите доступную папку в настройках."
            )

        self.data_dir = str(fallback)
        note = (
            f"Папка бэкапов недоступна ({problem}). "
            f"Бэкапы и журнал складываю в {fallback}."
        )
        return [note]

    # --- область сравнения --------------------------------------------------

    def code_scope(self, codes: list[str]) -> list[str]:
        """Какие из кодов ``codes`` надо сравнивать — в том же порядке."""
        include = parse_codes(self.compare_codes)
        exclude = parse_codes(self.ignore_codes)
        # Ключ и трансгаз сравнивать не нужно: по ключу строки сопоставляются,
        # а трансгаз показывается для наглядности.
        exclude |= {self.key_code, self.transgaz_code}
        return [
            code
            for code in codes
            if code not in exclude and (not include or code in include)
        ]

    def protected_codes(self) -> set[str]:
        """Столбцы, где формулу в боевом файле заменять значением нельзя.

        Во всех прочих столбцах формула в боевом — это ленивая ссылка на
        соседнюю ячейку, и её надо перезаписать: именно из-за неё фактические
        объёмы («Кол-во (факт)») не доезжали до боевого файла.
        """
        return parse_codes(self.formula_codes)

    # --- чтение/запись ------------------------------------------------------

    @classmethod
    def load(cls, path: str | Path | None = None) -> Config:
        """Прочитать config.json; если его нет — вернуть значения по умолчанию."""
        config_path = Path(path) if path else default_config_path()
        cfg = cls()
        if config_path.exists():
            data: dict[str, Any] = json.loads(config_path.read_text(encoding="utf-8"))
            known = {f.name for f in fields(cls) if not f.name.startswith("_")}
            for key, value in data.items():
                if key in known:
                    setattr(cfg, key, value)
            cfg._notes = cls._legacy_notes(data)
            if parse_codes(cfg.formula_codes) == parse_codes(_OLD_FORMULA_CODES):
                # Список по умолчанию до формы от 14.10.2026 — человек его не
                # менял, значит, переходит на новый сам. Свой список не трогаем.
                cfg.formula_codes = cls.formula_codes
            if "code_row" not in data and int(data.get("service_row") or 0) > 0:
                cfg.code_row = int(data["service_row"])
        cfg._path = str(config_path)
        cfg.normalize()
        return cfg

    def save(self, path: str | Path | None = None) -> Path:
        config_path = Path(path) if path else self.path
        problem = probe_dir(config_path.parent)
        if problem and not path:
            # Настройки важнее места: если основная папка закрыта, кладём
            # config.json в запасную, иначе программа забудет всё при выходе.
            fallback = fallback_data_dir() / CONFIG_FILENAME
            second = probe_dir(fallback.parent)
            if second:
                raise DataDirError(
                    f"Не удалось сохранить настройки.\n{problem}\n{second}"
                )
            config_path = fallback
        elif problem:
            raise DataDirError(f"Не удалось сохранить настройки.\n{problem}")
        data = {k: v for k, v in asdict(self).items() if not k.startswith("_")}
        config_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        self._path = str(config_path)
        return config_path

    # --- валидация ----------------------------------------------------------

    @staticmethod
    def _legacy_notes(data: dict[str, Any]) -> list[str]:
        """Что из старого config.json (столбцы буквами) не переехало само.

        Ключ, трансгаз, объект и расчётные столбцы переходят на коды по
        умолчанию молча: это те же столбцы ПП. А вот «свой блок» человек
        задавал сам, и перевести буквы в коды за него нельзя — буква зависит
        от того, в каком формате был файл.
        """
        notes: list[str] = []
        for old, title in (
            ("compare_columns", "сравнивать столбцы"),
            ("ignore_columns", "не сравнивать столбцы"),
        ):
            value = str(data.get(old) or "").strip()
            if value:
                notes.append(
                    f"Настройка «{title}» была задана буквами ({value}). Теперь "
                    "столбцы задаются кодами из строки кодов листа — впишите "
                    "коды заново, пока сравниваются все столбцы."
                )
        return notes

    def normalize(self) -> None:
        self.key_code = " ".join(str(self.key_code).split())
        self.transgaz_code = " ".join(str(self.transgaz_code).split())
        self.object_code = " ".join(str(self.object_code).split())
        self.sheet_name = str(self.sheet_name).strip()

    def validate(self) -> list[str]:
        """Список проблем в настройках (пустой список — всё в порядке)."""
        problems: list[str] = []

        if not self.target_file.strip():
            problems.append("Не указан боевой файл.")
        elif not self.target_path.exists():
            problems.append(f"Боевой файл не найден: {self.target_file}")

        if not self.source_file.strip():
            problems.append("Не указан файл с изменениями.")
        elif not self.source_path.exists():
            problems.append(f"Файл с изменениями не найден: {self.source_file}")
        elif (
            self.target_file.strip()
            and self.source_path.resolve() == self.target_path.resolve()
        ):
            problems.append(
                "Файл с изменениями и боевой файл — это один и тот же файл."
            )

        if not self.sheet_name:
            problems.append("Не указано имя листа.")
        if self.first_data_row <= self.header_row_bottom:
            problems.append(
                "Первая строка данных должна быть ниже строки заголовков "
                f"({self.first_data_row} <= {self.header_row_bottom})."
            )
        if not 0 < self.code_row < self.first_data_row:
            problems.append(
                "Строка кодов столбцов должна быть выше первой строки данных "
                f"(сейчас {self.code_row} и {self.first_data_row})."
            )
        if not self.key_code:
            problems.append("Не указан код столбца с ключом X/Y/Z.")
        for title, text in (
            ("сравнивать столбцы", self.compare_codes),
            ("не сравнивать столбцы", self.ignore_codes),
            ("столбцы с формулами", self.formula_codes),
        ):
            letters = letters_only(text)
            if letters:
                problems.append(
                    f"В настройке «{title}» стоят буквы столбцов "
                    f"({', '.join(letters)}), а нужны коды из строки "
                    f"{self.code_row} листа, например 316 или 259-267."
                )
        if self.conflict_resolution not in CONFLICT_MODES:
            problems.append(
                f"ConflictResolution должен быть 1, 2 или 3 "
                f"(сейчас {self.conflict_resolution})."
            )
        if self.conflict_resolution == XL_USER_RESOLUTION:
            problems.append(
                "ConflictResolution=1 показывает модальный диалог — "
                "скрытый Excel зависнет. Поставьте 2."
            )
        return problems
