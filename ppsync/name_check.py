"""Проверка и починка служебных имён книги — без запуска Excel.

Excel хранит встроенные имена (область печати, диапазон автофильтра и т.п.) с
префиксом ``_xlnm.``. Если в книге лежит имя ``_FilterDatabase`` без префикса,
Excel считает его пользовательским, видит совпадение со встроенным и при
**открытии** показывает модальное окно «Конфликт имён» с просьбой переименовать.
Невидимому экземпляру ответить некому: окно висит, а ``Workbooks.Open``
заканчивается ошибкой «Метод Open из класса Workbooks завершен неверно».

Поэтому книгу проверяем до Excel: xlsx/xlsm — это zip, а список имён лежит в
``xl/workbook.xml``. Чтение ничего не открывает и не блокирует файл.
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path

WORKBOOK_PART = "xl/workbook.xml"
BUILTIN_PREFIX = "_xlnm."

# Имена, зарезервированные Excel. В файле они обязаны быть с префиксом _xlnm.
BUILTIN_NAMES = (
    "_FilterDatabase",
    "Auto_Activate",
    "Auto_Close",
    "Auto_Deactivate",
    "Auto_Open",
    "Consolidate_Area",
    "Criteria",
    "Data_Form",
    "Database",
    "Extract",
    "Print_Area",
    "Print_Titles",
    "Recorder",
    "Sheet_Title",
)
_BUILTIN_LOWER = {name.lower(): name for name in BUILTIN_NAMES}

_DEFINED_NAME_RE = re.compile(r'<definedName\b[^>]*\bname="([^"]*)"', re.IGNORECASE)


class NameCheckError(Exception):
    """Книгу не удалось прочитать как zip — не наш случай, отдаём наверх."""


@dataclass(frozen=True)
class BadName:
    """Имя, из-за которого Excel спросит «Конфликт имён»."""

    name: str  # как записано в файле: _FilterDatabase
    builtin: str  # каноническое: _FilterDatabase
    count: int  # сколько раз встречается

    @property
    def fixed(self) -> str:
        return BUILTIN_PREFIX + self.builtin


def _read_workbook_xml(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as archive:
            return archive.read(WORKBOOK_PART).decode("utf-8", "replace")
    except (OSError, KeyError, zipfile.BadZipFile) as error:
        raise NameCheckError(f"{path.name}: {error}") from error


def scan(path: Path) -> list[BadName]:
    """Имена, совпадающие со встроенными, но записанные без префикса."""
    found: dict[str, int] = {}
    for raw in _DEFINED_NAME_RE.findall(_read_workbook_xml(path)):
        if raw.lower().startswith(BUILTIN_PREFIX):
            continue
        canonical = _BUILTIN_LOWER.get(raw.lower())
        if canonical:
            found[canonical] = found.get(canonical, 0) + 1
    return [
        BadName(name=name, builtin=name, count=count)
        for name, count in sorted(found.items())
    ]


def describe(problems: list[BadName]) -> list[str]:
    """Человеческое описание для окна с ошибкой."""
    lines = []
    for item in problems:
        times = "запись" if item.count == 1 else "записи"
        lines.append(
            f"«{item.name}» — {item.count} {times}; должно быть «{item.fixed}»"
        )
    return lines


def _patch_xml(xml: str, problems: list[BadName]) -> str:
    """Дописать префикс _xlnm. только в атрибутах name у definedName."""
    targets = {item.name.lower(): item.fixed for item in problems}

    def replace(match: re.Match[str]) -> str:
        whole = match.group(0)
        name = match.group(1)
        fixed = targets.get(name.lower())
        if not fixed:
            return whole
        return whole.replace(f'name="{name}"', f'name="{fixed}"', 1)

    return _DEFINED_NAME_RE.sub(replace, xml)


def lock_path(path: Path) -> Path:
    """Служебный файл Excel «~$книга.xlsm», который живёт рядом с открытой книгой."""
    return path.with_name("~$" + path.name)


def is_open_elsewhere(path: Path) -> bool:
    """Книга сейчас у кого-то открыта? Тогда трогать её мимо Excel нельзя.

    Признак не стопроцентный (бывают забытые ~$-файлы после падения Excel), но
    ошибается он в безопасную сторону: лишний раз не чиним.
    """
    return lock_path(path).exists()


def working_copy(source_path: Path, target_path: Path) -> list[BadName]:
    """Снять с книги рабочую копию, сразу с починенными именами.

    Для сравнения книга открывается не сама, а её копия во временной папке:
    оригинал при этом читается ровно один раз и не меняется вовсе. Если имена
    в порядке (или файл не zip — тогда о нём скажет Excel) — обычная копия.

    Копия временная, поэтому сжатие минимальное: пересборка 2-мегабайтной
    книги занимает 0,2 с вместо 0,5.
    """
    try:
        problems = scan(source_path)
    except NameCheckError:
        problems = []
    if not problems:
        shutil.copyfile(source_path, target_path)
        return []
    patched = _patch_xml(_read_workbook_xml(source_path), problems).encode("utf-8")
    _build_fixed_copy(source_path, target_path, patched, compresslevel=1)
    return problems


def _build_fixed_copy(
    source_path: Path,
    target_path: Path,
    patched: bytes,
    compresslevel: int | None = None,
) -> None:
    """Собрать рядом копию книги, где заменён только ``xl/workbook.xml``."""
    with (
        zipfile.ZipFile(source_path) as source,
        zipfile.ZipFile(
            target_path, "w", zipfile.ZIP_DEFLATED, compresslevel=compresslevel
        ) as target,
    ):
        for item in source.infolist():
            data = patched if item.filename == WORKBOOK_PART else source.read(item)
            # Сохраняем исходные атрибуты записи: дату, права, способ сжатия.
            info = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            info.compress_type = item.compress_type
            info.external_attr = item.external_attr
            info.internal_attr = item.internal_attr
            info.create_system = item.create_system
            target.writestr(info, data)


def _verify_copy(source_path: Path, fixed_path: Path) -> None:
    """Убедиться, что собранная копия — целая книга с тем же составом частей.

    Перезаписывать боевой файл содержимым, которое мы не перечитали, нельзя:
    ошибка здесь стоит дороже любой экономии времени.
    """
    with zipfile.ZipFile(source_path) as source, zipfile.ZipFile(fixed_path) as fixed:
        missing = set(source.namelist()) - set(fixed.namelist())
        if missing:
            raise NameCheckError(
                f"в починенной копии не хватает частей: {sorted(missing)[:3]}"
            )
        if not fixed.read(WORKBOOK_PART):
            raise NameCheckError("в починенной копии пустой " + WORKBOOK_PART)


def _overwrite(path: Path, fixed_path: Path) -> None:
    """Залить содержимое копии в существующий файл, не подменяя его.

    Заманчиво было бы просто переименовать копию поверх оригинала — это
    атомарно. Но боевая книга лежит в общей папке, и у файла бывают свои,
    назначенные не по наследству права: новый файл их не получит, и коллеги
    останутся без доступа. Поэтому пишем в тот же самый файл: inode, права и
    владелец остаются его собственными. Риск оборванной записи закрыт
    резервной копией, которую вызывающий код делает до правки.
    """
    with fixed_path.open("rb") as fixed, path.open("r+b") as original:
        shutil.copyfileobj(fixed, original)
        original.truncate()
        original.flush()
        os.fsync(original.fileno())


def repair(path: Path) -> list[BadName]:
    """Починить имена в книге на месте. Возвращает то, что исправлено.

    Переписывается ровно одна часть архива — ``xl/workbook.xml``; остальные,
    включая ``vbaProject.bin``, сводные кэши и элементы управления, копируются
    байт в байт. Это принципиально безопаснее пересохранения книги целиком.

    Чинит один запуск, а не файл навсегда: Excel при первом же сохранении
    записывает служебное имя обратно в кривом виде (замерено на боевой книге).
    """
    problems = scan(path)
    if not problems:
        return []

    patched = _patch_xml(_read_workbook_xml(path), problems).encode("utf-8")
    temporary = path.with_name(path.name + ".fixed")
    try:
        _build_fixed_copy(path, temporary, patched)
        _verify_copy(path, temporary)
        _overwrite(path, temporary)
    finally:
        with suppress(OSError):
            temporary.unlink()
    return problems
