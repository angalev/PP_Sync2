r"""Проверка починки служебных имён — без Excel.

    python -m tools.name_fix_check

Собирает тестовую книгу, портит в ней имя автофильтра ровно так, как это
выглядит в боевом файле (``_FilterDatabase`` без префикса ``_xlnm.``), и
проверяет три вещи, ради которых починка и написана:

* ``scan`` находит проблему, а ``repair`` её убирает;
* остальные части архива не изменились байт в байт — макросы, сводные и
  элементы управления должны пережить правку нетронутыми;
* файл остался тем же самым файлом, а не подменённым: на общей папке у книги
  бывают свои права доступа, и подмена файла оставила бы коллег без доступа.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppsync import name_check
from tools import make_test_workbook

failures: list[str] = []


def _break_name(path: Path) -> None:
    """Записать в книгу имя автофильтра без обязательного префикса _xlnm."""
    broken = path.with_name("broken_" + path.name)
    with (
        zipfile.ZipFile(path) as source,
        zipfile.ZipFile(broken, "w", zipfile.ZIP_DEFLATED) as target,
    ):
        for item in source.infolist():
            data = source.read(item)
            if item.filename == name_check.WORKBOOK_PART:
                xml = data.decode("utf-8")
                injected = (
                    '<definedNames><definedName name="_FilterDatabase" '
                    'localSheetId="0" hidden="1">Лист1!$A$1:$B$2</definedName>'
                    "</definedNames>"
                )
                xml = xml.replace("<sheets>", injected + "<sheets>", 1)
                data = xml.encode("utf-8")
            target.writestr(item, data)
    shutil.move(str(broken), str(path))


def _parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {item.filename: archive.read(item) for item in archive.infolist()}


def main() -> int:
    sandbox = Path(tempfile.mkdtemp(prefix="pp_sync_names_"))
    book = sandbox / "ПП_боевой.xlsx"
    make_test_workbook.build(book)
    _break_name(book)

    problems = name_check.scan(book)
    if [item.name for item in problems] != ["_FilterDatabase"]:
        failures.append(f"scan не нашёл сломанное имя: {problems}")

    before = _parts(book)
    identity_before = book.stat().st_ino if hasattr(book.stat(), "st_ino") else None

    fixed = name_check.repair(book)
    if [item.fixed for item in fixed] != ["_xlnm._FilterDatabase"]:
        failures.append(f"repair вернул неожиданное: {fixed}")
    if name_check.scan(book):
        failures.append("после починки имя всё ещё сломано")

    after = _parts(book)
    if set(before) != set(after):
        failures.append("состав частей книги изменился")
    for part, data in before.items():
        if part == name_check.WORKBOOK_PART:
            continue
        if after.get(part) != data:
            failures.append(f"часть {part} изменилась, хотя не должна была")

    identity_after = book.stat().st_ino if identity_before is not None else None
    if identity_before is not None and identity_before != identity_after:
        failures.append("файл подменён другим — права доступа потерялись")

    if (sandbox / (book.name + ".fixed")).exists():
        failures.append("временная копия .fixed осталась на диске")

    # Повторная починка на здоровой книге ничего не делает
    if name_check.repair(book):
        failures.append("repair сработал на книге, где чинить нечего")

    shutil.rmtree(sandbox, ignore_errors=True)

    if failures:
        print("ПРОВЕРКА НЕ ПРОЙДЕНА:")
        for line in failures:
            print(f"  • {line}")
        return 1
    print("Починка служебных имён: проверка пройдена.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
