r"""Починка служебных имён книги: _FilterDatabase -> _xlnm._FilterDatabase.

    python -m tools.fix_names "путь\к\книге.xlsm"          # только показать
    python -m tools.fix_names "путь\к\книге.xlsm" --repair  # починить

Переписывается ровно одна часть архива (``xl/workbook.xml``), остальные —
включая макросы, сводные и элементы управления — копируются байт в байт.
Рядом кладётся копия ``*.before_fix.xlsm``.

Важно понимать границы: Excel записывает это имя обратно в кривом виде при
первом же сохранении книги в общем доступе — проверено замером. Поэтому починка
лечит один конкретный запуск, а не файл навсегда, и запускать её имеет смысл
тогда, когда Excel показал окно «Конфликт имен».

Файл не должен быть открыт ни у кого: правка идёт мимо Excel, и запись под
чужой открытой сессией общего доступа сломает её слияние.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppsync import name_check


def main(argv: list[str]) -> int:
    args = [a for a in argv[1:] if not a.startswith("--")]
    repair = "--repair" in argv
    if not args:
        print(__doc__)
        return 2

    path = Path(args[0])
    if not path.exists():
        print(f"Файл не найден: {path}")
        return 1

    lock = path.with_name("~$" + path.name)
    if lock.exists():
        print(f"ВНИМАНИЕ: рядом лежит {lock.name} — книга у кого-то открыта.")
        if repair:
            print("Чинить открытую книгу нельзя. Закройте её и повторите.")
            return 1

    try:
        problems = name_check.scan(path)
    except name_check.NameCheckError as error:
        print(f"Не удалось прочитать книгу: {error}")
        return 1

    if not problems:
        print(f"{path.name}: служебные имена в порядке, чинить нечего.")
        return 0

    print(f"{path.name}: найдено то, из-за чего Excel спрашивает «Конфликт имен»:")
    for line in name_check.describe(problems):
        print(f"  • {line}")

    if not repair:
        print("\nЗапустите с ключом --repair, чтобы починить.")
        return 1

    backup = path.with_name(path.stem + ".before_fix" + path.suffix)
    shutil.copy2(path, backup)
    fixed = name_check.repair(path)
    print(f"\nКопия до правки: {backup}")
    print("Исправлено:", ", ".join(f"{b.name} -> {b.fixed}" for b in fixed))
    print("Проверка после правки:", name_check.scan(path) or "имена в порядке")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
