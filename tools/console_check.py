"""Консольная проверка сравнения без записи.

    python -m tools.console_check --target "tests/data/ПП_боевой.xlsx" \
                                  --source "tests/data/ПП_копия_с_изменениями.xlsx"

Делает ровно то же, что кнопка «Сравнить»: читает оба файла с рабочих копий и
печатает таблицу различий с временем каждого этапа. Ни один из файлов не
меняется — ни боевой, ни копия.
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppsync.applier import ApplyError, SyncSession
from ppsync.config import Config
from ppsync.differ import summarize
from ppsync.workbook_reader import WorkbookStructureError


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Сравнение двух копий ПП")
    parser.add_argument("--target", help="боевой файл (по умолчанию — из config.json)")
    parser.add_argument("--source", help="файл с изменениями")
    parser.add_argument("--sheet", help="имя листа")
    parser.add_argument(
        "--codes", help="коды столбцов, которые сравнивать, например 259-273, 316"
    )
    parser.add_argument("--config", help="путь к config.json")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = Config.load(args.config)
    if args.target:
        config.target_file = str(Path(args.target).resolve())
    if args.source:
        config.source_file = str(Path(args.source).resolve())
    if args.sheet:
        config.sheet_name = args.sheet
    if args.codes:
        config.compare_codes = args.codes

    print(f"Боевой : {config.target_file}")
    print(f"Копия  : {config.source_file}")
    print(f"Лист   : {config.sheet_name}")
    problems = config.validate()
    if problems:
        for problem in problems:
            print(f"  !! {problem}")
        return 2

    started = time.perf_counter()
    last = [started]

    def progress(text: str) -> None:
        now = time.perf_counter()
        print(f"  {now - last[0]:6.2f} с  {text}", flush=True)
        last[0] = now

    try:
        compare = SyncSession(config, progress=progress).compare()
    except (ApplyError, WorkbookStructureError) as error:
        print(f"\nНе удалось сравнить: {error.full_text()}")
        return 1
    print(f"Всего: {time.perf_counter() - started:.2f} с")
    print(
        f"Строк: в боевом {compare.target.count}, в копии {compare.source.count}; "
        f"столбцов в сравнении: {len(compare.codes)}"
    )
    for warning in compare.warnings:
        print(f"  ! {warning}")

    print_table(compare.rows)
    return 0


def print_table(rows: list) -> None:
    widths = (18, 18, 26, 8, 16, 16, 20)
    header = ("Трансгаз", "XYZ", "Столбец", "Ячейка", "В боевом", "В копии", "Статус")
    print()
    print(
        " | ".join(
            text[:width].ljust(width)
            for text, width in zip(header, widths, strict=True)
        )
    )
    print("-+-".join("-" * width for width in widths))
    for row in rows[:60]:
        cells = (
            row.transgaz,
            row.key,
            row.field_name,
            row.address or f"копия {row.source_address}",
            row.current_display or "—",
            row.new_display or "(очистить)",
            row.status.value,
        )
        print(
            " | ".join(str(t)[:w].ljust(w) for t, w in zip(cells, widths, strict=True))
        )
    if len(rows) > 60:
        print(f"… и ещё {len(rows) - 60}")

    print("\nИтого:")
    for status, count in summarize(rows).items():
        print(f"  {status.value:<22} {count}")
    notes = [r for r in rows if r.note]
    if notes:
        print("\nПояснения:")
        for row in notes[:10]:
            print(f"  • {row.key} {row.column}: {row.note}")


if __name__ == "__main__":
    raise SystemExit(main())
