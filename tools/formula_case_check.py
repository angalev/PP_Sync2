r"""Проверка: что программа делает с ячейками, где стоит формула.

    python -m tools.formula_case_check

Собирает пару книг через настоящий Excel (xlwings): нужен именно посчитанный
результат формулы, как в живом файле.

Заодно проверяется сопоставление столбцов по кодам: в «боевой» книге план и
факт лежат в ``DS``/``DT``, а в «копии» те же столбцы сдвинуты в ``DU``/``DV``
— как будто в копии на два столбца больше. Коды в строке 9 у них одинаковые
(311 и 312), и программа должна найти пары по ним, а не по буквам.

Формулы в ПП двух сортов, и проверяются оба.

**Расчётный столбец** (код 311 «Кол-во (план)» — он есть в
``Config.formula_codes``): формула держит лист, заменять её значением нельзя.

===========================  ==================================================
формула только в копии       статус «формула в копии», применить можно —
                             переносится результат расчёта
формула только в боевом      статус «формула в боевом», применить нельзя
формула с обеих сторон       строки нет: два расчёта сравнивать бессмысленно
===========================  ==================================================

**Обычный столбец** (код 312 «Кол-во (факт)» — расчётным не считается): формула
там — односложная ссылка на соседнюю ячейку, работник не стал впечатывать
цифру. Такую формулу надо заменить значением.

===========================  ==================================================
формула только в боевом      обычный статус «перезапись», применить можно,
                             ``replaces_formula`` поднят
формула с обеих сторон       строки нет: ссылка стоит у обоих, никто ничего
                             не впечатывал — разъехался тот столбец, на который
                             ссылаются
формул нет                   обычное «к применению» / «перезапись»
===========================  ==================================================
"""

from __future__ import annotations

import datetime
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppsync.config import Config
from ppsync.differ import build_diff
from ppsync.models import Status
from ppsync.values import code_text, display, is_empty_formula_result, is_error
from ppsync.workbook_reader import read_sheet

SHEET = "ПП 2026"
PLAN, FACT = "311", "312"  # коды столбцов «Кол-во (план)» и «Кол-во (факт)»

#: Где лежат столбцы в каждой книге: (ключ, план, факт). В копии всё правее
#: ключа сдвинуто на два столбца — буквы разные, коды одинаковые.
LAYOUT = {
    0: {"7": "A", "14": "F", "39": "U", PLAN: "DS", FACT: "DT"},
    1: {"7": "A", "14": "F", "39": "U", PLAN: "DU", FACT: "DV"},
}

#: (ключ, код столбца, ячейка боевого, ячейка копии, ожидаемый статус, ждём ли
#: замену формулы). ``{plan}``/``{fact}`` — буква соседнего столбца в своей
#: книге. ``None`` вместо статуса — строки в таблице быть не должно.
CASES: list[tuple[str, str, str, str, Status | None, bool]] = [
    # 311 — расчётный столбец: формулу в боевом не трогаем
    ("97575/1/000001", PLAN, "10", "={fact}{row}", Status.SOURCE_FORMULA, False),
    ("97575/1/000002", PLAN, "={fact}{row}", "40", Status.FORMULA, False),
    ("97575/1/000003", PLAN, "={fact}{row}", "={fact}{row}", None, False),
    ("97575/1/000004", PLAN, "10", "40", Status.OVERWRITE, False),
    # 312 — обычный столбец: формула в боевом это ленивая ссылка, заменяем
    ("97575/1/000005", FACT, "={plan}{row}", "40", Status.OVERWRITE, True),
    ("97575/1/000006", FACT, "10", "40", Status.OVERWRITE, False),
    # Тот самый случай с боевого файла: в боевом ячейка пустая, в копии
    # работник поставил ссылку вместо цифры. Раньше строка не появлялась
    # вовсе, и фактически выполненный объём не переносился.
    ("97575/1/000007", FACT, "", "={plan}{row}", Status.TO_APPLY, False),
    # И тут формулу трогать нельзя, хоть столбец и не расчётный: раз ссылка
    # стоит с обеих сторон, никто ничего не впечатывал — расходится план.
    ("97575/1/000008", FACT, "={plan}{row}", "={plan}{row}", None, False),
    # Формула копии посчитала ошибку — это не данные, переносить нечего.
    ("97575/1/000009", FACT, "10", "=1/0", None, False),
]

FIRST_ROW = 10

#: Опорные значения в «соседнем» столбце, на который ссылается формула.
#: Разные в боевом и в копии — иначе формула посчитает одно и то же, и
#: различия не будет вовсе.
ANCHOR_TARGET = 10
ANCHOR_SOURCE = 40


def _write_book(app, path: Path, *, side: int) -> None:
    """Сложить книгу с шапкой ПП и строками данных под каждый случай."""
    letters = LAYOUT[side]
    book = app.books.add()
    sheet = book.sheets[0]
    sheet.name = SHEET
    sheet.range("A6").value = "Общие сведения"
    sheet.range("A7").value = "Наименование трансгаза"
    sheet.range("U6").value = "Коды"
    sheet.range("U7").value = "Сводный код объекта ремонта (X/Y/Z)"
    sheet.range(f"{letters[PLAN]}6").value = "Физическое выполнение"
    sheet.range(f"{letters[PLAN]}7").value = "Кол-во (план)"
    sheet.range(f"{letters[FACT]}7").value = "Кол-во (факт)"
    for code, letter in letters.items():
        sheet.range(f"{letter}9").value = int(code)

    anchor = ANCHOR_TARGET if side == 0 else ANCHOR_SOURCE
    for offset, (key, code, target, source, _status, _flag) in enumerate(CASES):
        row = FIRST_ROW + offset
        sheet.range((row, 1)).value = "Волгоград ГТ"
        sheet.range((row, 6)).value = f"объект {offset + 1}"
        sheet.range(f"U{row}").value = key
        other = FACT if code == PLAN else PLAN
        sheet.range(f"{letters[other]}{row}").value = anchor
        cell = (target if side == 0 else source).format(
            row=row, plan=letters[PLAN], fact=letters[FACT]
        )
        if cell:  # пустая строка — ячейку не трогаем, она должна остаться пустой
            sheet.range(f"{letters[code]}{row}").value = cell

    book.api.SaveAs(str(path), 51)  # xlOpenXMLWorkbook
    book.close()


def main() -> int:
    import xlwings as xw

    workdir = Path(tempfile.mkdtemp(prefix="pp_formula_case_"))
    target_path = workdir / "боевой.xlsx"
    source_path = workdir / "копия.xlsx"

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    failures: list[str] = []
    try:
        _write_book(app, target_path, side=0)
        _write_book(app, source_path, side=1)

        config = Config()
        config.sheet_name = SHEET
        protected = config.protected_codes()
        print(f"Расчётные столбцы из настроек: {config.formula_codes}")
        print(f"код {PLAN} расчётный:", PLAN in protected)
        print(f"код {FACT} расчётный:", FACT in protected)
        print()

        target_book = app.books.open(str(target_path))
        source_book = app.books.open(str(source_path))
        try:
            target = read_sheet(target_book.sheets[SHEET], config, "боевой файл")
            source = read_sheet(source_book.sheets[SHEET], config, "копия")
        finally:
            source_book.close()
            target_book.close()
        rows = build_diff(source, target, config)

        print(
            f"Столбцы боевого: {PLAN} в {target.letter(PLAN)}, {FACT} в "
            f"{target.letter(FACT)}; копии: {PLAN} в {source.letter(PLAN)}, "
            f"{FACT} в {source.letter(FACT)}"
        )
        if (target.letter(PLAN), source.letter(PLAN)) != ("DS", "DU"):
            failures.append("столбцы не нашлись по кодам")

        found = {(r.key, r.code): r for r in rows}
        print(f"Строк данных в боевом: {target.count}, различий: {len(rows)}")
        print()
        for key, code, target_cell, source_cell, expected, wants_flag in CASES:
            row = found.get((key, code))
            actual = row.status if row else None
            flag = row.replaces_formula if row else False
            ok = actual == expected and flag == wants_flag
            if row is not None and ok:
                # Адреса обязаны указывать в свои книги: буквы у них разные
                letters = (row.column, row.source_column)
                ok = letters == (LAYOUT[0][code], LAYOUT[1][code])
            if not ok:
                failures.append(
                    f"код {code}: боевой={target_cell}, копия={source_cell} — ждали "
                    f"{expected} / замена формулы {wants_flag}, "
                    f"получили {actual} / {flag}"
                )
            shown = actual.value if actual else "строки нет"
            wanted = expected.value if expected else "строки нет"
            print(
                f"{'OK ' if ok else '!! '}код {code}: боевой={target_cell:<14} "
                f"копия={source_cell:<14} -> {shown:<18} (ждали: {wanted})"
            )
            if row is not None:
                print(
                    f"      {row.address} <- копия {row.source_address}: "
                    f"{row.current_display!r} -> {row.new_display!r}, "
                    f"применить {'можно' if row.applicable else 'нельзя'}, "
                    f"заменяет формулу: {row.replaces_formula}"
                )
                if row.note:
                    print(f"      {row.note}")
    finally:
        app.quit()
        shutil.rmtree(workdir, ignore_errors=True)

    print()
    # Формула, сосчитавшая ноль по пустой ссылке, — не значение. В ячейке с
    # форматом даты Excel показывает такой ноль как 00:00. На реальной паре
    # ПП это 143 из 300 «различий» в датах техакта.
    blanks = [
        (None, True),
        ("", True),
        ("   ", True),
        (datetime.time(0, 0), True),
        (datetime.datetime(1899, 12, 30), True),  # noqa: DTZ001 — Excel без TZ
        (datetime.time(9, 30), False),
        (datetime.datetime(2026, 8, 3), False),  # noqa: DTZ001 — Excel без TZ
        (0, False),
        ("б/н", False),
    ]
    for value, expected in blanks:
        actual = is_empty_formula_result(value)
        if actual != expected:
            failures.append(
                f"is_empty_formula_result({value!r}): ждали {expected}, "
                f"получили {actual}"
            )
        print(
            f"{'OK ' if actual == expected else '!! '}"
            f"результат формулы {value!r} — пусто: {actual}"
        )

    print()
    for left_value, right_value in ((24970662.1, 24970712.4), (4806.107, 4806.1071)):
        left, right = display(left_value), display(right_value)
        if left == right:
            failures.append(f"display: {left_value} и {right_value} склеились в {left}")
        print(
            f"{'OK ' if left != right else '!! '}"
            f"display({left_value})={left}, display({right_value})={right}"
        )

    print()
    for value, expected_code in ((39.0, "39"), (7, "7"), (" 12а ", "12а"), (None, "")):
        actual_code = code_text(value)
        if actual_code != expected_code:
            failures.append(f"code_text({value!r}): {actual_code!r}")
        print(
            f"{'OK ' if actual_code == expected_code else '!! '}"
            f"код столбца {value!r} -> {actual_code!r}"
        )
    for value, expected_flag in (("#DIV/0!", True), ("#REF!", True), ("б/н", False)):
        if is_error(value) != expected_flag:
            failures.append(f"is_error({value!r})")

    if failures:
        print()
        print("НЕ ПРОШЛО:")
        for line in failures:
            print(f"  • {line}")
        return 1
    print()
    print("Всё сошлось.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
