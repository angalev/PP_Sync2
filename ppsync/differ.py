"""Сравнение файла с изменениями и боевого файла.

Строки сопоставляются по сводному коду X/Y/Z, а не по номеру строки: в копии
могли добавиться или уехать строки, да и в боевом файле их порядок живой.

Столбцы сопоставляются по коду из строки кодов, а не по букве: копия и боевой
файл могут быть разной ширины (в ПП 2.0 удалено 40 столбцов). Сравниваются
только столбцы, код которых есть в обоих файлах.

Формулы в ПП двух разных сортов, и обращаться с ними надо по-разному.

**Расчётные столбцы** (``Config.formula_codes``: «Всего», «% выполнения»,
зеркальный блок «Работы по ДО»…) — там формула держит лист: сумма по строке,
процент выполнения, подстановка из парного столбца КОРР. Заменить её
значением — сломать расчёт, поэтому программа их не трогает.

**Все остальные столбцы.** Там формула — это односложная ссылка на соседнюю
ячейку (``=DS515``, ``=O152``, ``=DM11``): работник не стал впечатывать цифру.
Так быть не должно, и именно из-за этого фактически выполненные километры
(«Кол-во (факт)», 201 такая ячейка из 1278) не доезжали до боевого файла.
В таких столбцах формула считается обычным значением: сравнивается, показывается
и заменяется значением из копии — с пометкой в таблице и в журнале.

Само содержимое формулы не переносится никогда, ни в какую сторону: её ссылки
относятся к своему файлу. Переносится только результат.
"""

from __future__ import annotations

from openpyxl.utils import column_index_from_string

from .config import Config
from .models import DiffRow, Status
from .values import display, is_empty_formula_result, is_error
from .workbook_reader import RowData, SheetData, is_text_format

_TEXT_FORMAT_NOTE = "Формат ячейки «Текстовый» — значение запишется текстом."


def common_codes(source: SheetData, target: SheetData, config: Config) -> list[str]:
    """Коды столбцов, которые сравниваются: есть в обоих файлах и в настройках."""
    shared = [code for code in target.codes if code in source.columns]
    return config.code_scope(shared)


def build_diff(source: SheetData, target: SheetData, config: Config) -> list[DiffRow]:
    """Найти все различия между копией и боевым файлом."""
    rows: list[DiffRow] = []
    guarded = config.protected_codes()
    # Пары «столбец копии, столбец боевого» считаем один раз, а не на строку
    pairs = [
        (code, source.columns[code], target.columns[code], code in guarded)
        for code in common_codes(source, target, config)
    ]

    for key, source_row in source.rows.items():
        target_row = target.rows.get(key)

        for code, source_column, target_column, is_guarded in pairs:
            new_value = source_row.value(source_column)
            from_formula = bool(source_row.formula(source_column))

            if target_row is None:
                if display(new_value) and not from_formula:
                    rows.append(_missing_row(code, source, target, source_row))
                continue

            if from_formula and is_empty_formula_result(new_value):
                # Формула копии посчитала ноль (ссылка на пустую ячейку). Это
                # не данные: в ячейке с форматом даты такой ноль выглядит как
                # «00:00». Ни переносить, ни очищать им боевой файл нельзя.
                continue
            if is_error(new_value):
                # #REF!, #DIV/0! и прочее — тоже не данные: переносить в
                # боевой файл текст ошибки незачем.
                continue

            target_formula = target_row.formula(target_column)
            if from_formula and target_formula:
                # Формула с обеих сторон — синхронизировать нечего: ни там, ни
                # там человек ничего не впечатывал, оба значения посчитаны. Это
                # верно и для расчётных столбцов, и для обычных: если в «Кол-во
                # (факт)» с обеих сторон стоит ссылка на план, то разъехался
                # план, а не факт.
                continue

            current_value = target_row.value(target_column)
            current_display = display(current_value)
            new_display = display(new_value)
            if current_display == new_display:
                continue

            if target_formula and is_guarded:
                status = Status.FORMULA
            elif from_formula and is_guarded:
                status = Status.SOURCE_FORMULA
            else:
                status = _status(current_display, new_display)

            rows.append(
                DiffRow(
                    key=key,
                    column=target.letter(code),
                    field_name=_field_name(code, target),
                    status=status,
                    code=code,
                    source_column=source.letter(code),
                    row=target_row.row,
                    source_row=source_row.row,
                    transgaz=target_row.transgaz or source_row.transgaz,
                    object_name=target_row.object_name,
                    current_value=current_value,
                    new_value=new_value,
                    current_display=current_display,
                    new_display=new_display,
                    replaces_formula=bool(target_formula) and not is_guarded,
                    note=_formula_note(from_formula, target_formula, is_guarded),
                )
            )

    rows.sort(key=lambda r: (r.transgaz, r.row or 0, r.column))
    # Ничего не отмечаем заранее. Отметить всё «за человека» опасно вдвойне:
    # запись идёт в боевой файл, а фильтр по трансгазам строки только прячет —
    # человек видит в таблице одну галочку и не подозревает про остальные.
    # Отмечает он сам: галочками или кнопкой «Отметить все в таблице».
    return rows


def _field_name(code: str, sheet: SheetData) -> str:
    """Имя столбца для человека: код впереди — по нему столбец и ищут."""
    name = sheet.headers.get(code, "")
    return f"[{code}] {name}".strip()


def _formula_note(from_formula: bool, target_formula: str, guarded: bool) -> str:
    """Пояснение про формулы для строки таблицы."""
    if target_formula and guarded:
        return f"В боевом файле формула: {target_formula}"
    parts: list[str] = []
    if from_formula:
        parts.append(
            "В копии это значение посчитано формулой. Переносится результат, "
            "сама формула остаётся в копии."
        )
    if target_formula:
        parts.append(
            f"В боевом файле здесь стоит формула {target_formula} вместо "
            "значения — она будет заменена."
        )
    return " ".join(parts)


def _status(current_display: str, new_display: str) -> Status:
    if not new_display:
        return Status.CLEAR
    if current_display:
        return Status.OVERWRITE
    return Status.TO_APPLY


def _missing_row(
    code: str, source: SheetData, target: SheetData, source_row: RowData
) -> DiffRow:
    new_value = source_row.value(source.columns[code])
    return DiffRow(
        key=source_row.key,
        column=target.letter(code),
        field_name=_field_name(code, target),
        status=Status.KEY_NOT_FOUND,
        code=code,
        source_column=source.letter(code),
        source_row=source_row.row,
        transgaz=source_row.transgaz,
        object_name=source_row.object_name,
        new_value=new_value,
        new_display=display(new_value),
        note=(
            f"Код {source_row.key} есть в копии (строка {source_row.row}), "
            "но не найден в боевом файле. Новые строки программа не создаёт."
        ),
    )


def apply_column_formats(rows: list[DiffRow], target: SheetData) -> None:
    """Проставить формат по столбцу — там, где он у столбца один.

    Этого хватает, чтобы ещё в таблице предупредить про столбцы с форматом
    «Текстовый». Точный формат каждой ячейки дочитывается перед записью.
    """
    for row in rows:
        number_format = target.formats.get(row.code)
        if row.row and number_format:
            _set_format(row, number_format)


def apply_formats(rows: list[DiffRow], formats: dict[tuple[int, int], str]) -> None:
    """Проставить формат целевых ячеек и предупредить про «Текстовый».

    ``formats`` — из :func:`workbook_reader.read_formats`, ключ — ``(строка,
    номер столбца)`` боевого файла.
    """
    for row in rows:
        if not row.row:
            continue
        number_format = formats.get((row.row, column_index_from_string(row.column)))
        if number_format is not None:
            _set_format(row, number_format)


def _set_format(row: DiffRow, number_format: str) -> None:
    row.number_format = number_format
    row.text_format = is_text_format(number_format)
    if (
        row.text_format
        and row.applicable
        and not isinstance(row.new_value, str)
        and _TEXT_FORMAT_NOTE not in row.note
    ):
        row.note = (row.note + " " if row.note else "") + _TEXT_FORMAT_NOTE


def format_targets(rows: list[DiffRow]) -> list[tuple[int, int]]:
    """Ячейки боевого файла, формат которых надо знать: то, что можно записать."""
    return [
        (row.row, column_index_from_string(row.column))
        for row in rows
        if row.row and row.applicable
    ]


def summarize(rows: list[DiffRow]) -> dict[Status, int]:
    """Сколько строк в каждом статусе."""
    counts: dict[Status, int] = {}
    for row in rows:
        counts[row.status] = counts.get(row.status, 0) + 1
    return counts


def group_by_transgaz(rows: list[DiffRow]) -> dict[str, list[DiffRow]]:
    """Группировка для левой панели GUI (по убыванию количества)."""
    groups: dict[str, list[DiffRow]] = {}
    for row in rows:
        groups.setdefault(row.transgaz or "(без трансгаза)", []).append(row)
    return dict(sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0])))
