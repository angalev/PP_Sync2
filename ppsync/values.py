"""Нормализация значений.

Две задачи:

* ``display`` — привести любое значение (из Excel через COM или из пакета через
  openpyxl) к строке. Эта же строка используется для сравнения «было/станет»,
  поэтому дата 30.06.2026 из пакета и ``datetime(2026, 6, 30)`` из Excel должны
  давать один и тот же результат — иначе differ будет вечно показывать
  «перезапись» там, где ничего не меняется.
* ``coerce_for_write`` — превратить значение из пакета в тип, который имеет смысл
  писать в ячейку: дату как дату, число как число, остальное как текст.
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import Any

#: dd.mm.yyyy / dd-mm-yyyy / dd/mm/yyyy, опционально с временем
_DATE_RE = re.compile(
    r"^(?P<d>\d{1,2})[.\-/](?P<m>\d{1,2})[.\-/](?P<y>\d{4})"
    r"(?:\s+(?P<H>\d{1,2}):(?P<M>\d{2})(?::(?P<S>\d{2}))?)?$"
)
#: yyyy-mm-dd, опционально с временем
_ISO_RE = re.compile(
    r"^(?P<y>\d{4})-(?P<m>\d{1,2})-(?P<d>\d{1,2})"
    r"(?:[ T](?P<H>\d{1,2}):(?P<M>\d{2})(?::(?P<S>\d{2}))?)?$"
)
#: число: 1234, 12,5, -3.75 (без ведущих нулей — 007 это код, а не число)
_NUM_RE = re.compile(r"^-?(?:0|[1-9]\d*)(?:[.,]\d+)?$")

_MARKER_EMPTY = "<ПУСТО>"


def is_empty_marker(value: Any) -> bool:
    """Маркер осознанной очистки ячейки."""
    return isinstance(value, str) and value.strip().upper() == _MARKER_EMPTY.upper()


def is_empty_formula_result(value: Any) -> bool:
    """Формула вернула «ничего», а не значение.

    Пустую ссылку Excel считает нулём. В ячейке с форматом даты этот ноль
    выглядит как 00.01.1900, и openpyxl отдаёт его как ``time(0, 0)`` или как
    дату раньше 1900 года. Значение это не данные, а артефакт расчёта.

    Замер на реальной паре ПП: в копии столбцы DZ и EA («Дата начала/окончания
    работ, указанная в техакте») заполнены формулами ``=DO``/``=DP``. Там, где
    оперативные даты ещё не проставлены, формула даёт ноль — и без этой
    проверки в боевой файл уехало бы 143 фальшивых «00:00».

    Отсюда же правило: формула в копии может только добавить или изменить
    значение, но никогда не очищает ячейку боевого файла.
    """
    if value is None:
        return True
    if isinstance(value, str):
        return not clean_str(value)
    if isinstance(value, _dt.datetime):
        return value.year < 1900
    if isinstance(value, _dt.time):
        return (value.hour, value.minute, value.second) == (0, 0, 0)
    if isinstance(value, _dt.date):
        return value.year < 1900
    return False


#: Ошибки Excel так, как их отдаёт COM: целым числом вместо значения.
#: Обычные числа приходят через COM как float, поэтому путаницы нет.
EXCEL_ERRORS: dict[int, str] = {
    -2146826288: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
}
_ERROR_TEXTS = frozenset(EXCEL_ERRORS.values())


def from_excel(value: Any) -> Any:
    """Значение ячейки из COM -> значение для сравнения.

    Единственное, что здесь правится, — ошибки: ``-2146826281`` в таблице
    различий никому ничего не скажет, а ``#DIV/0!`` скажет.
    """
    if isinstance(value, int) and not isinstance(value, bool):
        return EXCEL_ERRORS.get(value, value)
    return value


def is_error(value: Any) -> bool:
    """Ячейка посчитала ошибку (#REF!, #DIV/0!…). Это не данные."""
    return isinstance(value, str) and value in _ERROR_TEXTS


def code_text(value: Any) -> str:
    """Код столбца из строки кодов -> строка.

    Excel отдаёт число как ``39.0``; человек в настройках пишет ``39``.
    Читать строку кодов надо через ``Value2``: ячейка с кодом 21 может стоять
    в формате даты, и тогда ``Value`` вернёт 21.01.1900 вместо числа.
    """
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else format_float(value)
    if isinstance(value, int):
        return str(value)
    return clean_str(str(value))


def looks_like_key(text: Any) -> bool:
    """Похоже ли значение на ключ строки.

    Формат ключей в боевом файле неоднороден: рядом с «97575/12/102770»
    встречается «97435.8-9100297753» (в ПП 2026 таких 86 строк). Поэтому
    ключом считается любая непустая строка, кроме чистого числа — числа лежат
    в служебной строке с номерами колонок, и путать их с ключами нельзя.
    """
    if text is None:
        return False
    if isinstance(text, (int, float)) and not isinstance(text, bool):
        return False
    value = clean_str(str(text))
    if not value:
        return False
    return parse_number(value) is None


def clean_str(value: str) -> str:
    """Убрать неразрывные пробелы и схлопнуть повторяющиеся."""
    return re.sub(r"\s+", " ", value.replace("\xa0", " ")).strip()


#: Сколько значащих цифр оставляем у дробного числа.
#:
#: 12 — это компромисс. Меньше нельзя: у восьмизначной суммы вроде
#: 24 970 662,10 руб. шесть значащих цифр (старое ``%g``) съедали всё после
#: сотен, и различие в рублях программа считала совпадением. Больше не нужно:
#: 3.6320000000000001 из одного файла и 3.632 из другого — это одно и то же
#: число, и разъезжаться из-за двоичного мусора они не должны.
FLOAT_DIGITS = 12


def format_float(value: float) -> str:
    """Дробное число строкой — без экспоненты и без потери значащих цифр.

    ``%g`` с шестью цифрами по умолчанию делал сразу две беды: показывал
    24 970 662,1 как ``2.49707e+07`` и приравнивал к нему любое число, у
    которого совпали первые шесть цифр.
    """
    text = f"{value:.{FLOAT_DIGITS}g}"
    if "e" in text or "E" in text:
        # Очень мелкие (1e-07) и очень крупные числа %g уводит в экспоненту —
        # разворачиваем обратно, иначе в таблице различий это не прочитать.
        text = f"{value:.{FLOAT_DIGITS}f}".rstrip("0").rstrip(".")
        return text or "0"
    return text


def display(value: Any) -> str:
    """Каноническая строка для показа и сравнения."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "ИСТИНА" if value else "ЛОЖЬ"
    if isinstance(value, _dt.datetime):
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.strftime("%d.%m.%Y")
        return value.strftime("%d.%m.%Y %H:%M")
    if isinstance(value, _dt.date):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, _dt.time):
        return value.strftime("%H:%M")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return format_float(value)
    if isinstance(value, int):
        return str(value)

    text = clean_str(str(value))
    if is_empty_marker(text):
        return ""
    # Текст, который на самом деле дата ("30.06.2026") — приводим к тому же виду,
    # что и datetime из Excel.
    parsed = parse_date(text)
    if parsed is not None:
        return display(parsed)
    return text


def parse_date(text: str) -> _dt.datetime | None:
    """Разобрать дату из строки. Возвращает None, если это не дата."""
    text = clean_str(text)
    for pattern in (_DATE_RE, _ISO_RE):
        m = pattern.match(text)
        if not m:
            continue
        try:
            # Excel хранит даты без таймзоны — naive datetime здесь намеренный.
            return _dt.datetime(  # noqa: DTZ001
                int(m.group("y")),
                int(m.group("m")),
                int(m.group("d")),
                int(m.group("H") or 0),
                int(m.group("M") or 0),
                int(m.group("S") or 0),
            )
        except ValueError:
            return None
    return None


def parse_number(text: str) -> float | int | None:
    """Разобрать число из строки. Коды с ведущими нулями числом не считаются."""
    text = clean_str(text).replace(" ", "")
    if not _NUM_RE.match(text):
        return None
    number = float(text.replace(",", "."))
    return int(number) if number.is_integer() else number


def coerce_for_write(value: Any, *, coerce_types: bool = True) -> Any:
    """Значение из пакета -> значение для записи в ячейку.

    ``coerce_types=False`` — писать ровно то, что пришло (кроме маркера очистки).
    """
    if value is None or is_empty_marker(value):
        return None
    if not isinstance(value, str):
        return value
    text = clean_str(value)
    if not text:
        return None
    if not coerce_types:
        return text
    parsed_date = parse_date(text)
    if parsed_date is not None:
        return parsed_date
    parsed_number = parse_number(text)
    if parsed_number is not None:
        return parsed_number
    return text
