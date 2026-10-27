"""Сессия: сравнить копию с боевым файлом, записать отмеченное.

Работа разбита на два независимых шага, и Excel открыт только пока идёт шаг.

**Сравнение** боевой файл не трогает вообще:

1. с каждого файла снимается рабочая копия во временную папку — оригинал
   читается один раз, по сети ходит ровно один файл в одну сторону;
2. в копии чинится служебное имя автофильтра (иначе Excel спросит «Конфликт
   имен») — в копии, а не в боевом файле;
3. копия открывается в скрытом Excel только на чтение, лист читается
   прямоугольником: значения и формулы;
4. Excel закрывается, временная папка удаляется.

Оба файла проходят эти шаги одновременно, каждый в своём потоке и в своём
экземпляре Excel: открытие книги — самая долгая часть, и ждать её дважды
подряд незачем. Различия считаются уже в памяти. Пока человек расставляет
галочки, боевой файл никем не открыт.

**Запись** — единственное место, где боевой файл меняется:

1. резервная копия — прямо перед записью, а не при сравнении;
2. боевой открывается в Excel (общий доступ, ``ConflictResolution``);
3. лист перечитывается заново: каждая отмеченная ячейка сверяется с тем, что
   человек видел в таблице. Если коллега успел её изменить — ячейка не
   пишется, о ней сообщается отдельно;
4. запись отмеченного, пересчёт, ``save()`` (слияние с правками коллег),
   перечитывание записанного, закрытие Excel, журнал.
"""

from __future__ import annotations

import datetime as _dt
import shutil
import tempfile
import threading
from collections.abc import Callable, Iterator
from concurrent.futures import Future, ThreadPoolExecutor
from contextlib import contextmanager, suppress
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl.utils import column_index_from_string, get_column_letter

from . import name_check
from .config import (
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
    Config,
    DataDirError,
    explain_path_error,
)
from .differ import (
    apply_column_formats,
    apply_formats,
    build_diff,
    common_codes,
    format_targets,
    summarize,
)
from .excel_session import ExcelError, ExcelInstance
from .logger import Journal
from .models import APPLICABLE, DiffRow, Status
from .values import coerce_for_write, display
from .workbook_reader import (
    SheetData,
    WorkbookStructureError,
    read_cells,
    read_formats,
    read_sheet,
)

Progress = Callable[[str], None]

TARGET_TITLE = "боевой файл"
SOURCE_TITLE = "файл с изменениями"

#: Как часто сообщать о ходе записи. Сообщение на каждую ячейку — это тысячи
#: строк в журнале окна, и рисуются они дольше, чем идёт сама запись.
WRITE_REPORT_EVERY = 50


class ApplyError(Exception):
    """Ошибка, которую нужно показать пользователю."""

    def __init__(self, message: str, problems: list[str] | None = None) -> None:
        super().__init__(message)
        self.problems = problems or []

    def full_text(self) -> str:
        if not self.problems:
            return str(self)
        return str(self) + "\n\n" + "\n".join(f"• {p}" for p in self.problems)


@dataclass
class CompareResult:
    """Что показать пользователю после «Сравнить»."""

    rows: list[DiffRow]
    target: SheetData
    source: SheetData
    codes: list[str]  # коды сравнивавшихся столбцов
    warnings: list[str] = field(default_factory=list)

    def counts(self) -> dict[Status, int]:
        return summarize(self.rows)


@dataclass
class ApplyResult:
    """Итог применения."""

    applied: int = 0
    failed: int = 0
    stale: int = 0  # изменилось в боевом после сравнения — не записано
    skipped: int = 0
    backup_path: Path | None = None
    journal_path: Path | None = None
    errors: list[str] = field(default_factory=list)
    verify_warnings: list[str] = field(default_factory=list)

    def summary(self) -> str:
        parts = [f"применено: {self.applied}"]
        if self.failed:
            parts.append(f"ошибок: {self.failed}")
        if self.stale:
            parts.append(f"изменилось в боевом файле и не записано: {self.stale}")
        if self.skipped:
            parts.append(f"не отмечено: {self.skipped}")
        return ", ".join(parts)


def _stamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d_%H%M")  # noqa: DTZ005


def _unique_path(path: Path) -> Path:
    """Не затирать уже лежащий файл: добавить _2, _3 …"""
    if not path.exists():
        return path
    for number in range(2, 100):
        candidate = path.with_name(f"{path.stem}_{number}{path.suffix}")
        if not candidate.exists():
            return candidate
    return path.with_name(f"{path.stem}_{_dt.datetime.now():%S%f}{path.suffix}")  # noqa: DTZ005


def _listing(sheet: SheetData, codes: list[str], limit: int = 6) -> str:
    """«8 (B, Примечание), 42 (V, …) и ещё 34» — столбцы для предупреждения."""
    text = ", ".join(sheet.describe(code) for code in codes[:limit])
    if len(codes) > limit:
        text += f" и ещё {len(codes) - limit}"
    return text


def _duplicate_lines(sheet: SheetData) -> list[str]:
    lines = [
        f"{key} — строки {', '.join(map(str, rows))}"
        for key, rows in list(sheet.duplicates.items())[:10]
    ]
    if len(sheet.duplicates) > 10:
        lines.append(f"… и ещё {len(sheet.duplicates) - 10}")
    lines.append("По повторяющемуся ключу строку не найти однозначно — уберите дубли.")
    return lines


class SyncSession:
    """Одна пара файлов: сравнение и, если человек решил, запись."""

    def __init__(self, config: Config, progress: Progress | None = None) -> None:
        self.config = config
        self.progress = progress or (lambda _message: None)
        self.rows: list[DiffRow] = []
        self._running: list[ExcelInstance] = []
        self._lock = threading.Lock()

    @property
    def excel_pids(self) -> list[int]:
        """PID работающих сейчас экземпляров Excel — для сторожа окон в GUI."""
        with self._lock:
            return [excel.pid for excel in self._running if excel.pid]

    @contextmanager
    def _excel(self) -> Iterator[ExcelInstance]:
        """Запустить Excel и гарантированно закрыть его — в этом же потоке."""
        excel = ExcelInstance(self.config, self.progress)
        with self._lock:
            self._running.append(excel)
        try:
            excel.start()
            yield excel
        finally:
            excel.close()
            with self._lock:
                self._running.remove(excel)

    # --- шаг 1: сравнение ---------------------------------------------------

    def compare(self) -> CompareResult:
        """Прочитать оба файла (параллельно, с рабочих копий) и собрать различия."""
        config = self.config
        try:
            warnings: list[str] = list(config.ensure_dirs())
        except DataDirError as error:
            raise ApplyError(
                str(error),
                ["Меню «Настройки» → «Папка бэкапов и журнала»."],
            ) from error

        workdir = Path(tempfile.mkdtemp(prefix="pp_sync_"))
        try:
            self.progress(
                f"Читаю оба файла: {config.target_path.name} (боевой) и "
                f"{config.source_path.name} (копия)…"
            )
            with ThreadPoolExecutor(max_workers=4) as pool:
                jobs = []
                # Имена рабочих копий разные намеренно: боевой файл и копия
                # обычно называются одинаково.
                for original, name, title in (
                    (config.target_path, "target", TARGET_TITLE),
                    (config.source_path, "source", SOURCE_TITLE),
                ):
                    # Копия снимается, пока запускается Excel: ждать одно,
                    # чтобы начать другое, незачем.
                    working = workdir / f"{name}{original.suffix}"
                    copy = pool.submit(self._working_copy, original, working)
                    jobs.append(pool.submit(self._read_file, copy, original, title))
                # Выход из блока ждёт все потоки: Excel закрывается в них самих.
            target, source = (job.result() for job in jobs)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)

        if target.duplicates:
            raise WorkbookStructureError(
                "В боевом файле повторяются ключи X/Y/Z", _duplicate_lines(target)
            )
        codes = common_codes(source, target, config)
        if not codes:
            raise ApplyError(
                "Нечего сравнивать: у файлов нет общих столбцов.",
                [
                    (
                        "Столбцы сопоставляются по кодам из строки "
                        f"{config.code_row}. Проверьте, что оба файла — "
                        "Производственная программа, и настройки «сравнивать "
                        "столбцы» / «не сравнивать столбцы»."
                    ),
                ],
            )
        warnings.extend(self._compare_warnings(source, target))

        self.progress(
            f"Сравниваю: {len(codes)} столбцов, строк в боевом {target.count}, "
            f"в копии {source.count}…"
        )
        self.rows = build_diff(source, target, config)
        apply_column_formats(self.rows, target)

        self.progress(f"Готово: найдено различий — {len(self.rows)}.")
        return CompareResult(
            rows=self.rows,
            target=target,
            source=source,
            codes=codes,
            warnings=warnings,
        )

    def _working_copy(self, original: Path, working: Path) -> Path:
        """Снять рабочую копию файла (с починенным служебным именем).

        Оригинал только читается: копию можно снять, даже если файл сейчас
        открыт у человека в его собственном Excel.
        """
        try:
            if self.config.fix_service_names:
                name_check.working_copy(original, working)
            else:
                shutil.copyfile(original, working)
        except OSError as error:
            raise ApplyError(
                f"Не удалось прочитать файл {original.name}",
                [
                    explain_path_error(original, error),
                    "Проверьте, подключён ли сетевой диск и открыт ли к нему доступ.",
                ],
            ) from error
        return working

    def _read_file(self, copy: Future[Path], original: Path, title: str) -> SheetData:
        """Excel -> рабочая копия -> лист в памяти. Целиком в одном потоке."""
        import pythoncom  # из pywin32, ставится вместе с xlwings

        pythoncom.CoInitialize()  # у каждого потока свой COM
        try:
            with self._excel() as excel:
                sheet = self._open(excel, copy.result(), original)
                return read_sheet(
                    sheet, self.config, title, with_formats=title == TARGET_TITLE
                )
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _open(
        excel: ExcelInstance, path: Path, shown: Path | None = None, **kwargs: Any
    ) -> Any:
        try:
            return excel.open_sheet(path, shown, **kwargs)
        except ExcelError as error:
            raise ApplyError(str(error), error.problems) from error

    def _compare_warnings(self, source: SheetData, target: SheetData) -> list[str]:
        """То, о чём стоит сказать, но из-за чего не надо останавливаться."""
        warnings: list[str] = []
        if source.duplicates:
            listing = ", ".join(list(source.duplicates)[:5])
            warnings.append(
                f"В копии {len(source.duplicates)} повторяющихся кодов X/Y/Z "
                f"({listing}) — учтено только первое вхождение каждого."
            )
        missing_rows = sum(1 for key in target.rows if key not in source.rows)
        if missing_rows:
            warnings.append(
                f"В копии нет {missing_rows} строк из боевого файла — "
                "они просто не сравнивались."
            )
        only_source = [code for code in source.codes if code not in target.columns]
        if only_source:
            warnings.append(
                f"В боевом файле нет столбцов ({len(only_source)}), которые есть "
                f"в копии: {_listing(source, only_source)}. Переносить их некуда."
            )
        only_target = [code for code in target.codes if code not in source.columns]
        if only_target:
            warnings.append(
                f"В копии нет столбцов ({len(only_target)}), которые есть в "
                f"боевом файле: {_listing(target, only_target)}. "
                "Они не сравнивались."
            )
        return warnings

    # --- шаг 2: запись ------------------------------------------------------

    def apply(self, rows: list[DiffRow] | None = None) -> ApplyResult:
        """Записать отмеченные различия, сохранить, закрыть, записать журнал."""
        config = self.config
        all_rows = self.rows if rows is None else rows
        selected = [r for r in all_rows if r.selected and r.status in APPLICABLE]
        if not selected:
            raise ApplyError("Не отмечено ни одного изменения — записывать нечего.")

        result = ApplyResult()
        journal_rows: list[tuple[DiffRow, str]] = []

        self.progress("Делаю резервную копию боевого файла…")
        result.backup_path = self._backup()
        for note in self._fix_service_names():
            self.progress(note)

        self.progress("Открываю боевой файл в Excel…")
        with self._excel() as excel:
            sheet = self._open(excel, config.target_path, read_only=False)

            self.progress("Сверяю отмеченное с тем, что в боевом файле сейчас…")
            current = read_sheet(sheet, config, TARGET_TITLE)
            if current.duplicates:
                raise WorkbookStructureError(
                    "В боевом файле повторяются ключи X/Y/Z",
                    _duplicate_lines(current),
                )
            writable: list[DiffRow] = []
            guarded = config.protected_codes()
            for row in selected:
                problem = self._recheck(row, current, guarded)
                if problem:
                    result.stale += 1
                    text = f"{row.address or row.key} ({row.key}): {problem}"
                    result.verify_warnings.append(text + " Не записано.")
                    journal_rows.append((row, f"не применено: {problem}"))
                else:
                    writable.append(row)

            if writable:
                apply_formats(writable, read_formats(sheet, format_targets(writable)))
                self._write(excel, sheet, writable, result, journal_rows)

                self.progress("Сохраняю книгу (слияние с правками коллег)…")
                excel.books[-1].save()

                if config.verify_after_save:
                    self.progress("Проверяю, что записанное легло в файл…")
                    mismatches = self._verify(sheet, writable)
                    result.verify_warnings.extend(text for _row, text in mismatches)
                    journal_rows.extend(
                        (row, f"внимание: {text}") for row, text in mismatches
                    )

        # По id: DiffRow изменяемый и сравнивается по всем полям, так что
        # «row not in selected» на тысячах строк — это миллионы сравнений.
        chosen = {id(row) for row in selected}
        skipped = [r for r in all_rows if r.applicable and id(r) not in chosen]
        result.skipped = len(skipped)
        journal_rows.extend((row, "не отмечено") for row in skipped)
        journal_rows.extend(
            (row, f"не применено: {row.status.value}")
            for row in all_rows
            if not row.applicable
        )

        # Журнал пишется уже после save(): изменения в файле, и потерять их
        # из-за недоступной папки нельзя — о сбое просто сообщим.
        try:
            journal = Journal(config.log_dir)
            result.journal_path = journal.write(config.source_path.name, journal_rows)
        except OSError as error:
            result.verify_warnings.append(
                "Изменения применены и сохранены, но журнал записать не удалось: "
                + explain_path_error(config.log_dir, error)
            )

        self.progress("Готово. " + result.summary())
        return result

    @staticmethod
    def _recheck(row: DiffRow, current: SheetData, guarded: set[str]) -> str:
        """Та ли ячейка в боевом файле, которую человек видел в таблице.

        Между сравнением и записью боевой файл никто не держал открытым, и
        коллеги могли его менять. Строка ищется заново по ключу, столбец — по
        коду (вдруг вставили строку или столбец), значение сверяется с тем,
        что было показано. Пустая строка — всё как было, можно писать.
        """
        target_row = current.rows.get(row.key)
        if target_row is None:
            return "строки с этим кодом X/Y/Z больше нет в боевом файле."
        column = current.columns.get(row.code)
        if column is None:
            return f"столбца с кодом {row.code} больше нет в боевом файле."

        formula = target_row.formula(column)
        now = display(target_row.value(column))
        if formula and (row.code in guarded or not row.replaces_formula):
            return f"в ячейке появилась формула {formula}."
        if now != row.current_display:
            if now == row.new_display:
                return "в боевом файле уже стоит это значение."
            return (
                "после сравнения значение в боевом файле изменилось: было "
                f"«{row.current_display}», сейчас «{now}»."
            )
        # Адрес берём свежий: строки и столбцы могли сдвинуться.
        row.row = target_row.row
        row.column = get_column_letter(column)
        row.replaces_formula = bool(formula)
        return ""

    def _backup(self) -> Path:
        source = self.config.target_path
        if not source.exists():
            raise ApplyError(
                f"Боевой файл не найден: {source}",
                ["Проверьте, подключён ли сетевой диск и открыт ли к нему доступ."],
            )
        target = _unique_path(self.config.backup_dir / f"{_stamp()}_{source.name}")
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)
        except OSError as error:
            # Сюда прилетают отказы и сетевого диска (читаем оттуда), и папки
            # бэкапов (пишем туда).
            raise ApplyError(
                "Не удалось снять резервную копию боевого файла.",
                [
                    explain_path_error(source, error),
                    explain_path_error(target.parent, error),
                ],
            ) from error
        self._cleanup_backups()
        return target

    def _cleanup_backups(self) -> None:
        keep = self.config.backup_keep
        if keep <= 0:
            return
        backups = sorted(
            self.config.backup_dir.glob(f"*{self.config.target_path.suffix}"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        for old in backups[keep:]:
            with suppress(OSError):
                old.unlink()

    def _fix_service_names(self) -> list[str]:
        """Убрать из боевого файла имя, из-за которого Excel спросит «Конфликт имен».

        Правится ровно ``xl/workbook.xml`` внутри книги — мимо Excel, после
        того как снят бэкап. Лечит один запуск: при сохранении Excel запишет
        имя обратно в кривом виде, поэтому проверяем каждый раз заново.
        """
        if not self.config.fix_service_names:
            return []
        path = self.config.target_path
        try:
            problems = name_check.scan(path)
        except name_check.NameCheckError:
            return []  # книга не читается как zip — не наш случай, скажет Excel
        if not problems:
            return []

        names = ", ".join(f"«{item.name}»" for item in problems)
        if name_check.is_open_elsewhere(path):
            # Правка под чужой открытой сессией общего доступа сломает слияние.
            return [
                (
                    f"В боевом файле служебное имя {names} без префикса _xlnm. — "
                    "из-за него Excel спрашивает «Конфликт имен». Починить сейчас "
                    "нельзя: книга у кого-то открыта. Если окно появится, отвечу сам."
                )
            ]
        try:
            fixed = name_check.repair(path)
        except (OSError, name_check.NameCheckError) as error:
            return [
                (
                    f"Не удалось починить служебное имя {names}: {error}. "
                    "Если Excel спросит «Конфликт имен», отвечу сам."
                )
            ]
        return [
            "Поправил служебные имена до открытия: "
            + ", ".join(f"{item.name} -> {item.fixed}" for item in fixed)
        ]

    @staticmethod
    @contextmanager
    def _calculation_paused(excel: ExcelInstance) -> Iterator[None]:
        """Выключить автопересчёт на время записи, в конце пересчитать один раз.

        В книге больше 40 000 формул, и при автопересчёте Excel перестраивает
        зависимости после КАЖДОЙ записанной ячейки: замерено 119 мс на ячейку
        против 8 мс с ручным режимом (tools/write_bench.py на боевом файле).
        Сотня правок — это 12 секунд против одной.

        Пересчёт обязателен до ``save()``: иначе в файл лягут старые
        закешированные значения формул, и коллеги увидят их до первого
        открытия книги в Excel. Режим восстанавливается в любом случае.
        """
        api = getattr(excel.app, "api", None)
        previous: Any = None
        if api is not None:
            with suppress(Exception):
                previous = api.Calculation
                api.Calculation = XL_CALCULATION_MANUAL
        try:
            yield
        finally:
            if api is not None:
                # Считаем только «грязные» ячейки — те, что зависят от записанных.
                with suppress(Exception):
                    api.Calculate()
                with suppress(Exception):
                    api.Calculation = previous or XL_CALCULATION_AUTOMATIC

    def _write(
        self,
        excel: ExcelInstance,
        sheet: Any,
        rows: list[DiffRow],
        result: ApplyResult,
        journal_rows: list[tuple[DiffRow, str]],
    ) -> None:
        total = len(rows)
        with self._calculation_paused(excel):
            for number, row in enumerate(rows, start=1):
                if number == 1 or number == total or number % WRITE_REPORT_EVERY == 0:
                    self.progress(f"Пишу {number} из {total}: {row.address}…")
                try:
                    # Пишем через xlwings, а не напрямую в COM: pywin32 сдвигает
                    # дату без часового пояса на разницу с UTC (проверено:
                    # 30.06.2026 ложится как 29.06.2026 21:00).
                    sheet.range(row.address).value = self._value_for(row)
                    result.applied += 1
                    # Замена формулы значением — не рядовая правка: в журнале
                    # это должно быть видно без чтения таблицы.
                    journal_rows.append(
                        (
                            row,
                            "применено, заменена формула"
                            if row.replaces_formula
                            else "применено",
                        )
                    )
                except Exception as exc:  # noqa: BLE001 — одна ячейка не валит всё
                    result.failed += 1
                    result.errors.append(f"{row.address} ({row.key}): {exc}")
                    journal_rows.append((row, f"ошибка: {exc}"))
            self.progress("Пересчитываю формулы…")

    @staticmethod
    def _value_for(row: DiffRow) -> Any:
        """Что именно писать в ячейку."""
        value = row.new_value
        if value is None or (isinstance(value, str) and not value.strip()):
            return None  # очистка
        if isinstance(value, str):
            value = coerce_for_write(value)
        if row.text_format and not isinstance(value, str):
            # В ячейку с форматом «Текстовый» Excel всё равно положит строку,
            # но по локали США («7/15/2026»). Пишем то, что видно в копии.
            return row.new_display
        return value

    @staticmethod
    def _verify(sheet: Any, rows: list[DiffRow]) -> list[tuple[DiffRow, str]]:
        """Перечитать записанные ячейки после save().

        Если значение отличается — значит правку перебил кто-то другой либо
        сработал ConflictResolution не в нашу пользу. Молчать об этом нельзя.
        """
        cells = {
            id(row): (row.row, column_index_from_string(row.column))
            for row in rows
            if row.row
        }
        values = read_cells(sheet, list(cells.values()))
        warnings: list[tuple[DiffRow, str]] = []
        for row in rows:
            cell = cells.get(id(row))
            if cell is None:
                continue
            actual = display(values.get(cell))
            if actual != row.new_display:
                warnings.append(
                    (
                        row,
                        (
                            f"{row.address} ({row.key}): ожидалось "
                            f"«{row.new_display}», в файле после сохранения «{actual}»"
                        ),
                    )
                )
        return warnings


__all__ = [
    "ApplyError",
    "ApplyResult",
    "CompareResult",
    "SyncSession",
    "WorkbookStructureError",
]
