"""Запись результата в оформленный Excel.

Шапка выделена и закреплена, включены фильтры, ширина столбцов по содержимому,
числа и даты — настоящие числа и даты Excel (по ним работают сортировка и формулы),
а не текст.
"""
import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
FILLS = {                                   # подсветка ячеек: изменения, новые и удалённые строки
    "changed": PatternFill("solid", fgColor="FFF2CC"),
    "added": PatternFill("solid", fgColor="E2EFDA"),
    "removed": PatternFill("solid", fgColor="FCE4D6"),
}
MAX_WIDTH = 50
INT_FORMAT = "#,##0"
FLOAT_FORMAT = "#,##0.00"
DATE_FORMAT = "DD.MM.YYYY"
DATETIME_FORMAT = "DD.MM.YYYY HH:MM"


@dataclass
class Sheet:
    """Лист результата. marks — подсветка: (номер строки таблицы, столбец или None, вид)."""
    name: str
    frame: pd.DataFrame
    total_row: bool = False                 # последняя строка — «Итого»: выделить жирным
    marks: list[tuple[int, str | None, str]] = field(default_factory=list)


def cell_value(value):
    """Значение pandas → значение для ячейки Excel."""
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if hasattr(value, "item") and not isinstance(value, (str, bytes)):
        return value.item()                 # numpy-числа → обычные числа Python
    return value


def number_format(value) -> str | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return DATE_FORMAT if (value.hour, value.minute, value.second) == (0, 0, 0) else DATETIME_FORMAT
    if isinstance(value, date):
        return DATE_FORMAT
    if isinstance(value, int):
        return INT_FORMAT
    if isinstance(value, float):
        return INT_FORMAT if value.is_integer() and abs(value) >= 1000 else FLOAT_FORMAT
    return None


def safe_sheet_name(name: str, taken: set[str]) -> str:
    """Excel: до 31 символа, без []:*?/\\ и без повторов."""
    base = re.sub(r"[\[\]:*?/\\]", "_", str(name)).strip() or "Лист"
    base = base[:31]
    candidate, n = base, 2
    while candidate.lower() in {t.lower() for t in taken}:
        suffix = f" ({n})"
        candidate = base[:31 - len(suffix)] + suffix
        n += 1
    taken.add(candidate)
    return candidate


def text_width(value) -> int:
    if value is None:
        return 0
    if isinstance(value, datetime):
        return 16 if (value.hour, value.minute) != (0, 0) else 10
    if isinstance(value, float):
        return len(f"{value:,.2f}")
    if isinstance(value, int):
        return len(f"{value:,}")
    return max(len(line) for line in str(value).splitlines() or [""])


def build_workbook(sheets: list[Sheet]) -> Workbook:
    book = Workbook()
    book.remove(book.active)
    taken: set[str] = set()
    for sheet in sheets:
        ws = book.create_sheet(safe_sheet_name(sheet.name, taken))
        frame = sheet.frame
        columns = [str(c) for c in frame.columns]
        ws.append(columns)
        widths = [len(c) + 3 for c in columns]          # + место под кнопку фильтра
        for row in frame.itertuples(index=False, name=None):
            values = [cell_value(v) for v in row]
            ws.append(values)
            r = ws.max_row
            for c, value in enumerate(values, start=1):
                fmt = number_format(value)
                if fmt:
                    ws.cell(r, c).number_format = fmt
                widths[c - 1] = max(widths[c - 1], text_width(value))

        for c in range(1, len(columns) + 1):
            cell = ws.cell(1, c)
            cell.fill, cell.font = HEADER_FILL, HEADER_FONT
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c)].width = min(MAX_WIDTH, widths[c - 1] + 3)
        if columns:
            ws.freeze_panes = "A2"
            last = ws.max_row - 1 if sheet.total_row else ws.max_row
            ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(last, 1)}"
        if sheet.total_row and len(frame):
            for c in range(1, len(columns) + 1):
                cell = ws.cell(ws.max_row, c)
                cell.font, cell.fill = TOTAL_FONT, TOTAL_FILL
        for row, column, kind in sheet.marks:
            cols = [columns.index(str(column)) + 1] if column is not None else range(1, len(columns) + 1)
            for c in cols:
                ws.cell(row + 2, c).fill = FILLS[kind]
    if not book.sheetnames:
        book.create_sheet("Пусто")
    return book


def save_workbook(sheets: list[Sheet], path: Path) -> Path:
    """Сохранить книгу. Если файл с таким именем открыт в Excel (Windows не даёт
    перезаписать), результат сохраняется рядом с отметкой времени — работа не теряется."""
    path = Path(path).with_suffix(".xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)
    book = build_workbook(sheets)
    try:
        book.save(path)
        return path
    except PermissionError:
        other = path.with_name(f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        book.save(other)
        return other
