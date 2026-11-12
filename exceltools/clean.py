"""Очистка таблицы: пробелы, пустые строки, числа и даты текстом, дубли."""
from dataclasses import dataclass, field

import pandas as pd

from .values import convert_types, is_text, normalize_text


@dataclass
class CleanReport:
    rows_before: int = 0
    rows_after: int = 0
    empty_rows: int = 0
    duplicates: int = 0
    trimmed_cells: int = 0
    numbers: list[str] = field(default_factory=list)
    dates: list[str] = field(default_factory=list)

    def lines(self) -> list[str]:
        out = [f"Строк: было {self.rows_before}, стало {self.rows_after}"]
        if self.empty_rows:
            out.append(f"Удалено пустых строк: {self.empty_rows}")
        if self.duplicates:
            out.append(f"Удалено повторов: {self.duplicates}")
        if self.trimmed_cells:
            out.append(f"Убраны лишние пробелы в ячейках: {self.trimmed_cells}")
        if self.numbers:
            out.append("Числа из текста: " + ", ".join(self.numbers))
        if self.dates:
            out.append("Даты из текста: " + ", ".join(self.dates))
        return out


def clean(frame: pd.DataFrame, drop_duplicates: bool = True) -> tuple[pd.DataFrame, CleanReport]:
    report = CleanReport(rows_before=len(frame))
    trimmed = 0

    def tidy(value):
        nonlocal trimmed
        if not is_text(value):
            return value
        text = normalize_text(value)
        if text != value:
            trimmed += 1
        return text or None

    frame = frame.astype(object).map(tidy)
    report.trimmed_cells = trimmed

    before = len(frame)
    frame = frame.dropna(how="all").dropna(axis=1, how="all")
    report.empty_rows = before - len(frame)

    frame, changed = convert_types(frame)
    report.numbers, report.dates = changed["numbers"], changed["dates"]

    if drop_duplicates:
        before = len(frame)
        frame = frame.drop_duplicates()
        report.duplicates = before - len(frame)

    frame = frame.reset_index(drop=True)
    report.rows_after = len(frame)
    return frame, report
