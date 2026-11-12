"""Объединение многих таблиц в одну.

Столбцы сопоставляются по смыслу названия, а не по позиции: «Сумма», « сумма »
и «СУММА» — один столбец, даже если в файлах он стоит на разных местах.
"""
from dataclasses import dataclass, field

import pandas as pd

from .tables import Table
from .values import column_key

SOURCE_FILE = "Файл"
SOURCE_SHEET = "Лист"


@dataclass
class MergeReport:
    tables: list[tuple[str, int]] = field(default_factory=list)     # источник, строк
    missing: dict[str, list[str]] = field(default_factory=dict)     # источник → каких столбцов нет
    rows: int = 0

    def lines(self) -> list[str]:
        out = [f"  {label}: {n} строк" for label, n in self.tables]
        out.append(f"Итого строк: {self.rows}")
        for label, columns in self.missing.items():
            out.append(f"  ⚠ в {label} нет столбцов: {', '.join(columns)}")
        return out


def merge(tables: list[Table], with_source: bool = True) -> tuple[pd.DataFrame, MergeReport]:
    report = MergeReport()
    seen: dict[str, list[str]] = {}         # ключ → как этот столбец назван в разных файлах
    parts = []
    for table in tables:
        frame = table.frame.copy()
        keys = []
        for column in frame.columns:
            key = column_key(column) or column
            while key in keys:              # два одинаковых по смыслу столбца в одном файле
                key += "_"
            seen.setdefault(key, []).append(column)
            keys.append(key)
        frame.columns = keys
        if with_source:
            frame.insert(0, SOURCE_FILE, table.file.name)
            if table.sheet is not None:
                frame.insert(1, SOURCE_SHEET, table.sheet)
        parts.append((table, frame))
        report.tables.append((table.label, len(frame)))

    names = {key: display_name(variants) for key, variants in seen.items()}
    all_keys = list(dict.fromkeys(k for _, f in parts for k in f.columns))
    for table, frame in parts:
        absent = [names[k] for k in all_keys if k in names and k not in frame.columns]
        if absent:
            report.missing[table.label] = absent

    frames = [f.astype(object) for _, f in parts if not f.empty]
    merged = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    merged = merged.reindex(columns=all_keys)
    merged.columns = [names.get(k, k) for k in all_keys]
    if with_source and SOURCE_SHEET in merged.columns and merged[SOURCE_SHEET].isna().all():
        merged = merged.drop(columns=SOURCE_SHEET)
    merged = merged.infer_objects()
    report.rows = len(merged)
    return merged, report


def display_name(variants: list[str]) -> str:
    """Название объединённого столбца — то, как он назван в большинстве файлов:
    «Товар» в двух файлах и «ТОВАР» в одном → «Товар»."""
    return max(dict.fromkeys(variants), key=variants.count)
