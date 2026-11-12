"""Сравнение двух версий таблицы: что добавилось, что пропало, что изменилось.

Строки сопоставляются по ключу — столбцу, который не меняется (артикул, код, ИНН),
а не по номеру строки: если в новом прайсе товары переставлены, это не «изменение».
"""
import math
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd

from .columns import find_columns
from .excel import Sheet
from .tables import TableError
from .values import column_key, is_text, normalize_text

CHANGE_COLUMNS = ["Столбец", "Было", "Стало", "Разница"]
STATUS = "Статус"


@dataclass
class Comparison:
    key: list[str]
    changes: pd.DataFrame                     # по строке на изменённую ячейку
    added: pd.DataFrame
    removed: pd.DataFrame
    full: pd.DataFrame                        # новая таблица + удалённые строки в конце, со статусом
    marks: list[tuple[int, str | None, str]] = field(default_factory=list)
    only_old: list[str] = field(default_factory=list)       # столбцы, которых нет в новой таблице
    only_new: list[str] = field(default_factory=list)
    duplicate_keys: list[str] = field(default_factory=list)

    @property
    def changed_rows(self) -> int:
        return self.changes[self.key].drop_duplicates().shape[0] if len(self.changes) else 0

    def lines(self) -> list[str]:
        out = [f"Изменено строк: {self.changed_rows} (ячеек: {len(self.changes)})",
               f"Новых строк: {len(self.added)}",
               f"Удалённых строк: {len(self.removed)}"]
        if self.only_new:
            out.append("  Новые столбцы: " + ", ".join(self.only_new))
        if self.only_old:
            out.append("  Пропавшие столбцы: " + ", ".join(self.only_old))
        if self.duplicate_keys:
            sample = ", ".join(self.duplicate_keys[:5])
            out.append(f"  ⚠ Ключ повторяется ({len(self.duplicate_keys)}): {sample} — "
                       "сравнивается первая из одинаковых строк")
        return out

    def sheets(self) -> list[Sheet]:
        changes = self.changes if len(self.changes) else pd.DataFrame(columns=self.key + CHANGE_COLUMNS)
        return [
            Sheet("Сравнение", self.full, marks=self.marks),
            Sheet("Изменения", changes),
            Sheet("Новые", self.added),
            Sheet("Удалённые", self.removed),
        ]


def same(a, b) -> bool:
    if _empty(a) and _empty(b):
        return True
    if _empty(a) or _empty(b):
        return False
    if _number(a) and _number(b):
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-9)
    if isinstance(a, (datetime, pd.Timestamp)) and isinstance(b, (datetime, pd.Timestamp)):
        return pd.Timestamp(a) == pd.Timestamp(b)
    if is_text(a) and is_text(b):
        return normalize_text(a) == normalize_text(b)
    return str(a) == str(b)


def _empty(v) -> bool:
    return v is None or (not isinstance(v, str) and pd.isna(v)) or (is_text(v) and not v.strip())


def _number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _key_text(values) -> str:
    """Ключ строки текстом: 123, 123.0 и «123» — один и тот же артикул."""
    parts = []
    for v in values:
        if _empty(v):
            parts.append("")
        elif isinstance(v, float) and v.is_integer():
            parts.append(str(int(v)))
        else:
            parts.append(normalize_text(str(v)))
    return " / ".join(parts)


def _align(old: pd.DataFrame, new: pd.DataFrame) -> tuple[pd.DataFrame, list[str], list[str]]:
    """Названия столбцов старой таблицы — как в новой («ЦЕНА» = «Цена, руб.»)."""
    new_keys = {column_key(c): c for c in new.columns}
    renamed, only_old = {}, []
    for column in old.columns:
        match = new_keys.get(column_key(column))
        if match is None:
            only_old.append(column)
        else:
            renamed[column] = match
    old = old.rename(columns=renamed)
    only_new = [c for c in new.columns if c not in old.columns]
    return old, only_old, only_new


def compare(old: pd.DataFrame, new: pd.DataFrame, key: list[str]) -> Comparison:
    if not key:
        raise TableError("Укажите ключ — столбец, по которому сопоставлять строки: --key Артикул")
    key = find_columns(new, key)
    old, only_old, only_new = _align(old, new)
    missing = [k for k in key if k not in old.columns]
    if missing:
        raise TableError(f"В старой таблице нет ключевого столбца: {', '.join(missing)}")
    common = [c for c in new.columns if c in old.columns and c not in key]
    # первый текстовый столбец (обычно название) — чтобы в «Изменениях» было понятно, о чём строка
    context = next((c for c in common if new[c].map(is_text).mean() > 0.5), None)
    shown = key + ([context] if context else [])

    old_index, duplicates = _index(old, key)
    new_index, new_duplicates = _index(new, key)
    duplicates = list(dict.fromkeys(duplicates + new_duplicates))

    changes, marks, statuses = [], [], []
    for position, (k, row) in enumerate(new_index.items()):
        before = old_index.get(k)
        if before is None:
            statuses.append("новая")
            marks.append((position, None, "added"))
            continue
        row_changed = False
        for column in common:
            a, b = before[column], row[column]
            if not same(a, b):
                row_changed = True
                marks.append((position, column, "changed"))
                diff = b - a if _number(a) and _number(b) else None
                changes.append({**{c: row[c] for c in shown}, "Столбец": column,
                                "Было": a, "Стало": b, "Разница": diff})
        statuses.append("изменена" if row_changed else "")

    removed_keys = [k for k in old_index if k not in new_index]
    full = pd.DataFrame(list(new_index.values()), columns=new.columns)
    full.insert(0, STATUS, statuses)
    removed = pd.DataFrame([old_index[k] for k in removed_keys], columns=old.columns)
    if len(removed):
        tail = removed.reindex(columns=new.columns).copy()
        tail.insert(0, STATUS, "удалена")
        start = len(full)
        full = pd.concat([full.astype(object), tail.astype(object)], ignore_index=True)
        marks += [(start + i, None, "removed") for i in range(len(tail))]

    added = full[full[STATUS] == "новая"].drop(columns=STATUS).reset_index(drop=True)
    changes_frame = pd.DataFrame(changes, columns=shown + CHANGE_COLUMNS)
    return Comparison(key=key, changes=changes_frame, added=added, removed=removed.reset_index(drop=True),
                      full=full.infer_objects(), marks=marks, only_old=only_old, only_new=only_new,
                      duplicate_keys=duplicates)


def _index(frame: pd.DataFrame, key: list[str]) -> tuple[dict[str, dict], list[str]]:
    """Строки по ключу (в порядке таблицы). Строки с пустым ключом пропускаются."""
    rows, duplicates = {}, []
    for record in frame.to_dict("records"):
        values = [record[k] for k in key]
        if all(_empty(v) for v in values):
            continue
        k = _key_text(values)
        if k in rows:
            duplicates.append(k)
            continue
        rows[k] = record
    return rows, duplicates
