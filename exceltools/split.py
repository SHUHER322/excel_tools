"""Разбивка таблицы по значению столбца: каждому менеджеру, филиалу, поставщику — свой файл
(или свой лист в одной книге)."""
import re

import pandas as pd

from .columns import find_column

EMPTY = "(пусто)"
RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
FORBIDDEN = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def split(frame: pd.DataFrame, by: str) -> tuple[str, dict[str, pd.DataFrame]]:
    """{значение: строки с этим значением} в порядке первого появления."""
    column = find_column(frame, by)
    labels = frame[column].map(lambda v: EMPTY if v is None or (not isinstance(v, str) and pd.isna(v))
                               else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip()))
    parts = {}
    for label in dict.fromkeys(labels):
        parts[label] = frame[labels == label].reset_index(drop=True)
    return column, parts


def safe_filename(name: str, taken: set[str]) -> str:
    """Имя файла, допустимое в Windows: без <>:"/\\|?*, не длиннее 100 символов, без повторов
    (в Windows «Иванов» и «иванов» — один файл)."""
    base = FORBIDDEN.sub("_", name).strip(" .") or "без_названия"
    if base.split(".")[0].upper() in RESERVED:
        base = f"_{base}"                       # CON, PRN, NUL… — служебные имена Windows
    base = base[:100]
    candidate, n = base, 2
    while candidate.casefold() in taken:
        candidate = f"{base} ({n})"
        n += 1
    taken.add(candidate.casefold())
    return candidate
