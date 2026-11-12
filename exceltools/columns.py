"""Поиск столбца по названию, как его написал человек: «менеджер» найдёт «Менеджер»,
«сумма» — «Сумма, руб.»."""
import pandas as pd

from .tables import TableError
from .values import column_key


def find_column(frame: pd.DataFrame, name: str) -> str:
    wanted = column_key(name)
    for column in frame.columns:
        if column == name or column_key(column) == wanted:
            return column
    available = ", ".join(f"«{c}»" for c in frame.columns)
    raise TableError(f"Нет столбца «{name}». Есть: {available}")


def find_columns(frame: pd.DataFrame, names) -> list[str]:
    return [find_column(frame, n) for n in names or []]


def numeric_columns(frame: pd.DataFrame, exclude=()) -> list[str]:
    return [c for c in frame.columns
            if c not in exclude and pd.api.types.is_numeric_dtype(frame[c])
            and not pd.api.types.is_bool_dtype(frame[c])]
