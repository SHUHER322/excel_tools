"""Сводный отчёт: итоги по группам — как сводная таблица Excel, но одной командой.

  --by Менеджер                    по менеджерам
  --by Дата:месяц --by Категория   по месяцам и категориям
  --pivot Категория                категории — в столбцы (перекрёстная таблица)
"""
import pandas as pd

from .columns import find_column, find_columns, numeric_columns
from .tables import TableError
from .values import column_key, is_text, parse_date

TOTAL = "Итого"
ROWS = "Строк"
AGGREGATIONS = {"sum": "сумма", "mean": "среднее", "min": "минимум", "max": "максимум"}
PERIODS = {"день": "D", "неделя": "W-SUN", "месяц": "M", "квартал": "Q", "год": "Y"}


def parse_group(spec: str) -> tuple[str, str | None]:
    """«Дата:месяц» → («Дата», «месяц»); «Менеджер» → («Менеджер», None)."""
    name, _, period = spec.partition(":")
    period = period.strip().lower() or None
    if period and period not in PERIODS:
        raise TableError(f"Неизвестный период «{period}». Можно: {', '.join(PERIODS)}")
    return name.strip(), period


def period_label(value: pd.Period, period: str) -> str:
    if period == "день":
        return value.start_time.strftime("%d.%m.%Y")
    if period == "неделя":
        return f"{value.start_time:%d.%m}–{value.end_time:%d.%m.%Y}"
    if period == "месяц":
        return value.start_time.strftime("%m.%Y")
    if period == "квартал":
        return f"{value.quarter} кв. {value.year}"
    return str(value.year)


def as_dates(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    return pd.to_datetime(series.map(lambda v: parse_date(v) if is_text(v) else v), errors="coerce")


def sort_column(columns: list[str]) -> str | None:
    """По какому столбцу упорядочить группы: сумма или выручка, если есть, иначе первый."""
    for column in columns:
        if any(word in column_key(column) for word in ("сумм", "выручк", "итог", "стоимост")):
            return column
    return columns[0] if columns else None


def group_keys(frame: pd.DataFrame, specs: list[str]) -> tuple[pd.DataFrame, list[str], dict[str, str]]:
    """Столбцы группировки. Для «Дата:месяц» добавляется столбец «Дата (месяц)» с периодами."""
    frame = frame.copy()
    keys, periods = [], {}
    for spec in specs:
        name, period = parse_group(spec)
        column = find_column(frame, name)
        if period is None:
            keys.append(column)
            continue
        dates = as_dates(frame[column])
        if dates.notna().sum() == 0:
            raise TableError(f"В столбце «{column}» нет дат — по периодам не сгруппировать")
        key = f"{column} ({period})"
        frame[key] = dates.dt.to_period(PERIODS[period])
        keys.append(key)
        periods[key] = period
    return frame, keys, periods


def summarize(frame: pd.DataFrame, by: list[str], values: list[str] | None = None,
              agg: str = "sum", pivot: str | None = None) -> pd.DataFrame:
    """Итоги по группам с последней строкой «Итого».
    values — какие числа считать (по умолчанию все числовые столбцы)."""
    if not by:
        raise TableError("Укажите, по чему группировать: --by Столбец")
    if agg not in AGGREGATIONS:
        raise TableError(f"Неизвестный способ подсчёта «{agg}». Можно: {', '.join(AGGREGATIONS)}")
    frame, keys, periods = group_keys(frame, by)
    columns = find_columns(frame, values) if values else numeric_columns(frame, exclude=keys)
    if pivot:
        return _pivot(frame, keys, periods, find_column(frame, pivot), columns, agg)

    grouped = frame.groupby(keys, dropna=False, sort=False)
    table = grouped.size().rename(ROWS).to_frame()
    for column in columns:
        table[column] = grouped[column].agg(agg)
    table = table.reset_index()
    table = _sort(table, keys, periods, columns)
    total = {keys[0]: TOTAL, ROWS: len(frame)}
    for column in columns:
        total[column] = getattr(frame[column], agg)()
    return pd.concat([table, pd.DataFrame([total])], ignore_index=True)


def _sort(table: pd.DataFrame, keys, periods, columns) -> pd.DataFrame:
    """По времени, если группировка по периодам; иначе — от большей суммы к меньшей."""
    if periods:
        table = table.sort_values(list(periods), kind="stable")
    elif sort_column(columns):
        table = table.sort_values(sort_column(columns), ascending=False, kind="stable")
    for key, period in periods.items():
        table[key] = [period_label(v, period) if isinstance(v, pd.Period) else v for v in table[key]]
    return table.reset_index(drop=True)


def _pivot(frame, keys, periods, pivot, columns, agg) -> pd.DataFrame:
    """Перекрёстная таблица: строки — группы, столбцы — значения pivot, в ячейках — первый
    из числовых столбцов."""
    if not columns:
        raise TableError("Для --pivot нужен числовой столбец: --values Сумма")
    value = columns[0]
    table = pd.pivot_table(frame, index=keys, columns=pivot, values=value, aggfunc=agg,
                           dropna=False, observed=True)
    table.columns = [str(c) for c in table.columns]
    table[TOTAL] = frame.groupby(keys, dropna=False)[value].agg(agg)
    table = table.reset_index()
    table = _sort(table, keys, periods, [TOTAL])
    total = {keys[0]: TOTAL}
    for column in table.columns[len(keys):]:
        part = frame[frame[pivot].astype(str) == column][value] if column != TOTAL else frame[value]
        total[column] = getattr(part, agg)()
    return pd.concat([table, pd.DataFrame([total])], ignore_index=True)
