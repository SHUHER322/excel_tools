"""Значения ячеек: текст, числа и даты, записанные «как попало».

«12 500,50 ₽», «1 234», «-3,5%» → числа; «05.10.2026», «2026-10-05» → даты.
Столбец переводится в числа или даты, только если так читается почти всё в нём —
иначе артикулы вроде «00123» превратились бы в 123.
"""
import re
from datetime import datetime

import pandas as pd

SHARE = 0.9                 # какая доля непустых значений должна читаться как число/дата

_SPACES = re.compile(r"\s+")
_NUMBER = re.compile(r"^[-+]?\d{1,3}(?:[   ]\d{3})+(?:[.,]\d+)?$|^[-+]?\d+(?:[.,]\d+)?$")
_CURRENCY = re.compile(r"\s*(?:₽|руб\.?|р\.|rub|\$|€|%)$", re.I)
_DATE_FORMATS = ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S",
                 "%Y-%m-%d %H:%M:%S")


def is_text(value) -> bool:
    return isinstance(value, str)


def normalize_text(value: str) -> str:
    """Пробелы по краям и двойные пробелы внутри убираются, неразрывные — тоже."""
    return _SPACES.sub(" ", value.replace(" ", " ")).strip()


def column_key(name: str) -> str:
    """Ключ для сопоставления столбцов разных файлов: «  Сумма, руб. », «СУММА» и
    «Сумма (₽)» — один столбец. Единицы измерения в конце названия не учитываются."""
    key = normalize_text(str(name)).lower().replace("ё", "е")
    key = re.sub(r"[^\w%]+", " ", key).strip()
    words = key.split()
    while len(words) > 1 and words[-1] in UNITS:
        words.pop()
    return " ".join(words)


UNITS = {"руб", "р", "rub", "rur", "шт", "ед"}


def parse_number(value):
    """Текст → число или None. Целые остаются целыми: «1 200» → 1200, «1 200,50» → 1200.5."""
    if not is_text(value):
        return None
    text = _CURRENCY.sub("", normalize_text(value))
    if not _NUMBER.match(text):
        return None
    text = re.sub(r"[   ]", "", text).replace(",", ".")
    number = float(text)
    return int(number) if "." not in text and abs(number) < 2 ** 53 else number


def parse_date(value):
    if not is_text(value):
        return None
    text = normalize_text(value)
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _converted(series: pd.Series, parse, native: tuple):
    """Столбец, переведённый в числа или даты, или None, если он не такой.
    native — типы, которые уже подходят (число из Excel не надо разбирать)."""
    values = [v for v in series if not _empty(v)]
    texts = [v for v in values if is_text(v)]
    if not texts:
        return None
    parsed = {t: parse(t) for t in texts}
    ok = sum(parsed[t] is not None for t in texts) + sum(isinstance(v, native) for v in values)
    if ok < SHARE * len(values):
        return None
    if parse is parse_number and any(re.match(r"^0\d", t) for t in texts):
        return None                     # «00123» — это код, а не число
    return series.map(lambda v: parsed.get(v) if is_text(v) and parsed.get(v) is not None else v)


def _empty(value) -> bool:
    return value is None or (not isinstance(value, str) and pd.isna(value))


def infer_types(frame: pd.DataFrame) -> pd.DataFrame:
    """Столбцы, где почти все значения — числа или даты текстом, переводятся в числа/даты."""
    return convert_types(frame)[0]


def convert_types(frame: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, list[str]]]:
    """То же, что infer_types, плюс какие столбцы стали числами и датами — для отчёта."""
    frame = frame.copy()
    changed = {"numbers": [], "dates": []}
    for column in frame.columns:
        series = frame[column].astype(object)
        for kind, parse, native in (("dates", parse_date, (datetime, pd.Timestamp)),
                                    ("numbers", parse_number, (int, float))):
            converted = _converted(series, parse, native)
            if converted is not None:
                frame[column] = _typed(converted, kind)
                changed[kind].append(column)
                break
    return frame, changed


def _typed(series: pd.Series, kind: str) -> pd.Series:
    """Если всё в столбце прочиталось — настоящий тип столбца (числа, даты).
    Если остались непрочитанные значения («нет», «уточнить») — они остаются как есть,
    ничего не теряется."""
    filled = series[[not _empty(v) for v in series]]
    if any(is_text(v) for v in filled):
        return series.astype(object)
    if kind == "dates":
        return pd.to_datetime(series)
    numbers = pd.to_numeric(series)
    if numbers.dropna().map(lambda v: float(v).is_integer()).all():
        return numbers.astype("Int64")
    return numbers.astype(float)
