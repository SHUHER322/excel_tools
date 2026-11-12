"""Чтение таблиц: Excel (.xlsx, .xlsm) и CSV.

В реальных файлах шапка часто не в первой строке (сверху заголовок «Отчёт за октябрь»),
CSV бывает в UTF-8 или Windows-1251, через «;» или «,», с десятичной запятой.
Всё это распознаётся само — на выходе обычная таблица pandas.
"""
import csv
import glob
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .values import infer_types, is_text, normalize_text

EXCEL = {".xlsx", ".xlsm"}
CSV = {".csv", ".txt"}
HEADER_SEARCH_ROWS = 20


class TableError(Exception):
    """Понятная ошибка для пользователя: что за файл и что с ним не так."""


@dataclass
class Table:
    """Лист Excel или CSV-файл, прочитанный в таблицу."""
    file: Path
    sheet: str | None
    frame: pd.DataFrame

    @property
    def label(self) -> str:
        return f"{self.file.name} [{self.sheet}]" if self.sheet else self.file.name


def expand_inputs(items) -> list[Path]:
    """Файлы, папки и шаблоны (*.xlsx) → список файлов таблиц без повторов.
    Шаблоны раскрываются здесь же: командная строка Windows сама этого не делает."""
    files: list[Path] = []
    for item in map(str, items):
        path = Path(item)
        if path.is_dir():
            found = sorted((p for p in path.iterdir() if _is_table(p)), key=_order)
        elif any(ch in item for ch in "*?["):
            found = sorted((Path(p) for p in glob.glob(item) if _is_table(Path(p))), key=_order)
        elif path.is_file():
            found = [path]
        else:
            raise TableError(f"Не найден файл или папка: {item}")
        if not found:
            raise TableError(f"Нет таблиц (.xlsx, .xlsm, .csv) в: {item}")
        files += [p for p in found if p not in files]
    return files


def _order(path: Path) -> str:
    """По алфавиту без учёта регистра — одинаково на Windows и Linux."""
    return path.name.casefold()


def _is_table(path: Path) -> bool:
    # ~$файл.xlsx — служебный файл, который Excel создаёт рядом с открытой книгой
    return path.is_file() and path.suffix.lower() in EXCEL | CSV and not path.name.startswith("~$")


def read_tables(path: Path, sheet: str | None = None) -> list[Table]:
    """Все листы книги Excel (или один, если указан sheet) либо CSV-файл."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in CSV:
        return [Table(path, None, frame_from_rows(read_csv_raw(path), infer=True))]
    if suffix not in EXCEL:
        raise TableError(f"{path.name}: поддерживаются .xlsx, .xlsm и .csv")
    try:
        raw = pd.read_excel(path, sheet_name=sheet if sheet else None, header=None, dtype=object)
    except PermissionError:
        raise TableError(f"{path.name}: файл занят — закройте его в Excel") from None
    except ValueError as e:                     # нет такого листа
        raise TableError(f"{path.name}: {e}") from None
    except Exception as e:                      # битый или зашифрованный файл
        raise TableError(f"{path.name}: не удалось открыть — {e}") from None
    sheets = raw if isinstance(raw, dict) else {sheet: raw}
    tables = [Table(path, name, frame_from_rows(rows)) for name, rows in sheets.items()]
    return [t for t in tables if not t.frame.empty]


def read_csv_raw(path: Path) -> pd.DataFrame:
    """CSV как есть, все значения текстом: кодировка и разделитель подбираются."""
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise TableError(f"{path.name}: неизвестная кодировка (ожидается UTF-8 или Windows-1251)")
    delimiter = sniff_delimiter(text)
    rows = list(csv.reader(text.splitlines(), delimiter=delimiter))
    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]
    return pd.DataFrame(rows, dtype=object)


def sniff_delimiter(text: str) -> str:
    """«;» в русском Excel, «,» в остальном мире, иногда табуляция — какой встречается
    одинаковое число раз в большинстве строк."""
    lines = [l for l in text.splitlines()[:30] if l.strip()]
    best, best_score = ",", -1
    for delimiter in (";", ",", "\t"):
        counts = [len(next(csv.reader([l], delimiter=delimiter))) - 1 for l in lines]
        if not counts or max(counts) == 0:
            continue
        common = max(set(counts), key=counts.count)
        score = counts.count(common) * 10 + common
        if common > 0 and score > best_score:
            best, best_score = delimiter, score
    return best


def frame_from_rows(raw: pd.DataFrame, infer: bool = False) -> pd.DataFrame:
    """Сырые строки листа → таблица: найти строку шапки, убрать пустые строки и столбцы.
    infer=True — превратить числа и даты, записанные текстом, в настоящие (для CSV)."""
    raw = raw.map(lambda v: None if is_text(v) and not normalize_text(v) else v)
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return pd.DataFrame()
    pos = header_position(raw)
    frame = raw.iloc[pos + 1:].copy()
    frame.columns = unique_names(raw.iloc[pos].tolist())
    frame = frame.dropna(how="all").reset_index(drop=True)
    return infer_types(frame) if infer else frame


def header_position(raw: pd.DataFrame) -> int:
    """Номер строки шапки: первая строка, где заполнена большая часть столбцов.
    Строки выше (название отчёта, дата выгрузки) обычно заполнены в одной-двух ячейках."""
    filled = raw.head(HEADER_SEARCH_ROWS).notna().sum(axis=1).tolist()
    widest = max(filled)
    need = max(2, round(widest * 0.6)) if widest > 1 else 1
    return next(i for i, n in enumerate(filled) if n >= need)


def unique_names(values) -> list[str]:
    """Названия столбцов: без лишних пробелов, пустые → «Столбец N», повторы → «Имя (2)»."""
    names: list[str] = []
    for i, value in enumerate(values, start=1):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            name = ""
        elif isinstance(value, float) and value.is_integer():
            name = str(int(value))              # год 2026.0 → «2026»
        else:
            name = normalize_text(str(value))
        name = name or f"Столбец {i}"
        base, n = name, 2
        while name in names:
            name = f"{base} ({n})"
            n += 1
        names.append(name)
    return names
