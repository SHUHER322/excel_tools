"""Запуск:
  python -m exceltools merge ФАЙЛЫ_ИЛИ_ПАПКИ   — объединить таблицы в одну
  python -m exceltools clean ФАЙЛ              — почистить таблицу
  python -m exceltools summary ФАЙЛЫ --by Менеджер   — сводный отчёт с итогами
  python -m exceltools compare СТАРЫЙ НОВЫЙ --key Артикул — что изменилось
  python -m exceltools split ФАЙЛ --by Менеджер — свой файл каждому менеджеру

У каждой команды есть справка: python -m exceltools merge --help
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

from .clean import clean
from .compare import compare
from .excel import Sheet, save_workbook
from .merge import merge
from .split import safe_filename, split
from .summary import AGGREGATIONS, PERIODS, summarize
from .tables import TableError, expand_inputs, read_tables

OUTPUT = Path("output")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="python -m exceltools",
                                     description="Обработка таблиц Excel и CSV")
    commands = parser.add_subparsers(dest="command", required=True, metavar="КОМАНДА")

    p = commands.add_parser("merge", help="объединить много файлов и листов в одну таблицу",
                            description="Объединить таблицы. Столбцы сопоставляются по названию, "
                                        "даже если в файлах они стоят в разном порядке.")
    p.add_argument("inputs", nargs="+", metavar="ФАЙЛ", help="файлы, папки или шаблоны (data/*.xlsx)")
    p.add_argument("--sheet", help="брать только этот лист (по умолчанию — все листы)")
    p.add_argument("--clean", action="store_true", help="заодно почистить: пробелы, дубли, числа текстом")
    p.add_argument("--keep-duplicates", action="store_true",
                   help="при --clean не удалять повторяющиеся строки")
    p.add_argument("--no-source", action="store_true", help="не добавлять столбцы «Файл» и «Лист»")
    p.add_argument("-o", "--out", type=Path, default=OUTPUT / "объединено.xlsx", help="куда сохранить")

    p = commands.add_parser("clean", help="почистить таблицу: пробелы, пустые строки, дубли, числа текстом")
    p.add_argument("input", metavar="ФАЙЛ")
    p.add_argument("--sheet", help="лист (по умолчанию — все листы, каждый чистится отдельно)")
    p.add_argument("--keep-duplicates", action="store_true", help="не удалять повторяющиеся строки")
    p.add_argument("-o", "--out", type=Path, help="куда сохранить (по умолчанию output/<имя>_чисто.xlsx)")

    p = commands.add_parser("summary", help="сводный отчёт: итоги по группам",
                            description="Итоги по группам — как сводная таблица Excel. "
                                        "Несколько файлов сначала объединяются.")
    p.add_argument("inputs", nargs="+", metavar="ФАЙЛ", help="файлы, папки или шаблоны")
    p.add_argument("--by", action="append", required=True, metavar="СТОЛБЕЦ",
                   help="по чему группировать; можно несколько раз. Для дат: "
                        f"«Дата:месяц» ({', '.join(PERIODS)})")
    p.add_argument("--values", action="append", metavar="СТОЛБЕЦ",
                   help="какие числа считать (по умолчанию все числовые столбцы)")
    p.add_argument("--agg", default="sum", choices=list(AGGREGATIONS),
                   help="сумма (sum), среднее (mean), минимум (min), максимум (max)")
    p.add_argument("--pivot", metavar="СТОЛБЕЦ", help="значения этого столбца — в столбцы отчёта")
    p.add_argument("--sheet", help="брать только этот лист")
    p.add_argument("-o", "--out", type=Path, default=OUTPUT / "сводка.xlsx", help="куда сохранить")

    p = commands.add_parser("compare", help="сравнить две версии таблицы",
                            description="Что добавилось, что пропало и что изменилось. "
                                        "Строки сопоставляются по ключу, а не по номеру.")
    p.add_argument("old", metavar="СТАРЫЙ")
    p.add_argument("new", metavar="НОВЫЙ")
    p.add_argument("--key", action="append", required=True, metavar="СТОЛБЕЦ",
                   help="столбец, по которому узнавать строку (артикул, код); можно несколько")
    p.add_argument("--sheet", help="лист (по умолчанию первый лист с данными)")
    p.add_argument("-o", "--out", type=Path, default=OUTPUT / "сравнение.xlsx", help="куда сохранить")

    p = commands.add_parser("split", help="разбить таблицу на файлы по значению столбца")
    p.add_argument("inputs", nargs="+", metavar="ФАЙЛ")
    p.add_argument("--by", required=True, metavar="СТОЛБЕЦ", help="по какому столбцу делить")
    p.add_argument("--sheets", action="store_true", help="не отдельные файлы, а листы одной книги")
    p.add_argument("--sheet", help="брать только этот лист")
    p.add_argument("-o", "--out", type=Path, help="папка (или файл при --sheets); "
                                                  "по умолчанию output/по_<столбцу>")
    return parser


def load(inputs, sheet=None) -> pd.DataFrame:
    """Файлы → одна таблица, приведённая в порядок (без удаления строк: для подсчётов
    повторы могут быть настоящими продажами). Из одной таблицы столбцы «Файл»/«Лист» не добавляются."""
    tables = [t for f in expand_inputs(inputs) for t in read_tables(f, sheet)]
    if not tables:
        raise TableError("Во входных файлах нет данных")
    frame, _ = merge(tables, with_source=len(tables) > 1)
    frame, _ = clean(frame, drop_duplicates=False)
    return frame


def run_summary(args) -> int:
    frame = load(args.inputs, args.sheet)
    table = summarize(frame, args.by, args.values, args.agg, args.pivot)
    print(f"Строк в данных: {len(frame)}, групп: {len(table) - 1}")
    path = save_workbook([Sheet("Сводка", table, total_row=True), Sheet("Данные", frame)], args.out)
    print(f"Готово: {path}")
    return 0


def first_table(path: str, sheet: str | None) -> pd.DataFrame:
    tables = read_tables(Path(path), sheet)
    if not tables:
        raise TableError(f"{Path(path).name}: нет данных")
    frame, _ = clean(tables[0].frame, drop_duplicates=False)
    return frame


def run_compare(args) -> int:
    result = compare(first_table(args.old, args.sheet), first_table(args.new, args.sheet), args.key)
    for line in result.lines():
        print(line)
    print(f"Готово: {save_workbook(result.sheets(), args.out)}")
    return 0


def run_split(args) -> int:
    frame = load(args.inputs, args.sheet)
    column, parts = split(frame, args.by)
    out = args.out or OUTPUT / f"по_{safe_filename(column, set())}"
    if args.sheets:
        path = save_workbook([Sheet(name, part) for name, part in parts.items()], out)
        print(f"Листов: {len(parts)}. Готово: {path}")
        return 0
    taken: set[str] = set()
    for name, part in parts.items():
        stem = Path(name).stem if Path(name).suffix.lower() in (".xlsx", ".xlsm", ".csv") else name
        path = save_workbook([Sheet(name, part)], Path(out) / f"{safe_filename(stem, taken)}.xlsx")
        print(f"  {path.name}: {len(part)} строк")
    print(f"Файлов: {len(parts)}. Папка: {out}")
    return 0


def run_merge(args) -> int:
    tables = [t for f in expand_inputs(args.inputs) for t in read_tables(f, args.sheet)]
    if not tables:
        raise TableError("Во входных файлах нет данных")
    print(f"Таблиц: {len(tables)}")
    merged, report = merge(tables, with_source=not args.no_source)
    for line in report.lines():
        print(line)
    if args.clean:
        merged, cleaned = clean(merged, drop_duplicates=not args.keep_duplicates)
        print("Очистка:")
        for line in cleaned.lines():
            print(f"  {line}")
    path = save_workbook([Sheet("Данные", merged)], args.out)
    print(f"Готово: {path}")
    return 0


def run_clean(args) -> int:
    path = Path(args.input)
    tables = read_tables(path, args.sheet)
    if not tables:
        raise TableError(f"{path.name}: нет данных")
    sheets = []
    for table in tables:
        frame, report = clean(table.frame, drop_duplicates=not args.keep_duplicates)
        print(table.label)
        for line in report.lines():
            print(f"  {line}")
        sheets.append(Sheet(table.sheet or path.stem, frame))
    out = args.out or OUTPUT / f"{path.stem}_чисто.xlsx"
    print(f"Готово: {save_workbook(sheets, out)}")
    return 0


def utf8_output() -> None:
    """Вывод в файл на Windows идёт в cp1251, где нет символов вроде ⚠ — переключаем на UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if (getattr(stream, "encoding", "") or "").lower().replace("-", "") != "utf8" \
                and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


COMMANDS = {"merge": run_merge, "clean": run_clean, "summary": run_summary,
            "compare": run_compare, "split": run_split}


def main(argv=None) -> int:
    utf8_output()
    args = build_parser().parse_args(argv)
    try:
        return COMMANDS[args.command](args)
    except TableError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
