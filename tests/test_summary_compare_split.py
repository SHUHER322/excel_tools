from datetime import datetime

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from exceltools.__main__ import main
from exceltools.compare import compare, same
from exceltools.split import safe_filename, split
from exceltools.summary import parse_group, summarize
from exceltools.tables import TableError


@pytest.fixture
def sales():
    return pd.DataFrame({
        "Дата": pd.to_datetime(["2026-09-30", "2026-10-01", "2026-10-02", "2026-10-15"]),
        "Менеджер": ["Иванова А.", "Петров С.", "Иванова А.", "Петров С."],
        "Категория": ["Мебель", "Канцтовары", "Канцтовары", "Мебель"],
        "Кол-во": [1, 2, 3, 1],
        "Сумма": [12500, 360, 540, 15900],
    })


# --- сводный отчёт ---

def test_summary_by_manager(sales):
    table = summarize(sales, ["менеджер"])
    assert list(table.columns) == ["Менеджер", "Строк", "Кол-во", "Сумма"]
    assert table["Менеджер"].tolist() == ["Петров С.", "Иванова А.", "Итого"]   # от большей суммы
    assert table["Сумма"].tolist() == [16260, 13040, 29300]
    assert table["Строк"].tolist() == [2, 2, 4]


def test_summary_by_month(sales):
    table = summarize(sales, ["Дата:месяц"], ["Сумма"])
    assert table["Дата (месяц)"].tolist() == ["09.2026", "10.2026", "Итого"]
    assert table["Сумма"].tolist() == [12500, 16800, 29300]


def test_summary_mean(sales):
    table = summarize(sales, ["Категория"], ["Сумма"], agg="mean")
    assert table.set_index("Категория").loc["Итого", "Сумма"] == 29300 / 4


def test_summary_pivot(sales):
    table = summarize(sales, ["Менеджер"], ["Сумма"], pivot="Категория")
    assert list(table.columns) == ["Менеджер", "Канцтовары", "Мебель", "Итого"]
    total = table.set_index("Менеджер").loc["Итого"]
    assert (total["Канцтовары"], total["Мебель"], total["Итого"]) == (900, 28400, 29300)


def test_summary_errors(sales):
    with pytest.raises(TableError, match="Нет столбца «Филиал». Есть: «Дата»"):
        summarize(sales, ["Филиал"])
    with pytest.raises(TableError, match="Неизвестный период"):
        summarize(sales, ["Дата:сезон"])
    with pytest.raises(TableError, match="нет дат"):
        summarize(sales, ["Менеджер:месяц"])
    assert parse_group(" Дата : Месяц ") == ("Дата", "месяц")


# --- сравнение ---

def price(rows, columns=("Артикул", "Наименование", "Цена", "Остаток")):
    return pd.DataFrame(rows, columns=list(columns))


def test_compare():
    old = price([("A-1", "Бумага", 420, 10), ("A-2", "Ручка", 180, 5), ("A-3", "Папка", 260, 1)])
    new = price([("A-2", "Ручка", 180, 5), ("A-1", "Бумага", 450, 10), ("A-4", "Мышь", 1250, 3)],
                columns=("артикул", "Наименование", "Цена, руб.", "Остаток"))
    result = compare(old, new, ["Артикул"])
    assert result.changes.to_dict("records") == [
        {"артикул": "A-1", "Наименование": "Бумага", "Столбец": "Цена, руб.", "Было": 420, "Стало": 450,
         "Разница": 30}]
    assert result.added["артикул"].tolist() == ["A-4"]
    assert result.removed["Наименование"].tolist() == ["Папка"]
    assert result.full["Статус"].tolist() == ["", "изменена", "новая", "удалена"]
    assert (0, None, "added") not in result.marks and (1, "Цена, руб.", "changed") in result.marks
    assert (2, None, "added") in result.marks and (3, None, "removed") in result.marks
    assert result.changed_rows == 1


def test_compare_numeric_keys_and_columns():
    old = price([(101, "Бумага", 420, 10)], columns=("Код", "Наименование", "Цена", "Цвет"))
    new = price([("101", "Бумага ", 420.0, 10)], columns=("Код", "Наименование", "Цена", "Размер"))
    result = compare(old, new, ["Код"])
    assert len(result.changes) == 0 and len(result.added) == 0
    assert result.only_old == ["Цвет"] and result.only_new == ["Размер"]


def test_compare_duplicate_keys():
    old = price([("A-1", "Бумага", 420, 1), ("A-1", "Бумага", 430, 1)])
    result = compare(old, old.copy(), ["Артикул"])
    assert result.duplicate_keys == ["A-1"]
    assert any("повторяется" in line for line in result.lines())


def test_compare_needs_key():
    with pytest.raises(TableError, match="ключ"):
        compare(price([]), price([]), [])


@pytest.mark.parametrize("a, b, expected", [
    (None, float("nan"), True), ("", None, True), (1, 1.0, True), (0.1 + 0.2, 0.3, True),
    (" Бумага", "Бумага", True), (1, 2, False), ("x", None, False),
    (datetime(2026, 10, 1), pd.Timestamp(2026, 10, 1), True),
])
def test_same(a, b, expected):
    assert same(a, b) is expected


# --- разбивка ---

def test_split(sales):
    column, parts = split(sales, "менеджер")
    assert column == "Менеджер"
    assert list(parts) == ["Иванова А.", "Петров С."]
    assert parts["Петров С."]["Сумма"].tolist() == [360, 15900]


def test_safe_filename():
    taken = set()
    assert safe_filename('ООО "Ромашка"/Склад', taken) == "ООО _Ромашка__Склад"
    assert safe_filename("Иванов", taken) == "Иванов"
    assert safe_filename("иванов", taken) == "иванов (2)"
    assert safe_filename("CON", taken) == "_CON"
    assert safe_filename("  ", taken) == "без_названия"


# --- командная строка ---

def write_book(path, rows):
    book = Workbook()
    for row in rows:
        book.active.append(row)
    book.save(path)
    return path


def test_cli_summary(tmp_path, sales_xlsx):
    out = tmp_path / "сводка.xlsx"
    assert main(["summary", str(sales_xlsx), "--by", "Менеджер", "--values", "Сумма", "-o", str(out)]) == 0
    ws = load_workbook(out)["Сводка"]
    assert [c.value for c in ws[1]] == ["Менеджер", "Строк", "Сумма"]
    assert [c.value for c in ws[ws.max_row]] == ["Итого", 3, 420]
    assert ws.cell(ws.max_row, 1).font.bold


def test_cli_compare(tmp_path):
    old = write_book(tmp_path / "old.xlsx", [["Артикул", "Цена"], ["A-1", 100], ["A-2", 200]])
    new = write_book(tmp_path / "new.xlsx", [["Артикул", "Цена"], ["A-1", 110], ["A-3", 300]])
    out = tmp_path / "сравнение.xlsx"
    assert main(["compare", str(old), str(new), "--key", "Артикул", "-o", str(out)]) == 0
    book = load_workbook(out)
    assert book.sheetnames == ["Сравнение", "Изменения", "Новые", "Удалённые"]
    ws = book["Сравнение"]
    assert [c.value for c in ws["A"]][1:] == ["изменена", "новая", "удалена"]
    assert ws["C2"].fill.fgColor.rgb.endswith("FFF2CC")


def test_cli_split(tmp_path, sales_xlsx):
    out = tmp_path / "по_менеджерам"
    assert main(["split", str(sales_xlsx), "--sheet", "Октябрь", "--by", "Менеджер", "-o", str(out)]) == 0
    assert sorted(p.name for p in out.iterdir()) == ["Иванова А.xlsx", "Петров С.xlsx"]
    book = tmp_path / "менеджеры.xlsx"
    assert main(["split", str(sales_xlsx), "--by", "Менеджер", "--sheets", "-o", str(book)]) == 0
    assert load_workbook(book).sheetnames == ["Иванова А.", "Петров С."]
