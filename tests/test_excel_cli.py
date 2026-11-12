from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

from exceltools import excel
from exceltools.__main__ import main
from exceltools.excel import Sheet, build_workbook, safe_sheet_name, save_workbook


def test_workbook_formatting():
    frame = pd.DataFrame({
        "Дата": [pd.Timestamp(2026, 10, 1), pd.NaT],
        "Сумма": pd.array([1200, None], dtype="Int64"),
        "Цена": [12.5, float("nan")],
        "Товар": ["Бумага", None],
    })
    ws = build_workbook([Sheet("Данные", frame)])["Данные"]
    assert [c.value for c in ws[1]] == ["Дата", "Сумма", "Цена", "Товар"]
    assert ws["A2"].value == datetime(2026, 10, 1) and ws["A2"].number_format == excel.DATE_FORMAT
    assert ws["B2"].value == 1200 and ws["B2"].number_format == excel.INT_FORMAT
    assert ws["C2"].number_format == excel.FLOAT_FORMAT
    assert [c.value for c in ws[3]] == [None, None, None, None]
    assert ws.freeze_panes == "A2" and ws.auto_filter.ref == "A1:D3"
    assert ws["A1"].font.bold


def test_total_row_and_marks():
    frame = pd.DataFrame({"Товар": ["A", "B", "Итого"], "Сумма": [1, 2, 3]})
    ws = build_workbook([Sheet("Сводка", frame, total_row=True, marks=[(0, "Сумма", "changed"),
                                                                      (1, None, "added")])])["Сводка"]
    assert ws["A4"].font.bold and ws.auto_filter.ref == "A1:B3"
    assert ws["B2"].fill.fgColor.rgb.endswith("FFF2CC") and not ws["A2"].fill.fgColor.rgb.endswith("FFF2CC")
    assert ws["A3"].fill.fgColor.rgb.endswith("E2EFDA") and ws["B3"].fill.fgColor.rgb.endswith("E2EFDA")


def test_safe_sheet_name():
    taken = set()
    assert safe_sheet_name("Продажи: октябрь/ноябрь", taken) == "Продажи_ октябрь_ноябрь"
    long = "Очень длинное название листа больше 31 символа"
    first = safe_sheet_name(long, taken)
    second = safe_sheet_name(long, taken)
    assert len(first) == 31 and len(second) <= 31 and first != second


def test_save_when_file_is_open(tmp_path, monkeypatch):
    calls = []
    original = excel.Workbook.save

    def locked(self, path):
        calls.append(path)
        if len(calls) == 1:
            raise PermissionError("файл открыт")
        original(self, path)

    monkeypatch.setattr(excel.Workbook, "save", locked)
    path = save_workbook([Sheet("A", pd.DataFrame({"x": [1]}))], tmp_path / "итог.xlsx")
    assert path != tmp_path / "итог.xlsx" and path.name.startswith("итог_") and path.exists()


def test_cli_merge(tmp_path, sales_xlsx, sales_csv):
    out = tmp_path / "out" / "всё.xlsx"
    assert main(["merge", str(tmp_path), "--clean", "-o", str(out)]) == 0
    frame = pd.read_excel(out)
    assert len(frame) == 5
    assert frame["Сумма"].sum() == 420 + 1250 - 1250 + 12500 + 420


def test_cli_clean(tmp_path, sales_csv):
    out = tmp_path / "чисто.xlsx"
    assert main(["clean", str(sales_csv), "-o", str(out)]) == 0
    book = load_workbook(out)
    ws = book[book.sheetnames[0]]
    assert ws["D3"].value == "Лебедева Т."


def test_cli_error(tmp_path):
    assert main(["merge", str(tmp_path / "нет")]) == 2
