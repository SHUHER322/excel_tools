import pandas as pd
import pytest

from exceltools.tables import (TableError, expand_inputs, frame_from_rows, read_tables, sniff_delimiter,
                               unique_names)


def test_header_below_report_title(sales_xlsx):
    tables = read_tables(sales_xlsx)
    assert [t.sheet for t in tables] == ["Октябрь", "Возвраты"]
    october = tables[0].frame
    assert list(october.columns) == ["Дата", "Менеджер", "Товар", "Сумма"]
    assert october["Сумма"].tolist() == [420, 1250]


def test_one_sheet(sales_xlsx):
    tables = read_tables(sales_xlsx, sheet="Возвраты")
    assert len(tables) == 1 and tables[0].frame["Сумма"].tolist() == [-1250]


def test_missing_sheet(sales_xlsx):
    with pytest.raises(TableError, match="Продажи_Томск.xlsx"):
        read_tables(sales_xlsx, sheet="Ноябрь")


def test_csv_cp1251_semicolon(sales_csv):
    frame = read_tables(sales_csv)[0].frame
    assert list(frame.columns) == ["товар", "СУММА, руб.", "Дата", "Менеджер"]
    assert frame["СУММА, руб."].tolist() == [12500, 420]
    assert frame["Дата"].iloc[0] == pd.Timestamp(2026, 10, 4)


def test_csv_utf8_comma(tmp_path):
    path = tmp_path / "data.csv"
    path.write_text("name,price\nA,\"1,5\"\nB,2\n", encoding="utf-8-sig")
    frame = read_tables(path)[0].frame
    assert frame["price"].tolist() == [1.5, 2.0]


@pytest.mark.parametrize("text, expected", [
    ("a;b;c\n1;2;3\n", ";"),
    ("a,b,c\n1,2,3\n", ","),
    ("a\tb\n1\t2\n", "\t"),
    ("name;comment\nA;да, конечно\nB;нет\n", ";"),
])
def test_sniff_delimiter(text, expected):
    assert sniff_delimiter(text) == expected


def test_unique_names():
    assert unique_names([" Сумма ", None, "Сумма", 2026.0, ""]) == [
        "Сумма", "Столбец 2", "Сумма (2)", "2026", "Столбец 5"]


def test_frame_from_rows_drops_empty_rows_and_columns():
    raw = pd.DataFrame([
        ["Отчёт", None, None, None],
        [None, None, None, None],
        ["A", None, "B", "C"],
        [1, None, 2, 3],
        [None, None, None, None],
        [4, None, 5, "   "],
    ], dtype=object)
    frame = frame_from_rows(raw)
    assert list(frame.columns) == ["A", "B", "C"]
    assert len(frame) == 2


def test_expand_inputs(tmp_path, sales_xlsx, sales_csv):
    (tmp_path / "~$Продажи_Томск.xlsx").write_bytes(b"lock")      # файл блокировки Excel
    (tmp_path / "заметки.txt.bak").write_text("x")
    assert expand_inputs([tmp_path]) == [sales_csv, sales_xlsx]      # без учёта регистра
    assert expand_inputs([str(tmp_path / "*.csv")]) == [sales_csv]
    assert expand_inputs([sales_csv, tmp_path]).count(sales_csv) == 1


def test_expand_inputs_errors(tmp_path):
    with pytest.raises(TableError, match="Не найден"):
        expand_inputs([tmp_path / "нет.xlsx"])
    with pytest.raises(TableError, match="Нет таблиц"):
        expand_inputs([tmp_path])


def test_unsupported_format(tmp_path):
    path = tmp_path / "old.xls"
    path.write_bytes(b"x")
    with pytest.raises(TableError, match="поддерживаются"):
        read_tables(path)
