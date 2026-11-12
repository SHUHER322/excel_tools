import pandas as pd

from exceltools.clean import clean
from exceltools.merge import display_name, merge
from exceltools.tables import read_tables


def test_clean():
    frame = pd.DataFrame({
        "Менеджер": [" Иванова  А.", "Иванова А.", None, "Петров С."],
        "Сумма": ["1 200", "1 200", None, "300,50"],
        "Пусто": [None, None, None, None],
    }, dtype=object)
    result, report = clean(frame)
    assert list(result.columns) == ["Менеджер", "Сумма"]
    assert result["Менеджер"].tolist() == ["Иванова А.", "Петров С."]
    assert result["Сумма"].tolist() == [1200.0, 300.5]
    assert (report.rows_before, report.rows_after) == (4, 2)
    assert (report.empty_rows, report.duplicates, report.trimmed_cells) == (1, 1, 1)
    assert report.numbers == ["Сумма"]
    assert "Удалено повторов: 1" in report.lines()


def test_clean_keep_duplicates():
    frame = pd.DataFrame({"A": ["x", "x"]}, dtype=object)
    result, report = clean(frame, drop_duplicates=False)
    assert len(result) == 2 and report.duplicates == 0


def test_merge_aligns_columns_by_name(sales_xlsx, sales_csv):
    tables = read_tables(sales_xlsx) + read_tables(sales_csv)
    merged, report = merge(tables)
    assert list(merged.columns) == ["Файл", "Лист", "Дата", "Менеджер", "Товар", "Сумма"]
    assert len(merged) == report.rows == 5
    csv_rows = merged[merged["Файл"] == sales_csv.name]
    assert csv_rows["Товар"].tolist() == ["Кресло", "Бумага"]
    assert csv_rows["Сумма"].tolist() == [12500, 420]
    assert merged["Лист"].tolist()[:3] == ["Октябрь", "Октябрь", "Возвраты"]
    assert pd.isna(merged["Лист"].iloc[-1])


def test_merge_reports_missing_columns(sales_xlsx, tmp_path):
    path = tmp_path / "без_суммы.csv"
    path.write_text("Дата;Менеджер;Товар\n01.10.2026;Иванова А.;Бумага\n", encoding="utf-8")
    merged, report = merge(read_tables(sales_xlsx, sheet="Октябрь") + read_tables(path))
    assert report.missing == {"без_суммы.csv": ["Сумма"]}
    assert pd.isna(merged["Сумма"].iloc[-1])
    assert any("нет столбцов: Сумма" in line for line in report.lines())


def test_merge_without_source(sales_xlsx):
    merged, _ = merge(read_tables(sales_xlsx), with_source=False)
    assert "Файл" not in merged.columns and "Лист" not in merged.columns


def test_display_name_majority():
    assert display_name(["ТОВАР", "Товар", "Товар"]) == "Товар"
    assert display_name(["Сумма, ₽", "Сумма"]) == "Сумма, ₽"       # поровну — как в первом файле
