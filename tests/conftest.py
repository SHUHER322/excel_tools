import csv
from datetime import datetime

import pytest
from openpyxl import Workbook


@pytest.fixture
def sales_xlsx(tmp_path):
    """Книга с заголовком отчёта над шапкой и вторым листом."""
    book = Workbook()
    ws = book.active
    ws.title = "Октябрь"
    ws.append(["Отчёт о продажах, октябрь 2026"])
    ws.append([])
    ws.append(["Дата", "Менеджер", "Товар", "Сумма"])
    ws.append([datetime(2026, 10, 1), "Иванова А.", "Бумага", 420])
    ws.append([datetime(2026, 10, 2), "Петров С.", "Мышь", 1250])
    other = book.create_sheet("Возвраты")
    other.append(["Дата", "Менеджер", "Товар", "Сумма"])
    other.append([datetime(2026, 10, 3), "Иванова А.", "Мышь", -1250])
    path = tmp_path / "Продажи_Томск.xlsx"
    book.save(path)
    return path


@pytest.fixture
def sales_csv(tmp_path):
    """CSV как из 1С: Windows-1251, «;», числа с пробелами и запятой, другой порядок столбцов."""
    path = tmp_path / "продажи_омск.csv"
    with open(path, "w", encoding="cp1251", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["  товар ", "СУММА, руб.", "Дата", "Менеджер"])
        writer.writerow(["Кресло", "12 500,00", "04.10.2026", "Морозов К."])
        writer.writerow(["Бумага", "420,00", "05.10.2026", " Лебедева  Т. "])
    return path
