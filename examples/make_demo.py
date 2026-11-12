"""Демо-данные: выгрузки продаж трёх филиалов — «как бывает в жизни».

Все данные выдуманы. Каждый файл «испорчен» по-своему, как настоящие выгрузки:
  Продажи_Томск.xlsx        — сверху заголовок отчёта, шапка не в первой строке, два листа
  Продажи_Новосибирск.xlsx  — другой порядок столбцов, лишние пробелы, повторы строк
  продажи_омск.csv          — Windows-1251, «;», числа «12 500,00» и даты текстом

И два прайса поставщика для сравнения (examples/compare/): в октябрьском поменялись цены
и остатки, два товара добавились, один пропал, строки идут в другом порядке.

Запуск: python examples/make_demo.py  → папки examples/data/ и examples/compare/
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).parent / "data"
PRODUCTS = [  # товар, категория, цена
    ("Бумага А4, 500 л.", "Канцтовары", 420),
    ("Ручка шариковая, 10 шт.", "Канцтовары", 180),
    ("Папка-регистратор", "Канцтовары", 260),
    ("Картридж для принтера", "Оргтехника", 3900),
    ("Мышь беспроводная", "Оргтехника", 1250),
    ("Клавиатура", "Оргтехника", 1890),
    ("Кресло офисное", "Мебель", 12500),
    ("Стол письменный", "Мебель", 15900),
    ("Кофе в зёрнах, 1 кг", "Продукты", 1490),
    ("Вода питьевая, 19 л", "Продукты", 390),
]
MANAGERS = {
    "Томск": ["Иванова А.", "Петров С.", "Сидорова Е."],
    "Новосибирск": ["Кузнецов Д.", "Смирнова О.", "Попов И.", "Волкова Н."],
    "Омск": ["Морозов К.", "Лебедева Т."],
}
START = date(2026, 10, 1)


def sales(city: str, rows: int, seed: int) -> list[dict]:
    rnd = random.Random(seed)
    result = []
    while len(result) < rows:
        product, category, price = rnd.choice(PRODUCTS)
        qty = rnd.choice([1, 1, 1, 2, 2, 3, 5, 10])
        row = {
            "Дата": START + timedelta(days=rnd.randrange(31)),
            "Менеджер": rnd.choice(MANAGERS[city]),
            "Товар": product,
            "Категория": category,
            "Кол-во": qty,
            "Цена": price,
            "Сумма": price * qty,
        }
        if row not in result:                   # повторы появятся только там, где задуманы
            result.append(row)
    return sorted(result, key=lambda r: r["Дата"])


def tomsk() -> None:
    """Сверху заголовок и пустая строка, два листа: «Октябрь» и «Возвраты»."""
    book = Workbook()
    ws = book.active
    ws.title = "Октябрь"
    ws.append(["Отчёт о продажах филиала Томск, октябрь 2026"])
    ws.append(["Выгружено из учётной системы"])
    ws.append([])
    columns = ["Дата", "Менеджер", "Товар", "Категория", "Кол-во", "Цена", "Сумма"]
    ws.append(columns)
    for row in sales("Томск", 45, seed=1):
        ws.append([row[c] for c in columns])
    returns = book.create_sheet("Возвраты")
    returns.append(columns)
    for row in sales("Томск", 4, seed=11):
        row["Кол-во"], row["Сумма"] = -row["Кол-во"], -row["Сумма"]
        returns.append([row[c] for c in columns])
    book.save(OUT / "Продажи_Томск.xlsx")


def novosibirsk() -> None:
    """Другой порядок и регистр столбцов, пробелы в ячейках, пустые строки, повторы."""
    book = Workbook()
    ws = book.active
    ws.title = "Продажи"
    columns = ["  менеджер ", "ТОВАР", "Категория", "Дата", "Цена", "Кол-во", "Сумма, ₽"]
    ws.append(columns)
    rows = sales("Новосибирск", 55, seed=2)
    rows += rows[5:8]                                   # выгрузили дважды
    for i, row in enumerate(rows):
        manager = f" {row['Менеджер']}  " if i % 4 == 0 else row["Менеджер"]
        ws.append([manager, row["Товар"], row["Категория"], row["Дата"], row["Цена"], row["Кол-во"],
                   row["Сумма"]])
        if i % 20 == 19:
            ws.append([])                               # пустые строки посреди данных
    book.save(OUT / "Продажи_Новосибирск.xlsx")


def omsk() -> None:
    """CSV из «1С»: Windows-1251, точка с запятой, «12 500,00» и даты текстом."""
    with open(OUT / "продажи_омск.csv", "w", encoding="cp1251", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Дата", "Менеджер", "Товар", "Категория", "Кол-во", "Цена", "Сумма"])
        for row in sales("Омск", 30, seed=3):
            writer.writerow([
                row["Дата"].strftime("%d.%m.%Y"), row["Менеджер"], row["Товар"], row["Категория"],
                row["Кол-во"], f"{row['Цена']:,.2f}".replace(",", " ").replace(".", ","),
                f"{row['Сумма']:,.2f}".replace(",", " ").replace(".", ","),
            ])


PRICE_LIST = [  # артикул, наименование, цена, остаток
    ("KC-1001", "Бумага А4, 500 л.", 420, 120),
    ("KC-1002", "Ручка шариковая, 10 шт.", 180, 300),
    ("KC-1003", "Папка-регистратор", 260, 85),
    ("OT-2001", "Картридж для принтера", 3900, 14),
    ("OT-2002", "Мышь беспроводная", 1250, 40),
    ("OT-2003", "Клавиатура", 1890, 22),
    ("MB-3001", "Кресло офисное", 12500, 6),
    ("MB-3002", "Стол письменный", 15900, 3),
    ("PR-4001", "Кофе в зёрнах, 1 кг", 1490, 50),
    ("PR-4002", "Вода питьевая, 19 л", 390, 200),
]


def price_lists() -> None:
    """Сентябрь и октябрь: изменились цены и остатки, KC-1003 пропал, добавились два товара."""
    folder = OUT.parent / "compare"
    folder.mkdir(parents=True, exist_ok=True)
    columns = ["Артикул", "Наименование", "Цена", "Остаток"]

    book = Workbook()
    ws = book.active
    ws.title = "Прайс"
    ws.append(columns)
    for row in PRICE_LIST:
        ws.append(list(row))
    book.save(folder / "Прайс_сентябрь.xlsx")

    changes = {"KC-1001": (450, 95), "OT-2001": (4100, 14), "MB-3001": (12500, 0), "PR-4001": (1590, 35)}
    october = []
    for sku, name, price, stock in PRICE_LIST:
        if sku == "KC-1003":
            continue
        price, stock = changes.get(sku, (price, stock))
        october.append((sku, name, price, stock))
    october += [("OT-2004", "Веб-камера", 2790, 18), ("PR-4003", "Чай чёрный, 100 пак.", 340, 60)]
    october.sort(key=lambda r: r[1])                    # поставщик отсортировал по названию
    book = Workbook()
    ws = book.active
    ws.title = "Прайс"
    ws.append(["артикул", "Наименование", "Цена, руб.", "Остаток, шт."])
    for row in october:
        ws.append(list(row))
    book.save(folder / "Прайс_октябрь.xlsx")


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    tomsk()
    novosibirsk()
    omsk()
    price_lists()
    print(f"Демо-данные: {OUT} и {OUT.parent / 'compare'}")
