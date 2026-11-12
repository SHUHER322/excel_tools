from datetime import datetime

import pandas as pd
import pytest

from exceltools.values import column_key, convert_types, normalize_text, parse_date, parse_number


@pytest.mark.parametrize("text, expected", [
    ("1200", 1200),
    ("1 200", 1200),
    ("12 500,50", 12500.5),
    ("12 500,50 ₽", 12500.5),
    ("-3,5%", -3.5),
    ("420.00", 420.0),
    ("1 890,00 руб.", 1890.0),
])
def test_parse_number(text, expected):
    assert parse_number(text) == expected


@pytest.mark.parametrize("text", ["abc", "1.234.567", "12-500", "", "1 23", None])
def test_parse_number_rejects(text):
    assert parse_number(text) is None


def test_parse_date():
    assert parse_date("05.10.2026") == datetime(2026, 10, 5)
    assert parse_date("2026-10-05") == datetime(2026, 10, 5)
    assert parse_date("05.10.26") == datetime(2026, 10, 5)
    assert parse_date("октябрь") is None


def test_normalize_and_key():
    assert normalize_text("  Иванова  А.  ") == "Иванова А."
    assert column_key("  Сумма, ₽ ") == column_key("СУММА") == "сумма"
    assert column_key("Сумма, руб.") == column_key("Сумма (₽)") == "сумма"
    assert column_key("Кол-во, шт.") == column_key("кол во")
    assert column_key("руб") == "руб"
    assert column_key("Ёмкость") == column_key("емкость")


def test_convert_types_numbers_and_dates():
    frame = pd.DataFrame({
        "Дата": ["05.10.2026", "06.10.2026", None],
        "Сумма": ["1 200", "300", "5"],
        "Цена": ["12 500,50", "180,00", "1 890,00"],
    })
    result, changed = convert_types(frame)
    assert changed == {"numbers": ["Сумма", "Цена"], "dates": ["Дата"]}
    assert result["Сумма"].tolist() == [1200, 300, 5]
    assert result["Цена"].tolist() == [12500.5, 180.0, 1890.0]
    assert result["Дата"].iloc[0] == pd.Timestamp(2026, 10, 5)


def test_codes_with_leading_zero_stay_text():
    result, changed = convert_types(pd.DataFrame({"Артикул": ["00123", "00124", "00125"]}))
    assert changed["numbers"] == []
    assert result["Артикул"].tolist() == ["00123", "00124", "00125"]


def test_mostly_text_column_is_not_converted():
    result, changed = convert_types(pd.DataFrame({"Комментарий": ["1", "нет", "уточнить", "2"]}))
    assert changed["numbers"] == []


def test_unreadable_values_are_kept():
    """10 значений: 9 чисел и «нет» — числа становятся числами, «нет» не теряется."""
    values = [str(i) for i in range(1, 10)] + ["нет"]
    result, changed = convert_types(pd.DataFrame({"Кол-во": values}))
    assert changed["numbers"] == ["Кол-во"]
    assert result["Кол-во"].tolist() == list(range(1, 10)) + ["нет"]


def test_numbers_from_excel_mixed_with_text_numbers():
    result, changed = convert_types(pd.DataFrame({"Цена": [420, 1250, "12 500,00"]}))
    assert result["Цена"].tolist() == [420, 1250, 12500]
