from types import SimpleNamespace

from backend.app.services.quote_math import validate_quote_items


def _item(**changes):
    values = {
        "product_name": "현수막",
        "width_mm": None,
        "height_mm": None,
        "quantity": None,
        "unit": None,
        "paper": None,
        "print_sides": None,
        "material": None,
        "unit_price": None,
        "confirmed": False,
    }
    values.update(changes)
    return SimpleNamespace(**values)


def test_product_name_only_passes_structural_validation():
    assert validate_quote_items([_item()]) == []


def test_missing_product_name_is_rejected():
    errors = validate_quote_items([
        _item(product_name="  ")
    ])
    assert errors == [
        "1번째 품목명이 비어 있습니다."
    ]


def test_empty_item_list_is_rejected():
    assert validate_quote_items([]) == [
        "견적서에 입력할 품목이 없습니다."
    ]


def test_multiple_items_report_only_blank_product_names():
    errors = validate_quote_items([
        _item(product_name="명함"),
        _item(product_name=""),
        _item(product_name="배너"),
    ])
    assert errors == [
        "2번째 품목명이 비어 있습니다."
    ]
