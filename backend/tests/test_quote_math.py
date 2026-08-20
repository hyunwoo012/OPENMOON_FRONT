from types import SimpleNamespace

from backend.app.services.quote_math import validate_quote_items


def _item(**changes):
    values = {
        "product_name": "사용자정의품목",
        "normalized_product": "사용자정의품목",
        "specification": "100*200mm",
        "quantity": 1,
        "unit": "개",
        "paper": None,
        "print_sides": None,
        "material": None,
        "spec_attributes": {},
        "unit_price": 10000,
        "confirmed": True,
    }
    values.update(changes)
    return SimpleNamespace(**values)


def test_complete_custom_item_passes():
    assert validate_quote_items([_item()]) == []


def test_missing_product_name_is_rejected():
    assert validate_quote_items([
        _item(product_name="  ")
    ]) == ["1번째 품목명이 비어 있습니다."]


def test_empty_item_list_is_rejected():
    assert validate_quote_items([]) == [
        "견적서에 입력할 품목이 없습니다."
    ]


def test_required_common_fields_are_rejected():
    errors = validate_quote_items([
        _item(
            quantity=None,
            unit="",
            unit_price=None,
            specification="",
        )
    ])
    joined = " / ".join(errors)
    assert "수량" in joined
    assert "단위" in joined
    assert "확정 단가" in joined
    assert "규격/사양" in joined


def test_catalog_product_requires_all_displayed_fields():
    item = _item(
        product_name="명함",
        normalized_product="명함",
        specification="90*50",
        paper="코팅명함",
        quantity=200,
        unit="매",
        print_sides="단면4도",
        spec_attributes={
            "finishing": "",
            "delivery_method": "납품",
        },
        unit_price=10000,
    )

    errors = validate_quote_items([item])
    assert any("후가공" in error for error in errors)

    item.spec_attributes["finishing"] = "귀도리"
    assert validate_quote_items([item]) == []
