from types import SimpleNamespace

from backend.app.services.quote_math import validate_quote_items


def _complete_item(**changes):
    values = {
        "product_name": "현수막",
        "width_mm": 5000,
        "height_mm": 600,
        "quantity": 1,
        "unit": "장",
        "paper": "해당 없음",
        "print_sides": "단면",
        "material": "현수막",
        "unit_price": 30000,
        "confirmed": True,
    }
    values.update(changes)
    return SimpleNamespace(**values)


def test_complete_quote_item_passes_required_field_validation():
    assert validate_quote_items([_complete_item()]) == []


def test_all_missing_quote_fields_are_reported():
    errors = validate_quote_items([
        _complete_item(
            width_mm=None,
            height_mm=None,
            quantity=None,
            unit=" ",
            paper=None,
            print_sides="",
            material=None,
            unit_price=None,
            confirmed=False,
        )
    ])
    message = " ".join(errors)
    for label in ("가로", "세로", "수량", "단위", "용지", "단면·양면", "재질", "확정단가"):
        assert label in message


def test_entered_unit_price_passes_even_without_internal_confirmation_flag():
    assert validate_quote_items([_complete_item(confirmed=False)]) == []
