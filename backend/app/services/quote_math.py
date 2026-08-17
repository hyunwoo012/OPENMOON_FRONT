from __future__ import annotations

from typing import Any


def calculate_supply_amount(
    quantity: float | int | None,
    unit_price: float | int | None,
) -> int | None:
    """견적 공급금액의 단일 계산 규칙: 수량 × 단가."""
    if quantity is None or unit_price is None:
        return None
    return int(round(float(quantity) * int(unit_price)))


def normalize_item_amount(item: Any) -> int | None:
    amount = calculate_supply_amount(
        getattr(item, "quantity", None),
        getattr(item, "unit_price", None),
    )
    item.amount = amount
    return amount


def quote_total(items: list[Any]) -> int | None:
    if not items:
        return None
    amounts: list[int] = []
    for item in items:
        amount = calculate_supply_amount(
            getattr(item, "quantity", None),
            getattr(item, "unit_price", None),
        )
        if amount is None:
            return None
        amounts.append(amount)
    return sum(amounts)


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def validate_quote_items(items: list[Any]) -> list[str]:
    errors: list[str] = []
    if not items:
        return ["견적서에 입력할 품목이 없습니다."]

    for index, item in enumerate(items, start=1):
        product = str(getattr(item, "product_name", "") or "").strip()
        label = product or f"{index}번째 품목"

        if not product:
            errors.append(f"{index}번째 품목명이 비어 있습니다.")
        required_fields = (
            ("width_mm", "가로"),
            ("height_mm", "세로"),
            ("quantity", "수량"),
            ("unit", "단위"),
            ("paper", "용지"),
            ("print_sides", "단면·양면"),
            ("material", "재질"),
            ("unit_price", "확정단가"),
        )
        missing = [
            display_name
            for field_name, display_name in required_fields
            if _is_blank(getattr(item, field_name, None))
        ]
        if missing:
            errors.append(f"{label}: {', '.join(missing)} 항목을 입력해 주세요.")

    return errors
