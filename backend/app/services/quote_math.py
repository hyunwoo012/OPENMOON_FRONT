from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any


CATALOG_PATH = (
    Path(__file__).resolve().parents[3]
    / "config"
    / "product_catalog.json"
)


def calculate_supply_amount(
    quantity: float | int | None,
    unit_price: float | int | None,
) -> int | None:
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


def _normalize_product(value: Any) -> str:
    return re.sub(
        r"[^0-9A-Za-z가-힣]",
        "",
        str(value or ""),
    ).casefold()


@lru_cache(maxsize=1)
def _catalog_products() -> dict[str, dict[str, Any]]:
    try:
        data = json.loads(
            CATALOG_PATH.read_text(encoding="utf-8")
        )
    except (OSError, ValueError, TypeError):
        return {}

    result: dict[str, dict[str, Any]] = {}

    for category in data.get("categories", []):
        for product in category.get("products", []):
            names = [
                product.get("name"),
                *(product.get("aliases") or []),
            ]
            for name in names:
                key = _normalize_product(name)
                if key:
                    result[key] = product

    return result


def _present(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (list, tuple, set)):
        return any(_present(part) for part in value)
    if isinstance(value, str):
        return bool(value.strip())
    return True


def _catalog_field_value(item: Any, field: dict[str, Any]) -> Any:
    legacy = field.get("legacy_field")

    if legacy:
        return getattr(item, str(legacy), None)

    attributes = getattr(item, "spec_attributes", None) or {}

    if isinstance(attributes, dict):
        return attributes.get(str(field.get("key") or ""))

    return None


def validate_quote_items(items: list[Any]) -> list[str]:
    if not items:
        return ["견적서에 입력할 품목이 없습니다."]

    errors: list[str] = []
    catalog = _catalog_products()

    for index, item in enumerate(items, start=1):
        product_name = str(
            getattr(item, "product_name", "") or ""
        ).strip()

        if not product_name:
            errors.append(f"{index}번째 품목명이 비어 있습니다.")
            continue

        quantity = getattr(item, "quantity", None)
        try:
            quantity_number = float(quantity)
        except (TypeError, ValueError):
            quantity_number = 0.0

        if quantity is None or quantity_number <= 0:
            errors.append(f"{index}번째 품목의 수량을 입력해주세요.")

        unit = str(getattr(item, "unit", "") or "").strip()
        if not unit:
            errors.append(f"{index}번째 품목의 단위를 입력해주세요.")

        unit_price = getattr(item, "unit_price", None)
        try:
            unit_price_number = float(unit_price)
        except (TypeError, ValueError):
            unit_price_number = 0.0

        if unit_price is None or unit_price_number <= 0:
            errors.append(f"{index}번째 품목의 확정 단가를 입력해주세요.")

        product = (
            catalog.get(
                _normalize_product(
                    getattr(item, "normalized_product", None)
                )
            )
            or catalog.get(_normalize_product(product_name))
        )

        if product:
            for field in product.get("fields", []):
                value = _catalog_field_value(item, field)

                if _present(value):
                    continue

                if (
                    field.get("legacy_field") == "quantity"
                    and (
                        quantity is None
                        or quantity_number <= 0
                    )
                ):
                    continue

                label = str(
                    field.get("label")
                    or field.get("key")
                    or "사양"
                )
                errors.append(
                    f"{index}번째 품목의 {label}을(를) 입력해주세요."
                )
        else:
            specification = str(
                getattr(item, "specification", "") or ""
            ).strip()
            if not specification:
                errors.append(
                    f"{index}번째 품목의 규격/사양을 입력해주세요."
                )

    return errors
