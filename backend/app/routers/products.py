from __future__ import annotations

import json

from fastapi import APIRouter, HTTPException

from ..config import PROJECT_ROOT


router = APIRouter(prefix="/api/products", tags=["products"])
CATALOG_PATH = PROJECT_ROOT / "config" / "product_catalog.json"


@router.get("/catalog")
def get_product_catalog():
    if not CATALOG_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f"품목 카탈로그 파일을 찾을 수 없습니다: {CATALOG_PATH}",
        )

    try:
        with CATALOG_PATH.open("r", encoding="utf-8") as handle:
            catalog = json.load(handle)
    except (OSError, json.JSONDecodeError) as error:
        raise HTTPException(
            status_code=500,
            detail=f"품목 카탈로그를 읽지 못했습니다: {error}",
        ) from error

    categories = catalog.get("categories")
    if not isinstance(categories, list):
        raise HTTPException(
            status_code=500,
            detail="품목 카탈로그의 categories 형식이 올바르지 않습니다.",
        )

    return catalog
