from __future__ import annotations

from backend.app.services.quotation_service import (
    CUSTOMER_PDF_EXPORT_SCRIPT,
    CUSTOMER_PRIVATE_CELLS,
)


def test_customer_private_cells_are_removed():
    assert CUSTOMER_PRIVATE_CELLS == (
        "L5",
        "L6",
        "L7",
    )
    for cell in CUSTOMER_PRIVATE_CELLS:
        assert f'"{cell}"' in CUSTOMER_PDF_EXPORT_SCRIPT


def test_customer_pdf_clears_internal_column():
    assert "Columns.Item(20).ClearContents()" in CUSTOMER_PDF_EXPORT_SCRIPT
    assert "Columns.Item(20).Hidden = $true" in CUSTOMER_PDF_EXPORT_SCRIPT


def test_customer_pdf_exports_single_sheet():
    assert "ExportAsFixedFormat" in CUSTOMER_PDF_EXPORT_SCRIPT
    assert "$sheet.ExportAsFixedFormat" in CUSTOMER_PDF_EXPORT_SCRIPT
