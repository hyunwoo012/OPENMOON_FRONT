from __future__ import annotations

import os
from pathlib import Path
from types import SimpleNamespace

import pytest

from backend.app.services.quotation_service import (
    _canonical_customer_product_name,
)
from backend.app.services.smtp_service import (
    _validate_customer_pdf_attachment,
)


def test_customer_product_uses_normalized_name_first():
    assert (
        _canonical_customer_product_name(
            "친환경 현수막",
            "현수막",
            ["현수막", "명함"],
        )
        == "현수막"
    )


def test_customer_product_collapses_modifier_to_catalog_name():
    assert (
        _canonical_customer_product_name(
            "친환경 현수막",
            None,
            ["현수막", "명함"],
        )
        == "현수막"
    )


def test_customer_product_keeps_exact_catalog_name():
    assert (
        _canonical_customer_product_name(
            "차량용스티커",
            None,
            [
                "스티커",
                "차량용스티커",
            ],
        )
        == "차량용스티커"
    )


def test_customer_pdf_requires_real_pdf_header(tmp_path: Path):
    fake = tmp_path / "quote.pdf"
    fake.write_bytes(b"NOT-PDF")

    with pytest.raises(
        ValueError,
        match="형식이 올바르지",
    ):
        _validate_customer_pdf_attachment(
            fake
        )


def test_customer_pdf_rejects_stale_file(tmp_path: Path):
    xlsx = tmp_path / "internal.xlsx"
    pdf = tmp_path / "customer.pdf"

    pdf.write_bytes(
        b"%PDF-test"
    )
    xlsx.write_bytes(
        b"XLSX"
    )

    os.utime(
        pdf,
        (100, 100),
    )
    os.utime(
        xlsx,
        (200, 200),
    )

    with pytest.raises(
        ValueError,
        match="오래된 버전",
    ):
        _validate_customer_pdf_attachment(
            pdf,
            xlsx,
        )


def test_customer_pdf_accepts_current_valid_pdf(tmp_path: Path):
    xlsx = tmp_path / "internal.xlsx"
    pdf = tmp_path / "customer.pdf"

    xlsx.write_bytes(
        b"XLSX"
    )
    pdf.write_bytes(
        b"%PDF-test"
    )

    os.utime(
        xlsx,
        (100, 100),
    )
    os.utime(
        pdf,
        (200, 200),
    )

    _validate_customer_pdf_attachment(
        pdf,
        xlsx,
    )
