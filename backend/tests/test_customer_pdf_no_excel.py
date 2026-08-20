from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import fitz

from backend.app.services.quotation_service import (
    _export_customer_pdf_python,
)


def test_python_customer_pdf_uses_official_template(tmp_path: Path):
    settings = SimpleNamespace(
        generated_quotes_dir=tmp_path,
        default_delivery_place="지정장소",
        default_payment_terms="현금 또는 카드결제",
        default_validity="견적일로부터",
    )

    draft = SimpleNamespace(
        id=77,
        customer_name="영인면행정복지센터",
    )

    item = SimpleNamespace(
        product_name="현수막",
        normalized_product="현수막",
        specification="3000*600mm",
        width_mm=None,
        height_mm=None,
        quantity=4,
        unit="장",
        paper=None,
        print_sides=None,
        material=None,
        spec_attributes={},
        unit_price=33000,
        amount=132000,
        cost_price=999999,
    )

    mail = SimpleNamespace(
        id=99,
        customer_organization="영인면행정복지센터",
        customer_department=None,
        customer_name="테스트담당자",
        customer_phone="010-1111-2222",
        customer_email="customer@example.com",
        original_sender_email="customer@example.com",
        delivery_place="지정장소",
        payment_terms="현금 또는 카드결제",
        items=[item],
    )

    pdf_path = _export_customer_pdf_python(
        settings,
        draft,
        tmp_path / "unused.xlsx",
        mail,
    )

    assert pdf_path.exists()
    assert pdf_path.read_bytes()[:5] == b"%PDF-"

    with fitz.open(pdf_path) as document:
        assert document.page_count == 1
        text = "\n".join(
            page.get_text()
            for page in document
        )

    # 공식 Excel 견적서 PDF의 고정 문구가 살아 있어야 한다.
    assert "(주)열린문디자인" in text
    assert "등록번호" in text
    assert "아래와 같이 견적합니다." in text
    assert "공급 금액" in text

    # 동적 견적 정보
    assert "영인면행정복지센터" in text
    assert "현수막" in text
    assert "33,000" in text
    assert "132,000" in text

    # 기존 샘플값 / 고객 개인정보 / 내부 원가는 노출 금지.
    assert "한솔산업" not in text
    assert "010-1111-2222" not in text
    assert "customer@example.com" not in text
    assert "999999" not in text
