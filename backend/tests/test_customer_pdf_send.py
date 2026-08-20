from __future__ import annotations

from email.message import EmailMessage
from pathlib import Path
from types import SimpleNamespace

import pytest

from backend.app.services import smtp_service


def _draft(tmp_path: Path):
    mail = SimpleNamespace(
        customer_email="customer@example.com",
        original_sender_email="sender@example.com",
        original_subject="원본 견적 요청",
        outer_subject=None,
        forward_depth=1,
        message_id=None,
        references=None,
        status=None,
    )

    return SimpleNamespace(
        id=17,
        mail=mail,
        email_subject="직접 수정한 발송 제목",
        file_path=str(tmp_path / "internal.xlsx"),
        status=smtp_service.DraftStatus.APPROVED,
    )


def _settings():
    return SimpleNamespace(
        allow_live_send=True,
        daum_login_id="sender@example.com",
        daum_app_password="secret",
        smtp_server="smtp.example.com",
        smtp_port=465,
    )


def test_validate_send_ready_requires_customer_pdf(tmp_path, monkeypatch):
    draft = _draft(tmp_path)
    pdf = tmp_path / "customer.pdf"

    monkeypatch.setattr(
        smtp_service,
        "customer_pdf_path",
        lambda settings, row: pdf,
    )

    with pytest.raises(
        FileNotFoundError,
        match="고객용 PDF",
    ):
        smtp_service.validate_send_ready(
            _settings(),
            draft,
        )


def test_validate_send_ready_returns_pdf_only(tmp_path, monkeypatch):
    draft = _draft(tmp_path)
    internal = Path(draft.file_path)
    internal.write_bytes(b"INTERNAL XLSX")

    pdf = tmp_path / "customer.pdf"
    pdf.write_bytes(b"%PDF-test")

    monkeypatch.setattr(
        smtp_service,
        "customer_pdf_path",
        lambda settings, row: pdf,
    )

    recipient, attachment = (
        smtp_service.validate_send_ready(
            _settings(),
            draft,
        )
    )

    assert recipient == "customer@example.com"
    assert attachment == pdf
    assert attachment != internal
    assert attachment.suffix == ".pdf"


def test_saved_subject_is_preferred_in_source():
    source = Path(
        smtp_service.__file__
    ).read_text(
        encoding="utf-8"
    )

    assert "saved_subject = (" in source
    assert 'message["Subject"] = saved_subject' in source
    assert "draft.email_subject" in source
