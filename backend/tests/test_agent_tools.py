from __future__ import annotations

from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from backend.app.config import Settings
from backend.app.database import Base
from backend.app.models import Attachment, Mail, MailItem
from backend.app.services import agent_tools
from backend.app.services.agent_tools import AgentToolContext, execute_agent_tool


def _context(tmp_path: Path) -> tuple[Session, AgentToolContext]:
    engine = create_engine(f"sqlite:///{tmp_path / 'agent.db'}")
    Base.metadata.create_all(engine)
    session = Session(engine, expire_on_commit=False)
    mail = Mail(
        uid="agent-mail",
        customer_organization="테스트기관",
        customer_name="홍길동",
        status="READY_FOR_QUOTE",
        analysis_payload={"is_order_related": True},
    )
    mail.items = [
        MailItem(
            position=1,
            product_name="기존 품목",
            specification="기존 규격",
            width_mm=100,
            height_mm=100,
            quantity=1,
            unit="개",
            paper="해당 없음",
            print_sides="해당 없음",
            material="기존 재질",
            unit_price=1000,
            amount=1000,
            confirmed=True,
        )
    ]
    session.add(mail)
    session.commit()
    settings = Settings(
        database_url=f"sqlite:///{tmp_path / 'agent.db'}",
        quotation_files_path=tmp_path / "quotes",
        use_external_price_engine=False,
    )
    return session, AgentToolContext(session=session, settings=settings, mail=mail)


def _item(product: str, price: int) -> dict[str, object]:
    return {
        "product_name": product,
        "specification": "5000*600mm",
        "width_mm": 5000,
        "height_mm": 600,
        "quantity": 2,
        "unit": "장",
        "paper": "해당 없음",
        "print_sides": "단면",
        "material": "현수막",
        "unit_price": price,
        "schedule_note": None,
    }


def test_agent_replaces_natural_language_quote_items(tmp_path: Path):
    session, ctx = _context(tmp_path)
    try:
        result, _ = execute_agent_tool(
            ctx,
            "replace_quote_items",
            {"items": [_item("현수막", 30000), _item("배너", 20000)]},
        )
        assert result["replaced"] is True
        assert [item.product_name for item in ctx.mail.items] == ["현수막", "배너"]
        assert [item.position for item in ctx.mail.items] == [1, 2]
        assert ctx.mail.items[0].amount == 60000
    finally:
        session.close()


def test_agent_quote_creation_requests_storage_modal(tmp_path: Path):
    session, ctx = _context(tmp_path)
    try:
        result, _ = execute_agent_tool(ctx, "create_quotation_draft", {})
        assert result["requires_storage_selection"] is True
        assert ctx.actions[-1]["type"] == "open_storage_modal"
    finally:
        session.close()


def test_agent_reads_extracted_attachment_text(tmp_path: Path):
    session, ctx = _context(tmp_path)
    try:
        attachment = Attachment(
            mail_id=ctx.mail.id,
            filename="견적요청서.pdf",
            saved_path=str(tmp_path / "견적요청서.pdf"),
            status="EXTRACTED",
            extracted_text="가로 500mm 세로 700mm 현수막 2장 방수천 재질 요청",
        )
        session.add(attachment)
        session.commit()

        result, evidence = execute_agent_tool(ctx, "read_mail_attachments", {"filename": None})

        assert result["count"] == 1
        assert "방수천" in result["attachments"][0]["text"]
        assert evidence[0]["type"] == "attachment"
    finally:
        session.close()


def test_agent_analyzes_image_attachment_lazily(tmp_path: Path, monkeypatch):
    session, ctx = _context(tmp_path)
    try:
        image_path = tmp_path / "시안.png"
        image_path.write_bytes(b"fake-bytes")
        attachment = Attachment(
            mail_id=ctx.mail.id,
            filename="시안.png",
            saved_path=str(image_path),
            status="IMAGE_PENDING",
        )
        session.add(attachment)
        session.commit()

        monkeypatch.setattr(
            agent_tools,
            "_analyze_attachment_image",
            lambda _ctx, _path: "현수막 시안: 2026년으로 연도 변경 요청",
        )

        result, _ = execute_agent_tool(ctx, "read_mail_attachments", {"filename": "시안"})

        assert result["count"] == 1
        assert "2026년" in result["attachments"][0]["text"]
        assert attachment.status == "EXTRACTED"
    finally:
        session.close()


def test_agent_opens_only_registered_history_source(tmp_path: Path, monkeypatch):
    session, ctx = _context(tmp_path)
    try:
        monkeypatch.setattr(agent_tools, "is_known_external_history_source", lambda *_args: True)
        monkeypatch.setattr(
            agent_tools,
            "open_excel_location",
            lambda path, sheet=None: {"opened": True, "file_path": path, "sheet": sheet},
        )
        result, _ = execute_agent_tool(
            ctx,
            "open_quotation_source",
            {"source_file": "C:/quotes/history.xlsx", "source_sheet": "0815"},
        )
        assert result["opened"] is True
        assert result["sheet"] == "0815"
    finally:
        session.close()
