from __future__ import annotations

import imaplib
import mimetypes
import re
import smtplib
from datetime import datetime
from email.message import EmailMessage
from email.policy import SMTP
from email.utils import format_datetime, make_msgid
from html import escape
from pathlib import Path

from sqlalchemy.orm import Session

from ..config import Settings
from ..enums import DraftStatus, MailStatus
from ..models import QuotationDraft
from .excel_open_service import resolve_quotation_file


SIGNATURE_ROOT = Path(__file__).resolve().parents[2] / "assets" / "email_signatures"
EMPLOYEES = {
    "moon_jeongseon": {
        "job": "업무총괄", "name": "문정선", "rank": "대표이사",
        "phone": "010-4420-5106", "image": "moon_jeongseon.png",
    },
    "shin_woohyun": {
        "job": "관리부서", "name": "신우현", "rank": "주임",
        "phone": "041-548-5106", "image": "shin_woohyun.png",
    },
    "kwon_jihye": {
        "job": "회계담당", "name": "권지혜", "rank": "대리",
        "phone": "070-8667-4730", "image": "kwon_jihye.png",
    },
    "kim_heejung": {
        "job": "관리부", "name": "김희정", "rank": "과장",
        "phone": "041-548-5106", "image": "kim_heejung.png",
    },
}


def _sender_address(login_id: str) -> str:
    """Daum accepts an account ID for login, but From must be a full address."""
    value = login_id.strip()
    return value if "@" in value else f"{value}@daum.net"


def _sent_mailboxes(imap: imaplib.IMAP4_SSL) -> list[str]:
    """Return the server's special-use Sent folder, followed by safe fallbacks."""
    candidates: list[str] = []
    status, rows = imap.list()
    if status == "OK":
        for raw in rows or []:
            if not raw:
                continue
            line = raw.decode("ascii", errors="replace")
            match = re.match(r"\((?P<flags>[^)]*)\)\s+\"[^\"]*\"\s+(?P<name>.+)$", line)
            if not match or "\\Sent" not in match.group("flags").split():
                continue
            name = match.group("name").strip()
            if name.startswith('"') and name.endswith('"'):
                name = name[1:-1].replace(r'\"', '"').replace(r"\\", "\\")
            candidates.append(name)
    for fallback in ("Sent", "Sent Messages", "INBOX.Sent"):
        if fallback not in candidates:
            candidates.append(fallback)
    return candidates


def _append_to_sent(settings: Settings, raw_message: bytes) -> None:
    imap = imaplib.IMAP4_SSL(settings.imap_server, settings.imap_port, timeout=120)
    try:
        imap.login(settings.daum_login_id, settings.daum_app_password)
        errors: list[str] = []
        for mailbox in _sent_mailboxes(imap):
            try:
                # Daum은 APPEND의 선택 인자(플래그·내부 날짜)를 거부하는 경우가 있어
                # 서버가 자체 처리하도록 최소 인자 형식으로 저장한다.
                status, response = imap.append(mailbox, None, None, raw_message)
                if status == "OK":
                    return
                errors.append(f"{mailbox}: {status} {response!r}")
            except imaplib.IMAP4.error as error:
                # 한 후보 폴더가 거부되어도 나머지 Sent 폴더 후보를 계속 확인한다.
                errors.append(f"{mailbox}: {error}")
        raise RuntimeError("보낸메일함을 찾거나 저장할 수 없습니다. " + " / ".join(errors))
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def _email_content(employee_key: str) -> tuple[str, str, Path]:
    employee = EMPLOYEES.get(employee_key)
    if not employee:
        raise ValueError("발송 담당 직원을 선택해주세요.")
    identity = f"{employee['job']} {employee['name']} {employee['rank']}"
    text = f"""안녕하세요.
(주)열린문디자인 {identity}입니다.

♥주문 주셔서 진심으로 감사드립니다.♥

요청주신 건에 대한 견적서를 첨부하여 보내드립니다.
견적서 검토 후 제작 진행 여부를 알려주시면, 담당 디자이너를 배정하여 시안을 받아보실 수 있도록 신속히 진행하겠습니다.

견적 또는 제작 관련하여 문의사항이 있으시면 아래 연락처로 편하게 연락 주시기 바랍니다.

또한, 저희 회사는 사회적기업 및 여성기업 확인서를 보유하고 있으니 관련 서류가 필요하실 경우 요청해 주시면 메일로 송부드리겠습니다.

무더워지는 날씨에 건강 유의하시고, 시원하고 기분 좋은 하루 보내시길 바랍니다.

감사합니다.

(주)열린문디자인 {identity}
☎ {employee['phone']}"""
    html = "".join(
        f"<p style=\"margin:0 0 16px;line-height:1.65\">{escape(block).replace(chr(10), '<br>')}</p>"
        for block in text.split("\n\n")
    )
    html = (
        '<div style="font-family:Arial,\'Malgun Gothic\',sans-serif;font-size:14px;color:#111">'
        f"{html}<p style=\"margin-top:18px\"><img src=\"cid:employee-signature\" "
        'alt="직원 안내 이미지" style="display:block;max-width:100%;height:auto"></p></div>'
    )
    return text, html, SIGNATURE_ROOT / str(employee["image"])


def validate_send_ready(settings: Settings, draft: QuotationDraft) -> tuple[str, Path]:
    """Validate all local prerequisites before approval changes the draft state."""
    if not settings.allow_live_send and not settings.approval_test_mode:
        raise PermissionError("ALLOW_LIVE_SEND=false입니다. 실제 발송 전 테스트를 완료하세요.")
    if not settings.daum_login_id or not settings.daum_app_password:
        raise RuntimeError("메일 계정 정보가 설정되지 않았습니다.")

    recipient = (
        settings.approval_test_recipient.strip()
        if settings.approval_test_mode
        else (draft.mail.customer_email or draft.mail.original_sender_email)
    )
    if not recipient:
        raise ValueError("테스트 수신자 또는 고객 이메일 주소가 없습니다.")
    attachment_path = Path(draft.file_path)
    if not attachment_path.exists():
        attachment_path = resolve_quotation_file(
            attachment_path,
            settings.quotation_files_path,
        )
    return recipient, attachment_path


def send_draft(
    session: Session,
    settings: Settings,
    draft: QuotationDraft,
    employee_key: str = "kim_heejung",
) -> QuotationDraft:
    if draft.status != DraftStatus.APPROVED:
        raise ValueError("승인된 견적서만 발송할 수 있습니다.")
    recipient, attachment_path = validate_send_ready(settings, draft)

    message = EmailMessage()
    message["From"] = _sender_address(settings.daum_login_id)
    message["To"] = recipient
    message["Date"] = format_datetime(datetime.now().astimezone())
    message["Message-ID"] = make_msgid(domain=_sender_address(settings.daum_login_id).split("@", 1)[1])
    text_body, html_body, signature_path = _email_content(employee_key)
    if not signature_path.exists():
        raise FileNotFoundError(f"직원 서명 이미지를 찾을 수 없습니다: {signature_path}")
    original_subject = (draft.mail.original_subject or draft.mail.outer_subject or "견적 문의").strip()
    message["Subject"] = original_subject if original_subject.lower().startswith("re:") else f"Re: {original_subject}"
    message.set_content(text_body)
    message.add_alternative(html_body, subtype="html")
    html_part = message.get_payload()[-1]
    html_part.add_related(
        signature_path.read_bytes(),
        maintype="image",
        subtype="png",
        cid="<employee-signature>",
        filename=signature_path.name,
        disposition="inline",
    )

    # 전달 메일의 Message-ID는 전달자와의 바깥 스레드이므로 직접 수신 메일에만 연결한다.
    if draft.mail.forward_depth == 0 and draft.mail.message_id:
        message["In-Reply-To"] = draft.mail.message_id
        references = (draft.mail.references or "").strip()
        message["References"] = f"{references} {draft.mail.message_id}".strip()

    content_type, _ = mimetypes.guess_type(attachment_path.name)
    maintype, subtype = (content_type or "application/octet-stream").split("/", 1)
    message.add_attachment(
        attachment_path.read_bytes(),
        maintype=maintype,
        subtype=subtype,
        filename=attachment_path.name,
    )

    try:
        # 견적서와 고해상도 서명 이미지를 함께 전송하므로 느린 회선도 허용한다.
        with smtplib.SMTP_SSL(settings.smtp_server, settings.smtp_port, timeout=120) as smtp:
            smtp.login(settings.daum_login_id, settings.daum_app_password)
            smtp.send_message(message)
        draft.status = DraftStatus.SENT
        draft.sent_at = datetime.now().astimezone().replace(tzinfo=None)
        draft.sent_to = recipient
        draft.mail.status = MailStatus.SENT
        draft.error_message = None
        try:
            _append_to_sent(settings, message.as_bytes(policy=SMTP))
        except Exception as append_error:
            # SMTP 발송은 이미 성공했으므로 재발송으로 인한 중복 메일을 막는다.
            draft.error_message = f"메일은 발송됐지만 보낸메일함 저장에 실패했습니다: {append_error}"
    except Exception as error:
        draft.status = DraftStatus.FAILED
        draft.error_message = f"{type(error).__name__}: {error}"
        raise
    finally:
        session.commit()
    return draft
