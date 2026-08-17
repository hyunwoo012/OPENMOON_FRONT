"""첨부파일 읽기 + LLM 분석 정확도를 먼저 확인하기 위한 임시 테스트 스크립트.

에이전트 Tool이나 견적 분석 파이프라인에 통합하기 전에, "첨부파일을 실제로 잘
읽어서 정확하게 판단하는가"만 독립적으로 검증한다. 결과는 화면이 아니라
txt 파일로 떨어뜨려서 사람이 눈으로 비교하기 쉽게 한다.

사용법:
    .venv/bin/python -m backend.scripts.test_attachment_analysis <mail_id> [<mail_id> ...]

mail_id를 안 주면 텍스트 추출/이미지/스캔 PDF가 섞인 표본 메일 몇 개로 실행한다.
결과는 backend/data/attachment_analysis_test/mail_<id>.txt 에 저장된다.
"""

from __future__ import annotations

import sys
from pathlib import Path

from openai import OpenAI
from sqlalchemy.orm import Session

from backend.app.config import get_settings
from backend.app.database import engine
from backend.app.enums import AttachmentStatus
from backend.app.models import Mail
from backend.app.services.attachment_service import extract_hwp_preview_text, extract_pptx_text
from backend.app.services.llm_service import _image_data_url, _pdf_page_data_urls

OUTPUT_DIR = Path("backend/data/attachment_analysis_test")
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
MAX_IMAGES = 4

PROMPT = (
    "너는 인쇄·디자인 회사 '(주)열린문디자인'의 견적 담당자야. 아래 메일 제목/본문과 "
    "첨부파일 내용을 보고 이 주문의 품목/규격/수량/재질/용도를 파악해줘. "
    "첨부파일에서 실제로 읽은 내용만 근거로 삼고, 확실하지 않은 값은 추측하지 말고 "
    "'확인 필요'라고 표시해. 가능하면 품목별로 표(품목/규격/수량/특이사항) 형태로 정리해줘."
)


def build_content(mail: Mail) -> tuple[list[dict], list[str]]:
    """LLM에 보낼 content 블록과, 사람이 확인할 첨부파일별 처리 로그를 만든다."""
    text_parts = [
        f"[메일 제목] {mail.original_subject or ''}",
        f"[메일 본문]\n{(mail.original_body or '')[:3000]}",
    ]
    log: list[str] = []
    content: list[dict] = []
    image_count = 0

    for attachment in mail.attachments:
        path = Path(attachment.saved_path)
        if attachment.extracted_text and attachment.extracted_text.strip():
            text_parts.append(
                f"\n[첨부파일: {attachment.filename} - 추출된 텍스트]\n{attachment.extracted_text[:4000]}"
            )
            log.append(f"- {attachment.filename}: 추출된 텍스트 사용 ({len(attachment.extracted_text)}자)")
            continue

        if not path.exists():
            log.append(f"- {attachment.filename}: 파일 없음 (saved_path={path})")
            continue

        suffix = path.suffix.lower()

        # DB에 아직 추출 결과가 없는 pptx/hwp는(과거에 import된 첨부라 예전 코드로
        # 처리됐을 수 있으니) 최신 추출 함수로 그 자리에서 다시 뽑아본다.
        if suffix == ".pptx":
            live_text = extract_pptx_text(path)
            if live_text.strip():
                text_parts.append(f"\n[첨부파일: {attachment.filename} - 추출된 텍스트]\n{live_text[:4000]}")
                log.append(f"- {attachment.filename}: pptx 실시간 추출 사용 ({len(live_text)}자)")
                continue
        elif suffix == ".hwp":
            live_text = extract_hwp_preview_text(path)
            if live_text.strip():
                text_parts.append(
                    f"\n[첨부파일: {attachment.filename} - HWP 미리보기 추출(참고용, 불완전할 수 있음)]\n{live_text[:4000]}"
                )
                log.append(f"- {attachment.filename}: hwp PrvText 실시간 추출 사용 ({len(live_text)}자)")
                continue

        if suffix in IMAGE_EXTENSIONS and image_count < MAX_IMAGES:
            content.append({"type": "input_image", "image_url": _image_data_url(path)})
            text_parts.append(f"\n[첨부파일: {attachment.filename}] 이미지로 첨부됨 (아래 참고)")
            log.append(f"- {attachment.filename}: 이미지로 Vision 전달")
            image_count += 1
        elif suffix == ".pdf" and attachment.status == AttachmentStatus.IMAGE_PENDING and image_count < MAX_IMAGES:
            urls = _pdf_page_data_urls(path, max_pages=2)
            for url in urls:
                if image_count >= MAX_IMAGES:
                    break
                content.append({"type": "input_image", "image_url": url})
                image_count += 1
            text_parts.append(f"\n[첨부파일: {attachment.filename}] 스캔 PDF 앞 {len(urls)}페이지를 이미지로 첨부됨")
            log.append(f"- {attachment.filename}: 스캔 PDF {len(urls)}페이지를 Vision 전달")
        else:
            log.append(f"- {attachment.filename}: 텍스트/이미지 처리 대상 아님 (status={attachment.status})")

    content.insert(0, {"type": "input_text", "text": PROMPT + "\n\n" + "\n".join(text_parts)})
    return content, log


def main(mail_ids: list[int]) -> None:
    settings = get_settings()
    if not settings.openai_api_key:
        raise SystemExit("OPENAI_API_KEY가 .env에 없습니다.")

    client = OpenAI(api_key=settings.openai_api_key)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with Session(engine) as session:
        for mail_id in mail_ids:
            mail = session.get(Mail, mail_id)
            if mail is None:
                print(f"mail_id={mail_id} 없음, 건너뜀")
                continue

            content, log = build_content(mail)
            response = client.responses.create(
                model=settings.openai_model,
                input=[{"role": "user", "content": content}],
            )

            out_path = OUTPUT_DIR / f"mail_{mail_id}.txt"
            out_path.write_text(
                f"mail_id={mail_id}\n"
                f"제목: {mail.original_subject}\n\n"
                "--- 첨부파일 처리 로그 ---\n"
                + "\n".join(log)
                + "\n\n--- LLM 분석 결과 ---\n"
                + (response.output_text or "(빈 응답)"),
                encoding="utf-8",
            )
            print(f"mail_id={mail_id} -> {out_path}")


if __name__ == "__main__":
    ids = [int(value) for value in sys.argv[1:]] or [7, 16, 25]
    main(ids)
