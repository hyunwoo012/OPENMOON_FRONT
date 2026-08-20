from __future__ import annotations

from pathlib import Path

from PIL import Image

from backend.app.enums import AttachmentStatus
from backend.app.models import Attachment
from backend.app.services.attachment_service import process_attachment
from backend.app.services.llm_service import (
    _chat_completion_user_content,
    _image_data_url,
)


class _FakeSession:
    def add(self, _value) -> None:
        pass

    def flush(self) -> None:
        pass


def _make_image(path: Path, image_format: str) -> None:
    image = Image.new("RGB", (16, 16), "white")
    image.save(path, format=image_format)


def test_png_attachment_is_marked_for_vision(tmp_path: Path):
    path = tmp_path / "sample.png"
    _make_image(path, "PNG")

    attachment = Attachment(
        mail_id=1,
        filename="sample.png",
        saved_path=str(path),
        size_bytes=path.stat().st_size,
    )

    process_attachment(_FakeSession(), attachment)

    assert attachment.status == AttachmentStatus.IMAGE_PENDING
    assert _image_data_url(path).startswith(
        "data:image/png;base64,"
    )


def test_jpg_attachment_is_marked_for_vision(tmp_path: Path):
    path = tmp_path / "sample.jpg"
    _make_image(path, "JPEG")

    attachment = Attachment(
        mail_id=1,
        filename="sample.jpg",
        saved_path=str(path),
        size_bytes=path.stat().st_size,
    )

    process_attachment(_FakeSession(), attachment)

    assert attachment.status == AttachmentStatus.IMAGE_PENDING
    assert _image_data_url(path).startswith(
        "data:image/jpeg;base64,"
    )


def test_chat_completion_fallback_keeps_images():
    content = _chat_completion_user_content(
        "메일 분석",
        [
            "data:image/png;base64,AAA",
            "data:image/jpeg;base64,BBB",
        ],
    )

    assert content[0] == {
        "type": "text",
        "text": "메일 분석",
    }

    assert content[1]["type"] == "image_url"
    assert (
        content[1]["image_url"]["url"]
        == "data:image/png;base64,AAA"
    )

    assert content[2]["type"] == "image_url"
    assert (
        content[2]["image_url"]["url"]
        == "data:image/jpeg;base64,BBB"
    )
