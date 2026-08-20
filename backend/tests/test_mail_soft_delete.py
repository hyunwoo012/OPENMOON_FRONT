from __future__ import annotations

from fastapi import FastAPI
from fastapi.testclient import TestClient
from sqlalchemy import create_engine, select
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import StaticPool

from backend.app.database import Base, get_db
from backend.app.models import Mail
from backend.app.routers import mails


def _setup():
    engine = create_engine(
        "sqlite://",
        connect_args={
            "check_same_thread": False
        },
        poolclass=StaticPool,
    )

    SessionLocal = sessionmaker(
        bind=engine,
        autoflush=False,
        autocommit=False,
        expire_on_commit=False,
    )

    Base.metadata.create_all(
        bind=engine
    )

    with SessionLocal() as session:
        first = Mail(
            account="test@example.com",
            uid="100",
            message_id="<mail-100@example.com>",
            outer_subject="삭제 대상",
            original_subject="삭제 대상",
            status="NEW",
        )

        second = Mail(
            account="test@example.com",
            uid="101",
            message_id="<mail-101@example.com>",
            outer_subject="유지 대상",
            original_subject="유지 대상",
            status="NEW",
        )

        session.add_all(
            [first, second]
        )
        session.commit()

        deleted_id = first.id
        kept_id = second.id

    app = FastAPI()
    app.include_router(
        mails.router
    )

    def override_db():
        session = SessionLocal()

        try:
            yield session
        finally:
            session.close()

    app.dependency_overrides[
        get_db
    ] = override_db

    return (
        TestClient(app),
        SessionLocal,
        deleted_id,
        kept_id,
    )


def test_delete_mail_is_soft_delete():
    client, SessionLocal, deleted_id, _ = _setup()

    response = client.delete(
        f"/api/mails/{deleted_id}"
    )

    assert response.status_code == 200
    assert response.json()["mode"] == "soft"
    assert response.json()["imap_deleted"] is False

    with SessionLocal() as session:
        row = session.get(
            Mail,
            deleted_id,
        )

        # 행 자체는 tombstone으로 남는다.
        assert row is not None
        assert row.deleted_at is not None
        assert row.uid == "100"


def test_deleted_mail_is_hidden_from_list():
    client, _, deleted_id, kept_id = _setup()

    assert client.delete(
        f"/api/mails/{deleted_id}"
    ).status_code == 200

    response = client.get(
        "/api/mails"
    )

    assert response.status_code == 200

    ids = {
        row["id"]
        for row in response.json()
    }

    assert deleted_id not in ids
    assert kept_id in ids


def test_deleted_mail_detail_returns_404():
    client, _, deleted_id, _ = _setup()

    assert client.delete(
        f"/api/mails/{deleted_id}"
    ).status_code == 200

    response = client.get(
        f"/api/mails/{deleted_id}"
    )

    assert response.status_code == 404


def test_deleted_uid_remains_as_sync_tombstone():
    client, SessionLocal, deleted_id, _ = _setup()

    assert client.delete(
        f"/api/mails/{deleted_id}"
    ).status_code == 200

    with SessionLocal() as session:
        existing = session.scalar(
            select(Mail.id).where(
                Mail.account == "test@example.com",
                Mail.uid == "100",
            )
        )

        # sync_imap의 기존 UID 확인에 계속 걸리므로 재수집되지 않는다.
        assert existing == deleted_id
