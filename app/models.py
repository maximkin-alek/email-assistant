from __future__ import annotations

import datetime as dt
from typing import Literal

from sqlalchemy import Boolean, DateTime, Integer, LargeBinary, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column

from app.db import Base


Provider = Literal["gmail", "imap"]


class Mailbox(Base):
    __tablename__ = "mailboxes"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    provider: Mapped[str] = mapped_column(String(32), nullable=False)  # gmail|imap
    name: Mapped[str] = mapped_column(String(128), nullable=False)  # отображаемое имя

    # Для IMAP (итерация 2): host/user/pass хранить в зашифрованном виде.
    imap_host_enc: Mapped[str | None] = mapped_column(Text, nullable=True)
    imap_user_enc: Mapped[str | None] = mapped_column(Text, nullable=True)
    imap_password_enc: Mapped[str | None] = mapped_column(Text, nullable=True)
    imap_port: Mapped[int | None] = mapped_column(Integer, nullable=True)
    imap_folder: Mapped[str | None] = mapped_column(String(128), nullable=True)
    imap_last_uid: Mapped[int | None] = mapped_column(Integer, nullable=True)
    imap_tls_verify: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)

    last_sync_at: Mapped[dt.datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    last_sync_status: Mapped[str | None] = mapped_column(String(32), nullable=True)  # ok|error|running
    last_sync_count: Mapped[int | None] = mapped_column(Integer, nullable=True)
    last_sync_error: Mapped[str | None] = mapped_column(Text, nullable=True)

    gmail_email: Mapped[str | None] = mapped_column(String(256), nullable=True)
    gmail_credentials_enc: Mapped[str | None] = mapped_column(Text, nullable=True)
    gmail_last_history_id: Mapped[str | None] = mapped_column(String(128), nullable=True)

    is_enabled: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=lambda: dt.datetime.now(dt.UTC))


class EmailMessage(Base):
    __tablename__ = "email_messages"
    __table_args__ = (
        UniqueConstraint("mailbox_id", "provider_message_id", name="uq_mailbox_provider_message_id"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    mailbox_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)

    provider_message_id: Mapped[str] = mapped_column(String(256), nullable=False)  # Message-ID / Gmail id / IMAP UID
    thread_id: Mapped[str | None] = mapped_column(String(256), nullable=True)

    from_email: Mapped[str | None] = mapped_column(String(512), nullable=True)
    subject: Mapped[str | None] = mapped_column(String(1024), nullable=True)
    date: Mapped[dt.datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)

    snippet: Mapped[str | None] = mapped_column(Text, nullable=True)
    body_text: Mapped[str | None] = mapped_column(Text, nullable=True)
    body_html: Mapped[str | None] = mapped_column(Text, nullable=True)
    extracted_links_json: Mapped[str | None] = mapped_column(Text, nullable=True)
    extracted_images_json: Mapped[str | None] = mapped_column(Text, nullable=True)

    category: Mapped[str | None] = mapped_column(String(32), nullable=True)  # important|normal|newsletter|spam_candidate
    score: Mapped[int | None] = mapped_column(Integer, nullable=True)  # 0..100
    summary: Mapped[str | None] = mapped_column(Text, nullable=True)

    ai_model: Mapped[str | None] = mapped_column(String(128), nullable=True)
    ai_explanation: Mapped[str | None] = mapped_column(Text, nullable=True)
    ai_processed_at: Mapped[dt.datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    ai_done: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)

    is_archived: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    is_read: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)

    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=lambda: dt.datetime.now(dt.UTC))


class EmailAttachment(Base):
    __tablename__ = "email_attachments"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    email_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)

    filename: Mapped[str | None] = mapped_column(String(512), nullable=True)
    content_type: Mapped[str | None] = mapped_column(String(256), nullable=True)
    size_bytes: Mapped[int | None] = mapped_column(Integer, nullable=True)
    content_id: Mapped[str | None] = mapped_column(String(512), nullable=True, index=True)  # для cid: изображений
    is_inline: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)

    data: Mapped[bytes | None] = mapped_column(LargeBinary, nullable=True)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=lambda: dt.datetime.now(dt.UTC))


class AppSetting(Base):
    __tablename__ = "app_settings"

    key: Mapped[str] = mapped_column(String(128), primary_key=True)
    value: Mapped[str | None] = mapped_column(Text, nullable=True)
    updated_at: Mapped[dt.datetime] = mapped_column(DateTime(timezone=True), default=lambda: dt.datetime.now(dt.UTC))

