from __future__ import annotations

from sqlalchemy import text

from app.db import engine


def ensure_schema() -> None:
    """
    Минимальная “миграция” для ранних итераций без Alembic.

    Делаем максимально просто: используем Postgres `IF NOT EXISTS`,
    чтобы схема могла обновляться поверх уже созданной БД.
    """
    with engine.begin() as conn:
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS imap_folder VARCHAR(128)"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS imap_last_uid INTEGER"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_at TIMESTAMPTZ"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_status VARCHAR(32)"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_count INTEGER"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_error TEXT"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS gmail_email VARCHAR(256)"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS gmail_credentials_enc TEXT"))
        conn.execute(text("ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS gmail_last_history_id VARCHAR(128)"))

        conn.execute(text("ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_model VARCHAR(128)"))
        conn.execute(text("ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_explanation TEXT"))
        conn.execute(text("ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_processed_at TIMESTAMPTZ"))
        conn.execute(text("ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_done BOOLEAN NOT NULL DEFAULT FALSE"))
        conn.execute(text("ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS is_archived BOOLEAN NOT NULL DEFAULT FALSE"))
        conn.execute(text("ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS is_read BOOLEAN NOT NULL DEFAULT TRUE"))

