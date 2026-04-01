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
        # Чтобы не зависать на старте при DDL-локах
        conn.execute(text("SET LOCAL lock_timeout = '5s'"))
        conn.execute(text("SET LOCAL statement_timeout = '30s'"))
        def _safe(sql: str) -> None:
            try:
                conn.execute(text(sql))
            except Exception:
                # Любой DDL-lock/timeout не должен валить приложение.
                pass

        for sql in [
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS imap_folder VARCHAR(128)",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS imap_last_uid INTEGER",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS imap_tls_verify BOOLEAN NOT NULL DEFAULT TRUE",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_at TIMESTAMPTZ",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_status VARCHAR(32)",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_count INTEGER",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS last_sync_error TEXT",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS gmail_email VARCHAR(256)",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS gmail_credentials_enc TEXT",
            "ALTER TABLE mailboxes ADD COLUMN IF NOT EXISTS gmail_last_history_id VARCHAR(128)",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_model VARCHAR(128)",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_explanation TEXT",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_processed_at TIMESTAMPTZ",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS ai_done BOOLEAN NOT NULL DEFAULT FALSE",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS is_archived BOOLEAN NOT NULL DEFAULT FALSE",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS is_read BOOLEAN NOT NULL DEFAULT TRUE",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS body_html TEXT",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS extracted_links_json TEXT",
            "ALTER TABLE email_messages ADD COLUMN IF NOT EXISTS extracted_images_json TEXT",
        ]:
            _safe(sql)

        try:
            conn.execute(
                text(
                    """
                    CREATE TABLE IF NOT EXISTS email_attachments (
                        id SERIAL PRIMARY KEY,
                        email_id INTEGER NOT NULL,
                        filename VARCHAR(512),
                        content_type VARCHAR(256),
                        size_bytes INTEGER,
                        content_id VARCHAR(512),
                        is_inline BOOLEAN NOT NULL DEFAULT FALSE,
                        data BYTEA,
                        created_at TIMESTAMPTZ DEFAULT NOW()
                    )
                    """
                )
            )
            conn.execute(text("CREATE INDEX IF NOT EXISTS ix_email_attachments_email_id ON email_attachments(email_id)"))
            conn.execute(text("CREATE INDEX IF NOT EXISTS ix_email_attachments_content_id ON email_attachments(content_id)"))
        except Exception:
            # Если не смогли получить DDL-лок быстро — не блокируем старт приложения.
            pass

        try:
            conn.execute(
                text(
                    """
                    CREATE TABLE IF NOT EXISTS app_settings (
                        key VARCHAR(128) PRIMARY KEY,
                        value TEXT,
                        updated_at TIMESTAMPTZ DEFAULT NOW()
                    )
                    """
                )
            )
        except Exception:
            pass

