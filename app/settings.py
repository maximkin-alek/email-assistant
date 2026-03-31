from __future__ import annotations

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    database_url: str = "postgresql+psycopg://email_assistant:email_assistant@localhost:5432/email_assistant"
    redis_url: str = "redis://localhost:6379/0"
    app_encryption_key: str | None = None

    yandex_imap_host: str = "imap.yandex.ru"
    yandex_imap_port: int = 993
    yandex_imap_user: str | None = None
    yandex_imap_password: str | None = None

    mailru_imap_host: str = "imap.mail.ru"
    mailru_imap_port: int = 993
    mailru_imap_user: str | None = None
    mailru_imap_password: str | None = None

    gmail_oauth_client_id: str | None = None
    gmail_oauth_client_secret: str | None = None
    gmail_oauth_redirect_uri: str = "http://localhost:8000/oauth2/google/callback"

    # AI (OpenAI-compatible)
    ai_base_url: str = "https://openrouter.ai/api/v1"
    ai_api_key: str | None = None
    ai_model: str = "anthropic/claude-3.5-haiku"

    # Backward-compat (старые переменные)
    openrouter_api_key: str | None = None
    openrouter_model: str | None = None


settings = Settings()

# Миграция конфигов: если заданы старые переменные — используем их как значения по умолчанию.
if settings.openrouter_api_key and not settings.ai_api_key:
    settings.ai_api_key = settings.openrouter_api_key
if settings.openrouter_model and (settings.ai_model == "anthropic/claude-3.5-haiku"):
    settings.ai_model = settings.openrouter_model

