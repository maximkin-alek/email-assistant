from __future__ import annotations

from cryptography.fernet import Fernet, InvalidToken

from app.settings import settings


def _fernet() -> Fernet:
    if not settings.app_encryption_key:
        raise RuntimeError("APP_ENCRYPTION_KEY не задан. Сгенерируй ключ и положи в .env")
    key = settings.app_encryption_key.encode("utf-8") if isinstance(settings.app_encryption_key, str) else settings.app_encryption_key
    return Fernet(key)


def encrypt_str(value: str) -> str:
    return _fernet().encrypt(value.encode("utf-8")).decode("utf-8")


def decrypt_str(value: str) -> str:
    try:
        return _fernet().decrypt(value.encode("utf-8")).decode("utf-8")
    except InvalidToken as e:
        raise ValueError("Не удалось расшифровать значение (неверный ключ?)") from e

