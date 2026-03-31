from __future__ import annotations

import secrets

from app.queue import get_redis


def issue_state(purpose: str, ttl_s: int = 600) -> str:
    state = secrets.token_urlsafe(24)
    r = get_redis()
    r.setex(f"oauth_state:{state}", ttl_s, purpose)
    return state


def consume_state(state: str) -> str | None:
    r = get_redis()
    key = f"oauth_state:{state}"
    val = r.get(key)
    if val is None:
        return None
    r.delete(key)
    return val.decode("utf-8", errors="ignore")

