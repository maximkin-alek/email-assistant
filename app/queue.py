from __future__ import annotations

from redis import Redis
from rq import Queue

from app.settings import settings


def get_redis() -> Redis:
    return Redis.from_url(settings.redis_url)


def get_queue(name: str = "default") -> Queue:
    return Queue(name, connection=get_redis())

