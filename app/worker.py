from __future__ import annotations

from rq import Worker

from app.db import Base, engine
from app.queue import get_queue, get_redis
from app.schema import ensure_schema


def main() -> None:
    # Чтобы воркер не падал на новых колонках при рестарте/миграциях
    Base.metadata.create_all(bind=engine)
    ensure_schema()
    queues = [get_queue("default")]
    worker = Worker(queues, connection=get_redis())
    worker.work(with_scheduler=False)


if __name__ == "__main__":
    main()

