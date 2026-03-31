from __future__ import annotations

from contextlib import contextmanager

from sqlalchemy import create_engine
from sqlalchemy import event
from sqlalchemy.pool import Pool
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from app.settings import settings


class Base(DeclarativeBase):
    pass


engine = create_engine(settings.database_url, pool_pre_ping=True)


@event.listens_for(Pool, "connect")
def _disable_prepared_statements(dbapi_connection, connection_record) -> None:
    """
    psycopg3 может автоматически использовать prepared statements, что конфликтует с форком воркера (RQ)
    и/или с пулерами. Самое простое и надёжное решение для MVP — отключить auto-prepare.
    """
    if hasattr(dbapi_connection, "prepare_threshold"):
        try:
            dbapi_connection.prepare_threshold = None
        except Exception:
            # Если драйвер не psycopg3 или атрибут read-only — просто игнорируем.
            pass

SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, expire_on_commit=False)


@contextmanager
def session_scope() -> Session:
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()

