"""Database engine, session, and declarative base.

Targets exaBase Studio PostgreSQL (schema ``fpgen``) in production and falls
back to a local SQLite file for development/tests. The ORM models are defined
against the configured schema; for SQLite the schema is translated away so the
same models work in both environments.
"""
from collections.abc import Iterator

from sqlalchemy import JSON, MetaData, create_engine, text
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from app.config import settings


def _is_sqlite(url: str) -> bool:
    return url.startswith("sqlite")


IS_SQLITE = _is_sqlite(settings.database_url)

# Schema is applied on PostgreSQL only; SQLite has no real schema support.
DB_SCHEMA = None if IS_SQLITE else (settings.db_schema or None)

# JSONB on PostgreSQL, generic JSON elsewhere (SQLite).
JSONType = JSON().with_variant(JSONB(), "postgresql")

metadata_obj = MetaData(schema=DB_SCHEMA)


class Base(DeclarativeBase):
    metadata = metadata_obj


_connect_args = {"check_same_thread": False} if IS_SQLITE else {}

engine = create_engine(
    settings.database_url,
    echo=settings.db_echo,
    pool_pre_ping=True,
    future=True,
    connect_args=_connect_args,
)

SessionLocal = sessionmaker(
    bind=engine,
    autoflush=False,
    autocommit=False,
    expire_on_commit=False,
    class_=Session,
)


def get_db() -> Iterator[Session]:
    """FastAPI dependency yielding a scoped database session."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db() -> None:
    """Create the schema (PostgreSQL) and all tables.

    Convenience for local development, tests, and first-run bootstrap. On
    exaBase PostgreSQL the schema is created by the ``rdb`` init SQL and the
    tables by Alembic migrations, but this remains safe and idempotent.
    """
    # Import models so they are registered on Base.metadata before create_all.
    from app import models  # noqa: F401

    if DB_SCHEMA:
        with engine.begin() as conn:
            conn.execute(text(f'CREATE SCHEMA IF NOT EXISTS "{DB_SCHEMA}"'))
    Base.metadata.create_all(bind=engine)
