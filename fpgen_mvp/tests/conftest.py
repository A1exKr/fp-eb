"""Test configuration: use an isolated SQLite DB and disable external services.

Environment must be set before any ``app`` import so that ``app.config.settings``
and ``app.db`` pick up the test database.
"""
import os

os.environ.setdefault("DATABASE_URL", "sqlite:///./data/test_fpgen.db")
os.environ.setdefault("FPGEN_AUTH_ENABLED", "false")
os.environ.setdefault("FPGEN_ENABLE_INDD_EXPORT", "false")
os.environ.setdefault("ENABLE_OPENAI_SYNTHESIS", "false")

from pathlib import Path  # noqa: E402

_db = Path("./data/test_fpgen.db")
if _db.exists():
    _db.unlink()
