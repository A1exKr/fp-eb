#!/bin/sh
# Container entrypoint: wait for DB, apply migrations, optionally seed, then serve.
set -e

# --- Wait for the database to accept connections before migrating. The rdb
# (Postgres) sideapp cold start can take ~30s (permission init + WAL recovery);
# without this the api races ahead, fails `alembic upgrade head`, and the whole
# app 502s until a retry happens to land after Postgres is finally ready.
echo "[entrypoint] Waiting for the database to become ready..."
python - <<'PY'
import sys
import time

from sqlalchemy import text

from app.db import engine

deadline = time.time() + 180
attempt = 0
while True:
    attempt += 1
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print(f"[entrypoint] Database ready (after {attempt} check(s)).", flush=True)
        break
    except Exception as exc:  # noqa: BLE001 - any connect error means "not ready yet"
        if time.time() >= deadline:
            print(f"[entrypoint] Database not ready after 180s: {exc}", file=sys.stderr)
            sys.exit(1)
        print("[entrypoint] Database not ready yet; retrying in 3s...", flush=True)
        time.sleep(3)
PY

echo "[entrypoint] Applying database migrations (alembic upgrade head)..."
# Retry to tolerate concurrent first-run migrations across replicas: once one
# replica applies 0001, the others see head and no-op.
n=0
until alembic upgrade head; do
  n=$((n + 1))
  if [ "$n" -ge 5 ]; then
    echo "[entrypoint] Migrations failed after $n attempts." >&2
    exit 1
  fi
  echo "[entrypoint] Migration attempt $n failed; retrying in 3s..."
  sleep 3
done

if [ "${FPGEN_SEED_ON_START}" = "true" ]; then
  echo "[entrypoint] FPGEN_SEED_ON_START=true -> seeding master data from JSON..."
  python scripts/import_json_to_db.py
fi

echo "[entrypoint] Starting API on :8080"
exec uvicorn app.main:app --host 0.0.0.0 --port 8080

