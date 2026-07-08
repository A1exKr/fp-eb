#!/bin/sh
# Container entrypoint: apply DB migrations, optionally seed, then serve.
set -e

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

