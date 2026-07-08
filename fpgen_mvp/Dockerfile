FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FPGEN_ENABLE_INDD_EXPORT=false

COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY data ./data
COPY alembic ./alembic
COPY alembic.ini ./alembic.ini
COPY scripts ./scripts
COPY entrypoint.sh ./entrypoint.sh
# Static reference data copied outside the persistent-volume mount path
# so a volume at /app/data does not shadow these files (used by the seed script).
RUN mkdir -p /app/appdata && \
    cp /app/data/personnel.json /app/appdata/personnel.json && \
    cp /app/data/team_presets.json /app/appdata/team_presets.json && \
    cp /app/data/reference_projects.json /app/appdata/reference_projects.json && \
    cp /app/data/assets_registry.json /app/appdata/assets_registry.json && \
    sed -i 's/\r$//' /app/entrypoint.sh && chmod +x /app/entrypoint.sh

ENV FPGEN_STORAGE_DIR=/app/data/proposals
ENV FPGEN_EXPORT_DIR=/app/data/exports
ENV FPGEN_ASSET_STORAGE_DIR=/app/data/asset_files
ENV FPGEN_REFERENCE_PROJECTS=/app/appdata/reference_projects.json
ENV FPGEN_PERSONNEL=/app/appdata/personnel.json
ENV FPGEN_TEAM_PRESETS=/app/appdata/team_presets.json
ENV FPGEN_ASSETS_REGISTRY=/app/appdata/assets_registry.json

EXPOSE 8080

# Applies DB migrations (and optional seed) before starting uvicorn.
ENTRYPOINT ["sh", "/app/entrypoint.sh"]
