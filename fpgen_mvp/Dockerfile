FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FPGEN_ENABLE_INDD_EXPORT=false

COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY data ./data
# Static reference data copied outside the persistent-volume mount path
# so a volume at /app/data does not shadow these files.
RUN mkdir -p /app/appdata && \
    cp /app/data/personnel.json /app/appdata/personnel.json && \
    cp /app/data/team_presets.json /app/appdata/team_presets.json && \
    cp /app/data/reference_projects.json /app/appdata/reference_projects.json && \
    cp /app/data/assets_registry.json /app/appdata/assets_registry.json

ENV FPGEN_STORAGE_DIR=/app/data/proposals
ENV FPGEN_EXPORT_DIR=/app/data/exports
ENV FPGEN_REFERENCE_PROJECTS=/app/appdata/reference_projects.json
ENV FPGEN_PERSONNEL=/app/appdata/personnel.json
ENV FPGEN_TEAM_PRESETS=/app/appdata/team_presets.json

EXPOSE 8080

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080"]
