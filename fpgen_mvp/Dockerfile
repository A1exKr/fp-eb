FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FPGEN_ENABLE_INDD_EXPORT=false

COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY data ./data

ENV FPGEN_STORAGE_DIR=/app/data/proposals
ENV FPGEN_EXPORT_DIR=/app/data/exports
ENV FPGEN_REFERENCE_PROJECTS=/app/data/reference_projects.json

EXPOSE 8080

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080"]
