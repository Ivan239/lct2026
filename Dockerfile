# Бэкенд + рендер превью (LibreOffice и pdftoppm нужны для картинок слайдов)
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress \
        poppler-utils \
        fonts-dejavu \
        fonts-liberation \
        curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src ./src
COPY prompts ./prompts
COPY scripts ./scripts
COPY vendor ./vendor
COPY samples ./samples
COPY docs ./docs
COPY README.md .

RUN mkdir -p output/templates output/parsed output/rendered output/generated output/evaluations

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s \
    CMD curl -fs http://localhost:8000/api/models || exit 1

CMD ["uvicorn", "src.api.main:app", "--app-dir", ".", "--host", "0.0.0.0", "--port", "8000"]
