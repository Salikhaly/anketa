# Парсер ПКБ/ГКБ — контейнер для Google Cloud Run
FROM python:3.11-slim

# poppler + tesseract (rus/kaz) — для сканированных (нетекстовых) PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
        poppler-utils \
        tesseract-ocr tesseract-ocr-rus tesseract-ocr-kaz \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PORT=8080
# Cloud Run передаёт порт через $PORT
CMD exec uvicorn api:app --host 0.0.0.0 --port ${PORT}
