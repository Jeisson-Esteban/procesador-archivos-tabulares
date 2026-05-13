FROM python:3.11-slim

WORKDIR /app

# libmagic es requerido por python-magic para detectar mime types reales.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libmagic1 \
    wget \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8010
ENV PORT=8010

# Gunicorn con 2 workers y 60s de timeout: archivos grandes tardan en procesarse.
CMD ["gunicorn", "--bind", "0.0.0.0:8010", "--workers", "2", "--timeout", "60", "app:app"]
