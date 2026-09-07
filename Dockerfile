# Python 3.11 temel imajı
FROM python:3.11-slim

# LibreOffice'i (pptx -> pdf dönüştürme için) sisteme kur
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python bağımlılıklarını kur
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Bot kodunu kopyala
COPY bot.py .

# Botu başlat
CMD ["python", "bot.py"]
