FROM python:3.11-slim

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        fontconfig \
        wget \
        cabextract \
        xfonts-utils \
        fonts-liberation \
        fonts-crosextra-carlito \
        fonts-crosextra-caladea && \
    rm -rf /var/lib/apt/lists/*

RUN sed -i 's/Components: main/Components: main contrib/' /etc/apt/sources.list.d/debian.sources && \
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get update && \
    apt-get install -y --no-install-recommends ttf-mscorefonts-installer && \
    rm -rf /var/lib/apt/lists/*

RUN fc-cache -f -v

WORKDIR /app

COPY requirements-4.txt .
RUN pip install --no-cache-dir -r requirements-4.txt

COPY bot-2.py .

CMD ["python", "bot-2.py"]
