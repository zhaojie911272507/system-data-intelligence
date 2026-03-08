FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-nogui \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /app/outputs /data

ENV PYTHONUNBUFFERED=1
ENV MPLBACKEND=Agg

CMD ["python", "-c", "print('system-data-intelligence ready')"]
