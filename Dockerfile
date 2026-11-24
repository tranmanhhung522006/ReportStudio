FROM python:3.11-slim
ENV PYTHONUNBUFFERED=1
WORKDIR /app
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer fonts-liberation fonts-dejavu-core && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN useradd -m appuser
USER appuser
CMD ["sh","-c","gunicorn production:app --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 4 --timeout 120"]
