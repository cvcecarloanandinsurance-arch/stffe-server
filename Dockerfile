FROM python:3.11-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr tesseract-ocr-eng tesseract-ocr-tam \
    && rm -rf /var/lib/apt/lists/*
# OMP_THREAD_LIMIT=1: Tesseract uses 1 thread (less RAM, no CPU fights on small plans)
# MALLOC_ARENA_MAX=2: stops glibc from holding extra memory per thread
ENV OMP_THREAD_LIMIT=1 MALLOC_ARENA_MAX=2 PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY server.py .
# 1 worker + 2 threads: one thread does the job, the other keeps /health answering.
# --max-requests recycles the worker now and then to give memory back.
CMD gunicorn -w 1 --threads 2 -t 180 --graceful-timeout 30 --max-requests 300 --max-requests-jitter 50 --worker-tmp-dir /dev/shm -b 0.0.0.0:${PORT:-10000} server:app
