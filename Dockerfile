# CoA -> Word Template Filler
FROM python:3.11-slim

# build-essential in case any dependency needs to compile (most ship manylinux
# wheels for linux/amd64 & arm64, this is just a safety net).
# tesseract-ocr is the actual OCR engine used to read scanned/photographed
# PDF pages in the "keep original layout" mode; libgl1/libglib2.0-0 are
# runtime libraries opencv-python-headless needs even without a display.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        build-essential \
        tesseract-ocr \
        libgl1 \
        libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/ backend/
COPY frontend/ frontend/
COPY templates/ templates/
COPY data/ data/
COPY forms/ forms/
RUN mkdir -p generated

WORKDIR /app/backend

EXPOSE 8420

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8420"]
