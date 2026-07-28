FROM python:3.11-slim

# Install system dependencies: Tesseract (OCR) and Poppler (PDF to image)
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-eng \
