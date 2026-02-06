FROM python:3.10-slim

# Install system dependencies
# poppler-utils: for pdf2image
# tesseract-ocr: for pytesseract
# libgl1: for opengl support (opencv usually)
RUN apt-get update && apt-get install -y \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-spa \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code
COPY . .

# code fix applied in api/services/preprocess.py

# Expose port
EXPOSE 8000

# Run the Fast API application
CMD ["uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8000"]
