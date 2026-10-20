FROM python:3.10-slim

# Устанавливаем Tesseract OCR с русским языком
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-rus \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Папки для загрузок и результатов
RUN mkdir -p uploads results

EXPOSE 5000

CMD ["python", "app.py"]