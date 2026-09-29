FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py Text_Extraction___.py ./
COPY frontend ./frontend
COPY input_images ./input_images
COPY comparison_results ./comparison_results

RUN mkdir -p outputs

CMD exec uvicorn app:app --host 0.0.0.0 --port ${PORT}