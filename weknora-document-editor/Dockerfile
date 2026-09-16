FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=7860 \
    OUTPUT_DIR=/data/generated

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY server_sse.py ./
RUN useradd --create-home --uid 10001 appuser && mkdir -p /data/generated && chown -R appuser:appuser /app /data
USER appuser
EXPOSE 7860
CMD ["python", "server_sse.py"]
