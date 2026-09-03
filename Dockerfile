# Gridlint runs anywhere Python does. Hugging Face Spaces and Render both use
# this file unchanged; PORT is read from the environment (Spaces uses 7860).
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    GRIDLINT_DATA=/data \
    PORT=7860

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY gridlint/ ./gridlint/
COPY samples/ ./samples/
COPY fixtures/ ./fixtures/
COPY README.md pyproject.toml ./

# A writable data directory for the SQLite database and stored workbooks.
RUN mkdir -p /data && chmod 777 /data

EXPOSE 7860
HEALTHCHECK --interval=60s --timeout=5s --start-period=20s \
  CMD python -c "import urllib.request,os;urllib.request.urlopen(f'http://127.0.0.1:{os.environ.get(\"PORT\",7860)}/api/health')"

CMD ["sh", "-c", "uvicorn gridlint.server:app --host 0.0.0.0 --port ${PORT:-7860}"]
