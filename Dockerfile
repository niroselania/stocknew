FROM python:3.11-alpine

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY server.py stock_processor.py stock-buscador.html app.js xlsx-worker.js manifest.webmanifest sw.js icon.svg ./

ENV APP_DIR=/app
ENV DATA_DIR=/data
ENV PORT=8000
ENV HISTORY_MAX=30

EXPOSE 8000

CMD ["python", "server.py"]
