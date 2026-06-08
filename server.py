from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
import cgi
import json
import os
import shutil
import time

from stock_processor import build_stock_payload, compare_inventories, make_history_id


APP_DIR = Path(os.environ.get("APP_DIR", "/app"))
DATA_DIR = Path(os.environ.get("DATA_DIR", "/data"))
HISTORY_DIR = DATA_DIR / "history"
HTML_FILE = APP_DIR / "stock-buscador.html"
STOCK_FILE = DATA_DIR / "stock.xlsx"
META_FILE = DATA_DIR / "stock.json"
DATA_FILE = DATA_DIR / "stock-data.json"
HISTORY_INDEX_FILE = DATA_DIR / "history-index.json"
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", str(100 * 1024 * 1024)))
HISTORY_MAX = int(os.environ.get("HISTORY_MAX", "30"))

STATIC_FILES = {
    "/app.js": ("app.js", "application/javascript; charset=utf-8"),
    "/xlsx-worker.js": ("xlsx-worker.js", "application/javascript; charset=utf-8"),
    "/manifest.webmanifest": ("manifest.webmanifest", "application/manifest+json; charset=utf-8"),
    "/sw.js": ("sw.js", "application/javascript; charset=utf-8"),
    "/icon.svg": ("icon.svg", "image/svg+xml"),
}


def load_history_index() -> list[dict]:
    if not HISTORY_INDEX_FILE.exists():
        return []
    return json.loads(HISTORY_INDEX_FILE.read_text(encoding="utf-8"))


def save_history_index(entries: list[dict]) -> None:
    HISTORY_INDEX_FILE.write_text(
        json.dumps(entries, ensure_ascii=False),
        encoding="utf-8",
    )


def trim_history(entries: list[dict]) -> list[dict]:
    trimmed = entries[:HISTORY_MAX]
    keep_ids = {entry["id"] for entry in trimmed}
    for entry in entries[HISTORY_MAX:]:
        history_xlsx = HISTORY_DIR / f"{entry['id']}.xlsx"
        history_json = HISTORY_DIR / f"{entry['id']}.json"
        if history_xlsx.exists():
            history_xlsx.unlink()
        if history_json.exists():
            history_json.unlink()
    return trimmed


def process_upload(stock_path: Path, original_name: str, uploaded_at: int | None = None) -> dict:
    uploaded_at = uploaded_at or int(time.time())
    payload = build_stock_payload(stock_path)
    history_id = make_history_id(uploaded_at)

    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(stock_path, HISTORY_DIR / f"{history_id}.xlsx")

    snapshot = {
        "id": history_id,
        "name": original_name,
        "uploadedAt": uploaded_at,
        "size": stock_path.stat().st_size,
        "productCount": payload["productCount"],
        "inventory": payload["inventory"],
    }
    (HISTORY_DIR / f"{history_id}.json").write_text(
        json.dumps(snapshot, ensure_ascii=False),
        encoding="utf-8",
    )

    entries = load_history_index()
    entries.insert(
        0,
        {
            "id": history_id,
            "name": original_name,
            "uploadedAt": uploaded_at,
            "size": snapshot["size"],
            "productCount": snapshot["productCount"],
        },
    )
    save_history_index(trim_history(entries))

    meta = {
        "id": history_id,
        "name": original_name,
        "uploadedAt": uploaded_at,
        "size": snapshot["size"],
        "productCount": payload["productCount"],
    }
    META_FILE.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    DATA_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return meta


class StockHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/" or path.startswith("/?"):
            return self.send_file(HTML_FILE, "text/html; charset=utf-8")

        if path in STATIC_FILES:
            filename, content_type = STATIC_FILES[path]
            return self.send_file(APP_DIR / filename, content_type)

        if path == "/api/health":
            return self.send_json({"ok": True, "hasStock": STOCK_FILE.exists()})

        if path == "/api/version":
            return self.send_json(
                {
                    "version": "2.1.0",
                    "uploadFix": True,
                    "features": ["history", "compare", "stock-data", "preprocess"],
                }
            )

        if path == "/api/current":
            if not STOCK_FILE.exists() or not META_FILE.exists():
                return self.send_text("No hay planilla cargada", status=404)
            return self.send_file(META_FILE, "application/json; charset=utf-8", no_cache=True)

        if path == "/api/stock-data":
            if not DATA_FILE.exists():
                return self.send_text("No hay datos procesados", status=404)
            return self.send_file(DATA_FILE, "application/json; charset=utf-8", no_cache=True)

        if path == "/api/stock.xlsx":
            if not STOCK_FILE.exists():
                return self.send_text("No hay planilla cargada", status=404)
            return self.send_file(
                STOCK_FILE,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                no_cache=True,
            )

        if path == "/api/history":
            return self.send_json(load_history_index())

        if path.startswith("/api/compare"):
            return self.handle_compare(parse_qs(parsed.query))

        return self.send_text("No encontrado", status=404)

    def handle_compare(self, query: dict):
        if not DATA_FILE.exists() or not META_FILE.exists():
            return self.send_text("No hay planilla actual para comparar", status=404)

        current_meta = json.loads(META_FILE.read_text(encoding="utf-8"))
        current_payload = json.loads(DATA_FILE.read_text(encoding="utf-8"))
        history = load_history_index()

        compare_id = (query.get("id") or [""])[0]
        if not compare_id:
            if len(history) < 2:
                return self.send_text("No hay una planilla anterior para comparar", status=404)
            compare_id = history[1]["id"]

        previous_file = HISTORY_DIR / f"{compare_id}.json"
        if not previous_file.exists():
            return self.send_text("No encontré esa planilla en el historial", status=404)

        previous_snapshot = json.loads(previous_file.read_text(encoding="utf-8"))
        comparison = compare_inventories(
            current_payload.get("inventory", {}),
            previous_snapshot.get("inventory", {}),
        )

        return self.send_json(
            {
                "current": current_meta,
                "previous": {
                    "id": previous_snapshot["id"],
                    "name": previous_snapshot["name"],
                    "uploadedAt": previous_snapshot["uploadedAt"],
                },
                **comparison,
            }
        )

    def do_POST(self):
        if self.path != "/api/upload":
            return self.send_text("No encontrado", status=404)

        length = int(self.headers.get("content-length", "0"))
        if length <= 0:
            return self.send_text("Archivo vacio", status=400)
        if length > MAX_UPLOAD_BYTES:
            return self.send_text("La planilla supera el limite permitido", status=413)

        content_type = self.headers.get("content-type", "")
        if "multipart/form-data" not in content_type:
            return self.send_text("La subida debe ser multipart/form-data", status=400)

        DATA_DIR.mkdir(parents=True, exist_ok=True)
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": content_type,
                "CONTENT_LENGTH": str(length),
            },
        )

        field = form["stock"] if "stock" in form else None
        if field is None or not getattr(field, "filename", ""):
            return self.send_text("No recibi el archivo stock", status=400)
        if not field.filename.lower().endswith(".xlsx"):
            return self.send_text("El archivo debe ser .xlsx", status=400)

        upload_file = DATA_DIR / "stock-upload.xlsx"
        with upload_file.open("wb") as output:
            shutil.copyfileobj(field.file, output)

        try:
            meta = process_upload(upload_file, Path(field.filename).name)
            upload_file.replace(STOCK_FILE)
            self.send_json(meta)
        except Exception as error:
            if upload_file.exists():
                upload_file.unlink()
            self.send_text(f"No pude procesar la planilla: {error}", status=400)

    def send_file(self, path, content_type, no_cache=False):
        if not path.exists():
            return self.send_text("No encontrado", status=404)

        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        if no_cache:
            self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def send_text(self, text, status=200):
        data = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        print("%s - %s" % (self.address_string(), fmt % args), flush=True)


def ensure_processed_stock() -> None:
    if not STOCK_FILE.exists() or DATA_FILE.exists():
        return

    try:
        meta = process_upload(STOCK_FILE, STOCK_FILE.name)
        print(f"Planilla existente procesada: {meta['name']}", flush=True)
    except Exception as error:
        print(f"No pude reprocesar la planilla existente: {error}", flush=True)


if __name__ == "__main__":
    try:
        import openpyxl

        print(f"openpyxl {openpyxl.__version__} OK", flush=True)
    except ImportError as error:
        print(f"ERROR: falta openpyxl ({error}). Rebuild la imagen Docker.", flush=True)

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    ensure_processed_stock()
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), StockHandler)
    print(f"Stock server v2.1.0 listening on {port}", flush=True)
    server.serve_forever()
