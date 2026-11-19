from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
import cgi
import gzip
import hmac
import json
import os
import re
import shutil
import threading
import time

from stock_processor import build_stock_payload, compare_inventories, make_history_id
from telegram_bot import start_bot

VERSION = "3.1.0"
TELEGRAM_BOT = None

APP_DIR = Path(os.environ.get("APP_DIR", "/app"))
DATA_DIR = Path(os.environ.get("DATA_DIR", "/data"))
HISTORY_DIR = DATA_DIR / "history"
HTML_FILE = APP_DIR / "stock-buscador.html"
STOCK_FILE = DATA_DIR / "stock.xlsx"
META_FILE = DATA_DIR / "stock.json"
DATA_FILE = DATA_DIR / "stock-data.json"
DATA_FILE_GZ = DATA_DIR / "stock-data.json.gz"
HISTORY_INDEX_FILE = DATA_DIR / "history-index.json"
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", str(100 * 1024 * 1024)))
HISTORY_MAX = int(os.environ.get("HISTORY_MAX", "30"))
# Opcional: si se define, subir o borrar planillas pide este código.
UPLOAD_TOKEN = os.environ.get("UPLOAD_TOKEN", "").strip()

SAFE_ID = re.compile(r"^[0-9A-Za-z_\-]+$")
PROCESS_LOCK = threading.Lock()

STATIC_FILES = {
    "/app.js": ("app.js", "application/javascript; charset=utf-8"),
    "/xlsx-worker.js": ("xlsx-worker.js", "application/javascript; charset=utf-8"),
    "/manifest.webmanifest": ("manifest.webmanifest", "application/manifest+json; charset=utf-8"),
    "/sw.js": ("sw.js", "application/javascript; charset=utf-8"),
    "/icon.svg": ("icon.svg", "image/svg+xml"),
}
COMPRESSIBLE = ("application/json", "application/javascript", "text/", "image/svg+xml", "application/manifest")


def write_json_atomic(path: Path, payload, compress: bool = False) -> None:
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_bytes(data)
    tmp.replace(path)
    if compress:
        gz_path = path.with_suffix(path.suffix + ".gz")
        gz_tmp = gz_path.with_suffix(".gz.tmp")
        gz_tmp.write_bytes(gzip.compress(data, compresslevel=6))
        gz_tmp.replace(gz_path)


def read_json(path: Path, default=None):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def load_history_index() -> list[dict]:
    return read_json(HISTORY_INDEX_FILE, [])


def trim_history(entries: list[dict]) -> list[dict]:
    for entry in entries[HISTORY_MAX:]:
        remove_history_files(entry["id"])
    return entries[:HISTORY_MAX]


def remove_history_files(history_id: str) -> None:
    for ext in ("xlsx", "json"):
        file = HISTORY_DIR / f"{history_id}.{ext}"
        if file.exists():
            file.unlink()


def unique_history_id(uploaded_at: int) -> str:
    base = make_history_id(uploaded_at)
    candidate, n = base, 1
    while (HISTORY_DIR / f"{candidate}.json").exists():
        n += 1
        candidate = f"{base}_{n}"
    return candidate


def process_upload(stock_path: Path, original_name: str, uploaded_at: int | None = None) -> dict:
    uploaded_at = uploaded_at or int(time.time())
    started = time.time()
    payload = build_stock_payload(stock_path)
    history_id = unique_history_id(uploaded_at)

    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(stock_path, HISTORY_DIR / f"{history_id}.xlsx")

    size = stock_path.stat().st_size
    write_json_atomic(
        HISTORY_DIR / f"{history_id}.json",
        {
            "id": history_id,
            "name": original_name,
            "uploadedAt": uploaded_at,
            "size": size,
            "productCount": payload["productCount"],
            "totals": payload["totals"],
            "inventory": payload["inventory"],
        },
    )

    meta = {
        "id": history_id,
        "name": original_name,
        "uploadedAt": uploaded_at,
        "size": size,
        "productCount": payload["productCount"],
        "totals": payload["totals"],
        "processSeconds": round(time.time() - started, 2),
        "priceColumn": payload.get("priceColumn"),
    }
    entries = load_history_index()
    entries.insert(0, meta)
    write_json_atomic(HISTORY_INDEX_FILE, trim_history(entries))

    # El navegador no necesita "inventory" (sólo se usa para comparar): payload más liviano.
    client_payload = {key: value for key, value in payload.items() if key != "inventory"}
    client_payload["dataVersion"] = history_id
    write_json_atomic(DATA_FILE, client_payload, compress=True)
    write_json_atomic(META_FILE, meta)
    return meta


def load_inventory(history_id: str):
    snapshot = read_json(HISTORY_DIR / f"{history_id}.json")
    return snapshot.get("inventory") if snapshot else None


class StockHandler(BaseHTTPRequestHandler):
    server_version = f"StockBuscador/{VERSION}"

    # ---------- GET ----------
    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)

        if path in ("/", "/index.html", "/stock-buscador.html"):
            return self.send_file(HTML_FILE, "text/html; charset=utf-8", revalidate=True)

        if path in STATIC_FILES:
            filename, content_type = STATIC_FILES[path]
            return self.send_file(APP_DIR / filename, content_type, revalidate=True)

        if path == "/api/health":
            telegram = dict(TELEGRAM_BOT.status) if TELEGRAM_BOT else {"enabled": False}
            return self.send_json({"ok": True, "hasStock": STOCK_FILE.exists(), "version": VERSION, "telegram": telegram})

        if path == "/api/version":
            return self.send_json(
                {
                    "version": VERSION,
                    "uploadFix": True,
                    "authRequired": bool(UPLOAD_TOKEN),
                    "historyMax": HISTORY_MAX,
                    "features": ["history", "compare", "stock-data", "preprocess", "gzip", "etag", "delete", "download", "price", "telegram"],
                    "telegram": bool(TELEGRAM_BOT and TELEGRAM_BOT.status.get("enabled")),
                }
            )

        if path == "/api/current":
            if not STOCK_FILE.exists() or not META_FILE.exists():
                return self.send_text("No hay planilla cargada", status=404)
            return self.send_file(META_FILE, "application/json; charset=utf-8", no_cache=True)

        if path == "/api/stock-data":
            if not DATA_FILE.exists():
                return self.send_text("No hay datos procesados", status=404)
            meta = read_json(META_FILE, {}) or {}
            return self.send_file(
                DATA_FILE,
                "application/json; charset=utf-8",
                revalidate=True,
                etag=f'"{meta.get("id", int(DATA_FILE.stat().st_mtime))}"',
                gz_path=DATA_FILE_GZ,
            )

        if path == "/api/stock.xlsx":
            if not STOCK_FILE.exists():
                return self.send_text("No hay planilla cargada", status=404)
            meta = read_json(META_FILE, {}) or {}
            return self.send_file(STOCK_FILE, XLSX_TYPE, no_cache=True, download_name=meta.get("name", "stock.xlsx"))

        if path == "/api/history":
            return self.send_json(load_history_index())

        if path == "/api/history/download":
            history_id = (query.get("id") or [""])[0]
            if not SAFE_ID.match(history_id):
                return self.send_text("Id inválido", status=400)
            entry = next((item for item in load_history_index() if item["id"] == history_id), None)
            file = HISTORY_DIR / f"{history_id}.xlsx"
            if not entry or not file.exists():
                return self.send_text("No encontré esa planilla", status=404)
            return self.send_file(file, XLSX_TYPE, no_cache=True, download_name=entry.get("name", file.name))

        if path == "/api/compare":
            return self.handle_compare(query)

        return self.send_text("No encontrado", status=404)

    def handle_compare(self, query: dict):
        current_meta = read_json(META_FILE)
        if not current_meta:
            return self.send_text("No hay planilla actual para comparar", status=404)
        history = load_history_index()

        compare_id = (query.get("id") or [""])[0]
        if not compare_id:
            older = [entry for entry in history if entry["id"] != current_meta.get("id")]
            if not older:
                return self.send_text("No hay una planilla anterior para comparar", status=404)
            compare_id = older[0]["id"]
        if not SAFE_ID.match(compare_id):
            return self.send_text("Id inválido", status=400)

        previous_file = HISTORY_DIR / f"{compare_id}.json"
        if not previous_file.exists():
            return self.send_text("No encontré esa planilla en el historial", status=404)
        previous_snapshot = read_json(previous_file)

        current_inventory = load_inventory(current_meta.get("id", ""))
        if current_inventory is None:  # datos de una versión anterior
            current_inventory = (read_json(DATA_FILE, {}) or {}).get("inventory", {})

        comparison = compare_inventories(current_inventory, previous_snapshot.get("inventory", {}))
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

    # ---------- POST / DELETE ----------
    def authorized(self) -> bool:
        if not UPLOAD_TOKEN:
            return True
        given = self.headers.get("X-Upload-Token", "")
        return hmac.compare_digest(given.encode("utf-8"), UPLOAD_TOKEN.encode("utf-8"))

    def do_POST(self):
        if urlparse(self.path).path != "/api/upload":
            return self.send_text("No encontrado", status=404)
        if not self.authorized():
            return self.send_text("Código de carga incorrecto", status=401)

        length = int(self.headers.get("content-length", "0") or 0)
        if length <= 0:
            return self.send_text("Archivo vacío", status=400)
        if length > MAX_UPLOAD_BYTES:
            return self.send_text("La planilla supera el límite permitido", status=413)

        content_type = self.headers.get("content-type", "")
        if "multipart/form-data" not in content_type:
            return self.send_text("La subida debe ser multipart/form-data", status=400)

        DATA_DIR.mkdir(parents=True, exist_ok=True)
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type, "CONTENT_LENGTH": str(length)},
        )
        field = form["stock"] if "stock" in form else None
        if field is None or not getattr(field, "filename", ""):
            return self.send_text("No recibí el archivo stock", status=400)
        if not field.filename.lower().endswith(".xlsx"):
            return self.send_text("El archivo debe ser .xlsx", status=400)

        if not PROCESS_LOCK.acquire(timeout=240):
            return self.send_text("Hay otra planilla procesándose. Probá en un minuto.", status=409)
        upload_file = DATA_DIR / "stock-upload.xlsx"
        try:
            with upload_file.open("wb") as output:
                shutil.copyfileobj(field.file, output)
            meta = process_upload(upload_file, Path(field.filename).name)
            upload_file.replace(STOCK_FILE)
            print(f"Planilla cargada: {meta['name']} ({meta['productCount']} productos, {meta['processSeconds']}s) · precio: {meta.get('priceColumn') or 'NO ENCONTRÉ columna RETAIL'}", flush=True)
            self.send_json(meta)
        except Exception as error:
            if upload_file.exists():
                upload_file.unlink()
            self.send_text(f"No pude procesar la planilla: {error}", status=400)
        finally:
            PROCESS_LOCK.release()

    def do_DELETE(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/history":
            return self.send_text("No encontrado", status=404)
        if not self.authorized():
            return self.send_text("Código de carga incorrecto", status=401)
        history_id = (parse_qs(parsed.query).get("id") or [""])[0]
        if not SAFE_ID.match(history_id):
            return self.send_text("Id inválido", status=400)
        current = read_json(META_FILE, {}) or {}
        if history_id == current.get("id"):
            return self.send_text("No se puede borrar la planilla actual", status=400)
        with PROCESS_LOCK:
            entries = load_history_index()
            remaining = [entry for entry in entries if entry["id"] != history_id]
            if len(remaining) == len(entries):
                return self.send_text("No encontré esa planilla", status=404)
            remove_history_files(history_id)
            write_json_atomic(HISTORY_INDEX_FILE, remaining)
        return self.send_json({"ok": True, "deleted": history_id})

    # ---------- helpers ----------
    def accepts_gzip(self) -> bool:
        return "gzip" in (self.headers.get("Accept-Encoding") or "").lower()

    def common_headers(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "same-origin")

    def send_file(self, path, content_type, no_cache=False, revalidate=False, etag=None, gz_path=None, download_name=None):
        if not path.exists():
            return self.send_text("No encontrado", status=404)

        stat = path.stat()
        etag = etag or f'"{int(stat.st_mtime)}-{stat.st_size}"'
        if revalidate and self.headers.get("If-None-Match") == etag:
            self.send_response(304)
            self.send_header("ETag", etag)
            self.common_headers()
            self.end_headers()
            return

        encoding = None
        if gz_path is not None and gz_path.exists() and self.accepts_gzip() and gz_path.stat().st_mtime >= stat.st_mtime - 1:
            data, encoding = gz_path.read_bytes(), "gzip"
        else:
            data = path.read_bytes()
            if self.accepts_gzip() and len(data) > 1024 and content_type.startswith(COMPRESSIBLE):
                data, encoding = gzip.compress(data, compresslevel=6), "gzip"

        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        if encoding:
            self.send_header("Content-Encoding", encoding)
            self.send_header("Vary", "Accept-Encoding")
        if no_cache:
            self.send_header("Cache-Control", "no-store")
        elif revalidate:
            self.send_header("Cache-Control", "no-cache")
            self.send_header("ETag", etag)
        if download_name:
            safe = re.sub(r'[^\w.\- ]', "_", download_name)
            self.send_header("Content-Disposition", f'attachment; filename="{safe}"')
        self.common_headers()
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        encoding = None
        if self.accepts_gzip() and len(data) > 1024:
            data, encoding = gzip.compress(data, compresslevel=5), "gzip"
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        if encoding:
            self.send_header("Content-Encoding", encoding)
            self.send_header("Vary", "Accept-Encoding")
        self.send_header("Cache-Control", "no-store")
        self.common_headers()
        self.end_headers()
        self.wfile.write(data)

    def send_text(self, text, status=200):
        data = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.common_headers()
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        print("%s - %s" % (self.address_string(), fmt % args), flush=True)


XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def ensure_processed_stock() -> None:
    """Reprocesa si no hay datos o si vienen de una versión anterior (sin dataVersion)."""
    if not STOCK_FILE.exists():
        return
    needs = not DATA_FILE.exists() or not DATA_FILE_GZ.exists()
    if not needs:
        return
    try:
        old_meta = read_json(META_FILE, {}) or {}
        name = old_meta.get("name", STOCK_FILE.name)
        meta = process_upload(STOCK_FILE, name, old_meta.get("uploadedAt"))
        # Evita duplicar en el historial la misma planilla reprocesada.
        if old_meta.get("id") and old_meta["id"] != meta["id"]:
            entries = [e for e in load_history_index() if e["id"] != old_meta["id"]]
            remove_history_files(old_meta["id"])
            write_json_atomic(HISTORY_INDEX_FILE, entries)
        print(f"Planilla existente procesada al formato v{VERSION}: {meta['name']}", flush=True)
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
    TELEGRAM_BOT = start_bot(DATA_FILE)
    if TELEGRAM_BOT is None:
        print("Bot de Telegram desactivado (definí TELEGRAM_BOT_TOKEN para activarlo).", flush=True)
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), StockHandler)
    print(f"Stock server v{VERSION} listening on {port}" + (" (carga protegida con código)" if UPLOAD_TOKEN else ""), flush=True)
    server.serve_forever()
