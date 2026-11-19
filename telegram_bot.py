"""Bot de Telegram para consultar stock en la base LOCAL.

Corre como un hilo dentro de server.py (mismo contenedor). Usa long polling,
así que no necesita abrir puertos ni tener dominio/HTTPS.

Variables de entorno:
  TELEGRAM_BOT_TOKEN      token que da @BotFather (obligatorio para activarlo)
  TELEGRAM_ALLOWED_CHATS  ids de chat permitidos separados por coma (recomendado)
  TELEGRAM_BASE           base a consultar (default LOCAL)
  TELEGRAM_MAX_ITEMS      máximo de variantes por respuesta (default 40)
"""
from __future__ import annotations

import html
import json
import os
import re
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Callable

from stock_processor import normalize

def env(name: str, default: str = "") -> str:
    """Lee una variable y limpia espacios y comillas (Portainer las deja tal cual si se pegan)."""
    value = os.environ.get(name, default) or ""
    return value.strip().strip("\"'").strip()


API_BASE = env("TELEGRAM_API_BASE", "https://api.telegram.org").rstrip("/")
BASE = normalize(env("TELEGRAM_BASE", "LOCAL")) or "LOCAL"
MAX_ITEMS = int(env("TELEGRAM_MAX_ITEMS", "40") or 40)

HELP_TEXT = (
    "📦 <b>Consulta de stock · {base}</b>\n\n"
    "Mandame un código y te digo qué hay en {base}:\n"
    "• <code>25551</code> → todas las variantes del SKU\n"
    "• <code>25551 NLSA</code> → un color\n"
    "• <code>25551 NLSA M</code> → color y talle\n\n"
    "Te respondo código, color, talle, cantidad y precio. "
    "Si no hay unidades en {base}, te aviso que no hay stock."
)


_NOT_ALNUM = re.compile(r"[^0-9A-Z]")


def compact(value: Any) -> str:
    return _NOT_ALNUM.sub("", normalize(value))


def fmt_number(value: float) -> str:
    if value == int(value):
        return f"{int(value):,}".replace(",", ".")
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_price(value: Any) -> str:
    if value in (None, "", 0):
        return "sin precio"
    return "$ " + fmt_number(float(value))


def sort_key(item: tuple[str, dict]) -> tuple:
    key, product = item
    talles = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "2XL", "3XL"]
    talle = normalize(product.get("talle"))
    talle_rank = talles.index(talle) if talle in talles else (100 + (int(talle) if talle.isdigit() else 0))
    return (normalize(product.get("sku")), normalize(product.get("color")), talle_rank, key)


class StockLookup:
    """Lee el JSON procesado del servidor y lo recarga cuando cambia la planilla."""

    def __init__(self, data_file: Path):
        self.data_file = data_file
        self._mtime = None
        self._index: dict[str, dict] = {}
        self._lock = threading.Lock()

    def index(self) -> dict[str, dict]:
        try:
            mtime = self.data_file.stat().st_mtime
        except FileNotFoundError:
            return {}
        with self._lock:
            if mtime != self._mtime:
                data = json.loads(self.data_file.read_text(encoding="utf-8"))
                self._index = data.get("index", {})
                self._mtime = mtime
            return self._index

    def find(self, query: str) -> list[tuple[str, dict]]:
        """Búsqueda tolerante (igual de flexible que la web).

        Ignora espacios, guiones, puntos y barras, así "25551NLSAM", "25551 NLSA M"
        y "25551-NLSA-M" dan lo mismo.
        """
        q = normalize(query)
        qc = compact(q)
        if not qc:
            return []
        index = self.index()
        if q in index:
            return [(q, index[q])]
        entries = [(k, v, compact(k), compact(v.get("sku"))) for k, v in index.items()]
        exact = [(k, v) for k, v, kc, _ in entries if kc == qc]
        if len(exact) == 1:
            return exact
        # SKU completo -> todas sus variantes
        tiers = [
            [(k, v) for k, v, _, sc in entries if sc and sc == qc],
            # SKU+color, SKU+color+talle parcial, etc.
            [(k, v) for k, v, kc, _ in entries if kc.startswith(qc)],
            [(k, v) for k, v, kc, sc in entries if qc in kc or (sc and qc in sc)],
        ]
        for tier in tiers:
            if tier:
                return sorted(tier, key=sort_key)
        # Último recurso: todas las palabras presentes (en cualquier orden)
        tokens = q.split(" ")
        loose = [
            (k, v) for k, v in index.items()
            if all(t in f"{k} {normalize(v.get('sku'))} {normalize(v.get('nombre'))} {normalize(v.get('color'))}" for t in tokens)
        ]
        return sorted(loose, key=sort_key)

    def answer(self, query: str) -> str:
        if not self.index():
            return "⚠️ Todavía no hay una planilla cargada en el sistema."
        shown = html.escape(query.strip().upper())
        matches = self.find(query)
        if not matches:
            return f"❓ <b>{shown}</b>: ese código no existe en la planilla."

        in_stock = [(k, v) for k, v in matches if float(v.get("bases", {}).get(BASE, 0) or 0) > 0]
        if not in_stock:
            return f"❌ <b>{shown}</b>: no hay stock en {BASE}."

        lines: list[str] = []
        current_sku = None
        for key, product in in_stock[:MAX_ITEMS]:
            sku = str(product.get("sku") or "")
            if sku != current_sku:
                current_sku = sku
                if lines:
                    lines.append("")
                nombre = html.escape(str(product.get("nombre") or ""))
                lines.append(f"📦 <b>{html.escape(sku or key)}</b>" + (f" · {nombre}" if nombre else ""))
            qty = float(product["bases"][BASE])
            color = html.escape(str(product.get("color") or "-"))
            talle = html.escape(str(product.get("talle") or "-"))
            lines.append(
                f"• <code>{html.escape(key)}</code>\n"
                f"   Color <b>{color}</b> · Talle <b>{talle}</b> · {fmt_number(qty)} u · <b>{fmt_price(product.get('precio'))}</b>"
            )
        total = sum(float(v["bases"][BASE]) for _, v in in_stock)
        footer = f"\n\n✅ {len(in_stock)} variante{'s' if len(in_stock) != 1 else ''} con stock en {BASE} · {fmt_number(total)} u"
        if len(in_stock) > MAX_ITEMS:
            footer += f"\n(mostrando {MAX_ITEMS}; escribí también el color para acotar)"
        return "\n".join(lines) + footer


class TelegramBot:
    def __init__(self, token: str, lookup: StockLookup, allowed: set[int], log: Callable[[str], None] = print):
        self.token = token
        self.lookup = lookup
        self.allowed = allowed
        self.log = log
        self.offset = 0
        self.status: dict[str, Any] = {
            "enabled": True, "connected": False, "username": None, "lastError": None,
            "base": BASE, "allowedChats": len(allowed), "lastMessageAt": None,
        }

    # --- API ---
    def call(self, method: str, payload: dict | None = None, timeout: float = 30) -> Any:
        data = json.dumps(payload or {}).encode("utf-8")
        request = urllib.request.Request(
            f"{API_BASE}/bot{self.token}/{method}", data=data, headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = json.loads(response.read().decode("utf-8"))
        if not body.get("ok"):
            raise RuntimeError(body.get("description", "error de Telegram"))
        return body.get("result")

    def send(self, chat_id: int, text: str, reply_to: int | None = None) -> None:
        payload = {"chat_id": chat_id, "text": text, "parse_mode": "HTML", "disable_web_page_preview": True}
        if reply_to:
            payload["reply_to_message_id"] = reply_to
            payload["allow_sending_without_reply"] = True
        self.call("sendMessage", payload)

    # --- lógica ---
    def handle(self, message: dict) -> None:
        chat_id = message["chat"]["id"]
        text = (message.get("text") or "").strip()
        if not text:
            return
        command = text.split()[0].split("@")[0].lower() if text.startswith("/") else ""

        if command == "/id":
            self.send(chat_id, f"🆔 El id de este chat es <code>{chat_id}</code>")
            return
        if self.allowed and chat_id not in self.allowed:
            self.log(f"[telegram] chat no autorizado {chat_id} ({message['chat'].get('username') or message['chat'].get('title') or ''})")
            self.send(chat_id, f"⛔ Este chat no está autorizado.\nPasale este id al administrador: <code>{chat_id}</code>")
            return
        if command in ("/start", "/ayuda", "/help"):
            self.send(chat_id, HELP_TEXT.format(base=BASE))
            return
        if command in ("/stock", "/s", "/buscar"):
            text = text.split(maxsplit=1)[1] if len(text.split(maxsplit=1)) > 1 else ""
            if not text:
                self.send(chat_id, "Escribí el código después del comando, ej: <code>/stock 25551 NLSA</code>")
                return
        elif command:
            self.send(chat_id, "No conozco ese comando. Mandame directamente el código o /ayuda.")
            return

        self.status["lastMessageAt"] = int(time.time())
        answer = self.lookup.answer(text)
        self.log(f"[telegram] consulta '{text}' de {chat_id} -> {answer.splitlines()[0][:80]}")
        self.send(chat_id, answer, reply_to=message.get("message_id"))

    def run(self) -> None:
        backoff = 2
        while True:
            try:
                if not self.status["connected"]:
                    me = self.call("getMe")
                    self.call("deleteWebhook", {"drop_pending_updates": False})
                    try:
                        self.call("setMyCommands", {"commands": [
                            {"command": "stock", "description": f"Consultar stock en {BASE}"},
                            {"command": "ayuda", "description": "Cómo usar el bot"},
                            {"command": "id", "description": "Ver el id de este chat"},
                        ]})
                    except Exception:
                        pass
                    self.status.update(connected=True, username=me.get("username"), lastError=None)
                    self.log(f"[telegram] bot @{me.get('username')} conectado · base {BASE}"
                             + (f" · {len(self.allowed)} chat(s) autorizados" if self.allowed else " · SIN lista de chats (cualquiera puede consultar)"))
                updates = self.call("getUpdates", {"offset": self.offset, "timeout": 50, "allowed_updates": ["message"]}, timeout=65)
                backoff = 2
                for update in updates:
                    self.offset = update["update_id"] + 1
                    message = update.get("message")
                    if not message:
                        continue
                    try:
                        self.handle(message)
                    except Exception as error:  # un mensaje malo no tumba el bot
                        self.log(f"[telegram] error respondiendo: {error}")
            except urllib.error.HTTPError as error:
                detail = error.read().decode("utf-8", "ignore")[:200]
                self.status.update(connected=False, lastError=f"HTTP {error.code}")
                if error.code == 401:
                    self.log("[telegram] token inválido: revisá TELEGRAM_BOT_TOKEN. Bot detenido.")
                    self.status["enabled"] = False
                    return
                if error.code == 409:
                    self.log("[telegram] otro proceso está usando este bot (¿dos contenedores con el mismo token?).")
                else:
                    self.log(f"[telegram] HTTP {error.code}: {detail}")
                time.sleep(backoff)
                backoff = min(backoff * 2, 60)
            except Exception as error:
                self.status.update(connected=False, lastError=str(error)[:200])
                self.log(f"[telegram] sin conexión ({error}); reintento en {backoff}s")
                time.sleep(backoff)
                backoff = min(backoff * 2, 60)


def start_bot(data_file: Path) -> TelegramBot | None:
    token = env("TELEGRAM_BOT_TOKEN")
    if not token:
        return None
    allowed = set()
    for part in re.split(r"[,; \n]+", env("TELEGRAM_ALLOWED_CHATS")):
        part = part.strip().strip("\"'")
        if part.lstrip("-").isdigit():
            allowed.add(int(part))
    if ":" not in token:
        print("[telegram] OJO: el token no tiene el formato 123456789:AAxxxx. Revisá TELEGRAM_BOT_TOKEN.", flush=True)
    bot = TelegramBot(token, StockLookup(data_file), allowed, log=lambda m: print(m, flush=True))
    threading.Thread(target=bot.run, name="telegram-bot", daemon=True).start()
    return bot
