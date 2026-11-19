from __future__ import annotations

import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

BASE_COLUMNS = [
    {"base": "CARRITO", "col": "C"},
    {"base": "LOCAL", "col": "L"},
    {"base": "MAYORISTA", "col": "M"},
    {"base": "RETAIL", "col": "RE"},
    {"base": "RIO", "col": "RI"},
    {"base": "BARI", "col": "B"},
    {"base": "DEPOSITO", "col": "D"},
    {"base": "EXPORTA", "col": "E"},
    {"base": "CONTROL", "col": "CO"},
    {"base": "PENDIENTES", "col": "P.PAGO"},
]

import os

# El precio se toma ÚNICAMENTE de la columna de precio RETAIL.
# Ojo: la columna "RE" es el STOCK de la base RETAIL, no el precio; nunca se usa como precio.
# Si en la planilla el encabezado se llama distinto, se puede forzar con PRICE_COLUMN.
PRICE_CANDIDATES = ["RETAIL", "PRECIO RETAIL", "P.RETAIL", "P. RETAIL", "PRECIO RETAIL $", "RETAIL $"]
PRICE_SHEETS = ["PRECIOS", "LISTA DE PRECIOS", "LISTA PRECIOS"]

_NUMBER_CLEAN = re.compile(r"[^\d.\-]")
_SPACES = re.compile(r"\s+")


def normalize(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return _SPACES.sub(" ", text).strip().upper()


def as_number(value: Any) -> float:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value) if value == value else 0.0
    cleaned = _NUMBER_CLEAN.sub("", str(value or "").replace(",", "."))
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def clean_value(value: Any) -> Any:
    """Valores serializables y compactos (sin None, enteros sin .0)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def find_header_index(header: list[Any], name: str) -> int:
    wanted = normalize(name)
    for index, cell in enumerate(header):
        if normalize(cell) == wanted:
            return index
    return -1


def is_total_row(row: list[Any]) -> bool:
    return any(isinstance(cell, str) and normalize(cell) == "TOTAL" for cell in row)


def parse_price(value: Any) -> float | None:
    """Convierte '$ 45.990,50', '45990.5' o 45990 en número. Vacío -> None."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value) if value == value and value > 0 else None
    text = re.sub(r"[^\d.,\-]", "", str(value))
    if not text:
        return None
    if "," in text and "." in text:
        # el último separador es el decimal
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        head, _, tail = text.rpartition(",")
        text = head.replace(",", "") + ("." + tail if len(tail) != 3 else tail)
    elif text.count(".") > 1 or (text.count(".") == 1 and len(text.rpartition(".")[2]) == 3):
        text = text.replace(".", "")
    try:
        number = float(text)
    except ValueError:
        return None
    return number if number > 0 else None


def find_price_column(header: list[Any]) -> int:
    stock_cols = {normalize(base["col"]) for base in BASE_COLUMNS}
    forced = os.environ.get("PRICE_COLUMN", "").strip()
    if forced:
        return find_header_index(header, forced)
    for name in PRICE_CANDIDATES:
        index = find_header_index(header, name)
        if index >= 0:
            return index
    # Cualquier encabezado que contenga RETAIL (p.ej. "PRECIO RETAIL ARS"), nunca una columna de stock.
    for index, cell in enumerate(header):
        name = normalize(cell)
        if "RETAIL" in name and name not in stock_cols:
            return index
    return -1


def build_price_lookup(rows: list[list[Any]] | None) -> tuple[dict[str, float], dict[str, float]]:
    """Pestaña de precios opcional: por CONCAT y/o por SKU."""
    by_concat: dict[str, float] = {}
    by_sku: dict[str, float] = {}
    if not rows:
        return by_concat, by_sku
    for start, header in enumerate(rows[:10]):
        price_col = find_price_column(list(header))
        if price_col < 0:
            continue
        concat_col = find_header_index(list(header), "CONCAT")
        sku_col = find_header_index(list(header), "SKU")
        if concat_col < 0 and sku_col < 0:
            continue
        for row in rows[start + 1:]:
            if not row or price_col >= len(row):
                continue
            price = parse_price(row[price_col])
            if price is None:
                continue
            if 0 <= concat_col < len(row) and normalize(row[concat_col]):
                by_concat[normalize(row[concat_col])] = price
            if 0 <= sku_col < len(row) and normalize(clean_value(row[sku_col])):
                by_sku.setdefault(normalize(clean_value(row[sku_col])), price)
        break
    return by_concat, by_sku


def read_sheets(workbook_path: Path, names: list[str]) -> dict[str, list[list[Any]] | None]:
    """Abre el Excel una sola vez y devuelve las pestañas pedidas."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        by_name = {normalize(name): name for name in workbook.sheetnames}
        result: dict[str, list[list[Any]] | None] = {}
        for wanted in names:
            real = by_name.get(normalize(wanted))
            result[wanted] = (
                [list(row) for row in workbook[real].iter_rows(values_only=True)] if real else None
            )
        return result
    finally:
        workbook.close()


def detect_pending_columns(pending_rows: list[list[Any]]) -> tuple[int, int]:
    for row in pending_rows[:10]:
        qty_index = -1
        for index, cell in enumerate(row):
            normalized = normalize(cell)
            if "QUANTITY" in normalized or "CANT" in normalized:
                qty_index = index
                break
        if qty_index >= 0:
            concat_index = 0
            for index, cell in enumerate(row):
                if "CONCAT" in normalize(cell):
                    concat_index = index
                    break
            return qty_index, concat_index
    return -1, 0


def build_stock_payload(workbook_path: Path) -> dict[str, Any]:
    sheets = read_sheets(workbook_path, ["STOCK", "PENDIENTES", *PRICE_SHEETS])
    stock_rows = sheets["STOCK"]
    if stock_rows is None:
        raise ValueError("No encontré una pestaña llamada STOCK.")
    if not stock_rows:
        raise ValueError("La pestaña STOCK está vacía.")
    pending_rows = sheets["PENDIENTES"] or []

    header = list(stock_rows[0])
    concat_col = find_header_index(header, "CONCAT")
    sku_col = find_header_index(header, "SKU")
    color_col = find_header_index(header, "COLOR")
    talle_col = find_header_index(header, "TALLE")
    nombre_col = find_header_index(header, "NOMBRE")
    linea_col = find_header_index(header, "LINEA")
    # Índices de columnas de bases calculados una sola vez (antes era por fila).
    base_cols = [(base["base"], find_header_index(header, base["col"])) for base in BASE_COLUMNS]
    price_col = find_price_column(header)
    price_sheet = next((sheets[name] for name in PRICE_SHEETS if sheets.get(name)), None)
    prices_by_concat, prices_by_sku = build_price_lookup(price_sheet)

    pending_qty_col, pending_concat_col = detect_pending_columns(pending_rows)
    pending_by_concat: dict[str, float] = {}
    if pending_rows and pending_qty_col >= 0:
        for row in pending_rows[1:]:
            if not row:
                continue
            concat = normalize(row[pending_concat_col] if pending_concat_col < len(row) else "")
            if not concat:
                continue
            qty = as_number(row[pending_qty_col] if pending_qty_col < len(row) else 0)
            pending_by_concat[concat] = pending_by_concat.get(concat, 0.0) + qty

    index: dict[str, dict[str, Any]] = {}
    inventory: dict[str, dict[str, Any]] = {}
    data_rows: list[list[Any]] = []
    lineas: set[str] = set()
    total_row = None
    width = len(header)

    def cell(values: list[Any], col: int) -> Any:
        return clean_value(values[col]) if col >= 0 else ""

    for row in stock_rows[1:]:
        if is_total_row(row):
            total_row = list(row) + [""] * max(0, width - len(row))
            continue

        row_values = list(row) + [""] * max(0, width - len(row))
        concat_value = normalize(row_values[concat_col]) if concat_col >= 0 else ""
        built_value = normalize(
            " ".join(
                str(clean_value(row_values[col]))
                for col in (sku_col, color_col, talle_col)
                if col >= 0 and row_values[col] not in (None, "")
            )
        )
        key = concat_value or built_value
        if not key:
            continue

        bases = {name: (as_number(row_values[col]) if col >= 0 else 0.0) for name, col in base_cols}
        pending_qty = pending_by_concat.get(key, 0.0)
        if pending_qty > 0:
            bases["PENDIENTES"] = pending_qty

        product = {
            "sku": cell(row_values, sku_col),
            "color": cell(row_values, color_col),
            "talle": cell(row_values, talle_col),
            "nombre": cell(row_values, nombre_col),
            "linea": cell(row_values, linea_col),
            "precio": (
                (parse_price(row_values[price_col]) if price_col >= 0 else None)
                or prices_by_concat.get(key)
                or prices_by_sku.get(normalize(cell(row_values, sku_col)))
            ),
            "bases": bases,
        }
        linea = normalize(product["linea"])
        if linea:
            lineas.add(linea)

        data_rows.append([clean_value(value) for value in row_values[:width]])
        index[key] = product
        inventory[key] = {"sku": product["sku"], "nombre": product["nombre"], "bases": bases}

    totals: dict[str, float] = {name: 0.0 for name, _ in base_cols}
    for name, col in base_cols:
        if col < 0:
            continue
        if total_row is not None and total_row[col] not in (None, ""):
            totals[name] = as_number(total_row[col])
        else:
            totals[name] = sum(as_number(row[col]) for row in data_rows)
    if pending_by_concat:
        totals["PENDIENTES"] = sum(pending_by_concat.values())

    return {
        "header": [clean_value(value) for value in header],
        "rows": data_rows,
        "index": index,
        "inventory": inventory,
        "pendingByConcat": pending_by_concat,
        "totals": totals,
        "lineas": sorted(lineas),
        "productCount": len(data_rows),
        "priceColumn": header[price_col] if price_col >= 0 else ("pestaña de precios" if price_sheet else None),
    }


def make_history_id(uploaded_at: int | None = None) -> str:
    timestamp = uploaded_at or int(datetime.now(tz=timezone.utc).timestamp())
    return datetime.fromtimestamp(timestamp, tz=timezone.utc).strftime("%Y%m%d_%H%M%S")


def compare_inventories(
    current: dict[str, dict[str, Any]],
    previous: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    changes: list[dict[str, Any]] = []

    def add(key, item, base, before, after, kind):
        changes.append(
            {
                "concat": key,
                "sku": item.get("sku", ""),
                "nombre": item.get("nombre", ""),
                "base": base,
                "before": before,
                "after": after,
                "delta": after - before,
                "type": kind,
            }
        )

    for key in sorted(set(current) | set(previous)):
        current_item = current.get(key)
        previous_item = previous.get(key)

        if current_item and not previous_item:
            for base, qty in current_item["bases"].items():
                if qty > 0:
                    add(key, current_item, base, 0.0, float(qty), "new")
            continue

        if previous_item and not current_item:
            for base, qty in previous_item["bases"].items():
                if qty > 0:
                    add(key, previous_item, base, float(qty), 0.0, "removed")
            continue

        current_bases = current_item["bases"]
        previous_bases = previous_item["bases"]
        for base in sorted(set(current_bases) | set(previous_bases)):
            before = float(previous_bases.get(base, 0.0))
            after = float(current_bases.get(base, 0.0))
            if before != after:
                add(key, current_item, base, before, after, "changed")

    by_base: dict[str, dict[str, float]] = {}
    for item in changes:
        entry = by_base.setdefault(item["base"], {"in": 0.0, "out": 0.0, "net": 0.0, "lines": 0})
        entry["lines"] += 1
        entry["net"] += item["delta"]
        if item["delta"] > 0:
            entry["in"] += item["delta"]
        else:
            entry["out"] += item["delta"]

    new_keys = set(current) - set(previous)
    removed_keys = set(previous) - set(current)
    summary = {
        "newProducts": len(new_keys),
        "removedProducts": len(removed_keys),
        "changedProducts": len({item["concat"] for item in changes}),
        "changedLines": len(changes),
        "positive": sum(1 for item in changes if item["delta"] > 0),
        "negative": sum(1 for item in changes if item["delta"] < 0),
        "unitsIn": sum(item["delta"] for item in changes if item["delta"] > 0),
        "unitsOut": sum(item["delta"] for item in changes if item["delta"] < 0),
        "byBase": by_base,
    }
    return {"changes": changes, "summary": summary}
