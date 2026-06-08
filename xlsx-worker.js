const BASE_COLUMNS = [
  { base: "CARRITO", col: "C" },
  { base: "LOCAL", col: "L" },
  { base: "MAYORISTA", col: "M" },
  { base: "RETAIL", col: "R" },
  { base: "BARI", col: "B" },
  { base: "DEPOSITO", col: "D" },
  { base: "EXPORTA", col: "E" },
  { base: "CONTROL", col: "CO" },
  { base: "PENDIENTES", col: "P.PAGO" }
];

function normalize(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function asNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value ?? "").replace(",", ".").replace(/[^\d.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function findHeaderIndex(header, name) {
  const wanted = normalize(name);
  return header.findIndex((cell) => normalize(cell) === wanted);
}

function isTotalRow(row) {
  return row.some((cell) => normalize(cell) === "TOTAL");
}

function parseWorkbook(zip) {
  const workbookXml = getXml(zip, "xl/workbook.xml");
  const relsXml = getXml(zip, "xl/_rels/workbook.xml.rels");
  const sharedStrings = zip["xl/sharedStrings.xml"]
    ? parseSharedStrings(getXml(zip, "xl/sharedStrings.xml"))
    : [];

  const rels = {};
  for (const rel of relsXml.getElementsByTagName("Relationship")) {
    const target = rel.getAttribute("Target") || "";
    rels[rel.getAttribute("Id")] = normalizePath(
      target.startsWith("/") ? target.slice(1) : "xl/" + target
    );
  }

  const sheets = [];
  for (const sheet of workbookXml.getElementsByTagName("sheet")) {
    const relId = sheet.getAttribute("r:id") || sheet.getAttribute("id");
    sheets.push({ name: sheet.getAttribute("name"), path: rels[relId] });
  }

  return { sheets, sharedStrings };
}

function parseSharedStrings(xml) {
  return Array.from(xml.getElementsByTagName("si")).map((si) =>
    Array.from(si.getElementsByTagName("t"))
      .map((node) => node.textContent || "")
      .join("")
  );
}

function parseSheet(zip, path, sharedStrings) {
  const xml = getXml(zip, path);
  const rows = [];
  for (const rowNode of xml.getElementsByTagName("row")) {
    const row = [];
    for (const cell of rowNode.getElementsByTagName("c")) {
      const ref = cell.getAttribute("r") || "";
      const index = ref ? columnIndex(ref.replace(/\d/g, "")) : row.length;
      row[index] = readCell(cell, sharedStrings);
    }
    rows.push(row);
  }
  return rows;
}

function readCell(cell, sharedStrings) {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") {
    return Array.from(cell.getElementsByTagName("t"))
      .map((node) => node.textContent || "")
      .join("");
  }
  const valueNode = cell.getElementsByTagName("v")[0];
  if (!valueNode) return "";
  const raw = valueNode.textContent || "";
  if (type === "s") return sharedStrings[Number(raw)] ?? "";
  if (type === "b") return raw === "1";
  if (type === "str") return raw;
  const numeric = Number(raw);
  return Number.isFinite(numeric) ? numeric : raw;
}

function columnIndex(letters) {
  let total = 0;
  for (const char of letters.toUpperCase()) {
    total = total * 26 + char.charCodeAt(0) - 64;
  }
  return total - 1;
}

function getXml(zip, path) {
  const text = zip[normalizePath(path)];
  if (!text) throw new Error(`Falta ${path} dentro del Excel.`);
  return new DOMParser().parseFromString(text, "application/xml");
}

function normalizePath(path) {
  const parts = [];
  for (const part of path.replace(/\\/g, "/").split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
}

async function unzipXlsx(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const entries = {};
  const eocdOffset = findEndOfCentralDirectory(view);
  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  let offset = centralDirectoryOffset;
  const end = centralDirectoryOffset + centralDirectorySize;

  while (offset < end) {
    if (view.getUint32(offset, true) !== 0x02014b50) break;
    const compression = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const nameBytes = bytes.slice(offset + 46, offset + 46 + fileNameLength);
    const name = normalizePath(new TextDecoder().decode(nameBytes));
    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);

    if (!name.endsWith("/")) {
      let content;
      if (compression === 0) content = compressed;
      else if (compression === 8) content = await inflateRaw(compressed);
      else throw new Error(`Compresión no soportada en ${name}.`);
      if (name.endsWith(".xml") || name.endsWith(".rels")) {
        entries[name] = new TextDecoder("utf-8").decode(content);
      }
    }
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function findEndOfCentralDirectory(view) {
  for (let offset = view.byteLength - 22; offset >= Math.max(0, view.byteLength - 70000); offset--) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }
  throw new Error("No pude abrir el ZIP interno del .xlsx.");
}

async function inflateRaw(bytes) {
  try {
    return await inflateWithFormat(bytes, "deflate-raw");
  } catch {
    return await inflateWithFormat(bytes, "deflate");
  }
}

async function inflateWithFormat(bytes, format) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(format));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function detectPendingColumns(pendingRows) {
  for (const row of pendingRows.slice(0, 10)) {
    const qtyIndex = row.findIndex(
      (cell) => normalize(cell).includes("QUANTITY") || normalize(cell).includes("CANT")
    );
    if (qtyIndex >= 0) {
      const concatIndex = row.findIndex((cell) => normalize(cell).includes("CONCAT"));
      return { pendingQuantityCol: qtyIndex, pendingConcatCol: concatIndex >= 0 ? concatIndex : 0 };
    }
  }
  return { pendingQuantityCol: -1, pendingConcatCol: 0 };
}

function buildPayload(stockRows, pendingRows, pendingQuantityCol, pendingConcatCol) {
  const header = stockRows[0] || [];
  const concatCol = findHeaderIndex(header, "CONCAT");
  const skuCol = findHeaderIndex(header, "SKU");
  const colorCol = findHeaderIndex(header, "COLOR");
  const talleCol = findHeaderIndex(header, "TALLE");
  const nombreCol = findHeaderIndex(header, "NOMBRE");
  const lineaCol = findHeaderIndex(header, "LINEA");

  const pendingByConcat = {};
  if (pendingRows.length && pendingQuantityCol >= 0) {
    for (let i = 1; i < pendingRows.length; i++) {
      const row = pendingRows[i];
      const concat = normalize(row[pendingConcatCol]);
      if (!concat) continue;
      pendingByConcat[concat] = (pendingByConcat[concat] || 0) + asNumber(row[pendingQuantityCol]);
    }
  }

  const index = {};
  const inventory = {};
  const rows = [];
  const lineas = new Set();

  for (let i = 1; i < stockRows.length; i++) {
    const row = stockRows[i];
    if (isTotalRow(row)) continue;

    const concatValue = concatCol >= 0 ? normalize(row[concatCol]) : "";
    const builtValue = normalize(
      [row[skuCol], row[colorCol], row[talleCol]].filter(Boolean).join(" ")
    );
    const key = concatValue || builtValue;
    if (!key) continue;

    const bases = {};
    for (const base of BASE_COLUMNS) {
      const colIndex = findHeaderIndex(header, base.col);
      bases[base.base] = colIndex >= 0 ? asNumber(row[colIndex]) : 0;
    }
    if (pendingByConcat[key] > 0) bases.PENDIENTES = pendingByConcat[key];

    const product = {
      sku: row[skuCol] ?? "",
      color: row[colorCol] ?? "",
      talle: row[talleCol] ?? "",
      nombre: row[nombreCol] ?? "",
      linea: row[lineaCol] ?? "",
      bases
    };

    if (normalize(product.linea)) lineas.add(normalize(product.linea));
    rows.push(row);
    index[key] = product;
    inventory[key] = { ...product, concat: key, bases, row };
  }

  const totals = {};
  for (const base of BASE_COLUMNS) totals[base.base] = 0;

  let totalRow = null;
  for (let i = stockRows.length - 1; i >= 1; i--) {
    if (isTotalRow(stockRows[i])) {
      totalRow = stockRows[i];
      break;
    }
  }

  for (const base of BASE_COLUMNS) {
    const colIndex = findHeaderIndex(header, base.col);
    if (colIndex < 0) continue;
    if (totalRow) totals[base.base] = asNumber(totalRow[colIndex]);
    else {
      for (const row of rows) totals[base.base] += asNumber(row[colIndex]);
    }
  }
  if (Object.keys(pendingByConcat).length) {
    totals.PENDIENTES = Object.values(pendingByConcat).reduce((sum, qty) => sum + qty, 0);
  }

  return {
    header,
    rows,
    index,
    inventory,
    pendingByConcat,
    totals,
    lineas: Array.from(lineas).sort(),
    productCount: rows.length,
    pendingRows,
    pendingQuantityCol,
    pendingConcatCol
  };
}

self.onmessage = async (event) => {
  try {
    const zip = await unzipXlsx(event.data.buffer);
    const workbook = parseWorkbook(zip);
    const stockPath = workbook.sheets.find((sheet) => normalize(sheet.name) === "STOCK")?.path;
    if (!stockPath) throw new Error("No encontré una pestaña llamada STOCK.");

    const stockRows = parseSheet(zip, stockPath, workbook.sharedStrings);
    const pendingSheet = workbook.sheets.find((sheet) => normalize(sheet.name) === "PENDIENTES");
    const pendingRows = pendingSheet
      ? parseSheet(zip, pendingSheet.path, workbook.sharedStrings)
      : [];
    const pendingCols = detectPendingColumns(pendingRows);
    const payload = buildPayload(stockRows, pendingRows, pendingCols.pendingQuantityCol, pendingCols.pendingConcatCol);
    self.postMessage({ ok: true, payload });
  } catch (error) {
    self.postMessage({ ok: false, error: error.message || String(error) });
  }
};
