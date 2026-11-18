const BASE_COLUMNS = [
  { base: "CARRITO", col: "C" },
  { base: "LOCAL", col: "L" },
  { base: "MAYORISTA", col: "M" },
  { base: "RETAIL", col: "RE" },
  { base: "RIO", col: "RI" },
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
  return row.some((cell) => typeof cell === "string" && normalize(cell) === "TOTAL");
}

// Nota: DOMParser NO existe dentro de un Web Worker, por eso el XML se lee con expresiones regulares.
function decodeXml(text) {
  return String(text ?? "").replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_, code) => {
    if (code === "lt") return "<";
    if (code === "gt") return ">";
    if (code === "amp") return "&";
    if (code === "quot") return '"';
    if (code === "apos") return "'";
    if (code[1] === "x") return String.fromCodePoint(parseInt(code.slice(2), 16));
    return String.fromCodePoint(parseInt(code.slice(1), 10));
  });
}

function attr(attrs, name) {
  const match = new RegExp(`(?:^|\\s)${name.replace(":", "\\:")}="([^"]*)"`).exec(attrs);
  return match ? decodeXml(match[1]) : null;
}

function joinTextNodes(xml) {
  let out = "";
  const re = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g;
  let m;
  while ((m = re.exec(xml))) out += m[1];
  return decodeXml(out);
}

function parseWorkbook(zip) {
  const workbookXml = getXml(zip, "xl/workbook.xml");
  const relsXml = getXml(zip, "xl/_rels/workbook.xml.rels");
  const sharedStrings = zip["xl/sharedStrings.xml"]
    ? parseSharedStrings(getXml(zip, "xl/sharedStrings.xml"))
    : [];

  const rels = {};
  for (const m of relsXml.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)) {
    const target = attr(m[1], "Target") || "";
    rels[attr(m[1], "Id")] = normalizePath(target.startsWith("/") ? target.slice(1) : "xl/" + target);
  }

  const sheets = [];
  for (const m of workbookXml.matchAll(/<(?:\w+:)?sheet\b([^>]*)>/g)) {
    const relId = attr(m[1], "r:id") || attr(m[1], "id");
    sheets.push({ name: attr(m[1], "name"), path: rels[relId] });
  }
  return { sheets, sharedStrings };
}

function parseSharedStrings(xml) {
  const out = [];
  // Ignora las guías fonéticas (<rPh>) para no duplicar texto.
  const clean = xml.replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, "");
  for (const m of clean.matchAll(/<(?:\w+:)?si\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?si>)/g)) {
    out.push(m[1] ? joinTextNodes(m[1]) : "");
  }
  return out;
}

function parseSheet(zip, path, sharedStrings) {
  const xml = getXml(zip, path);
  const rows = [];
  const rowRe = /<(?:\w+:)?row\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g;
  const cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
  let rowMatch;
  while ((rowMatch = rowRe.exec(xml))) {
    const row = [];
    const body = rowMatch[1] || "";
    let cellMatch;
    cellRe.lastIndex = 0;
    while ((cellMatch = cellRe.exec(body))) {
      const attrs = cellMatch[1];
      const ref = attr(attrs, "r") || "";
      const index = ref ? columnIndex(ref.replace(/\d/g, "")) : row.length;
      row[index] = readCell(attr(attrs, "t"), cellMatch[2] || "", sharedStrings);
    }
    for (let i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = "";
    rows.push(row);
  }
  return rows;
}

function readCell(type, inner, sharedStrings) {
  if (type === "inlineStr") return joinTextNodes(inner);
  const valueMatch = /<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/.exec(inner);
  if (!valueMatch) return "";
  const raw = decodeXml(valueMatch[1]);
  if (type === "s") return sharedStrings[Number(raw)] ?? "";
  if (type === "b") return raw === "1";
  if (type === "str" || type === "e") return raw;
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
  return text;
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

function readZipCatalog(bytes, view) {
  const catalog = new Map();
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
    if (!name.endsWith("/")) {
      catalog.set(name, { compression, dataStart, compressedSize });
    }
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return catalog;
}

async function readZipEntry(bytes, meta) {
  const compressed = bytes.slice(meta.dataStart, meta.dataStart + meta.compressedSize);
  if (meta.compression === 0) return compressed;
  if (meta.compression === 8) return await inflateRaw(compressed);
  throw new Error(`Compresión no soportada (${meta.compression}).`);
}

async function unzipXlsx(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const catalog = readZipCatalog(bytes, view);
  const entries = {};

  async function loadText(path) {
    const name = normalizePath(path);
    if (entries[name]) return entries[name];
    const meta = catalog.get(name);
    if (!meta) throw new Error(`Falta ${name} dentro del Excel.`);
    const content = await readZipEntry(bytes, meta);
    entries[name] = new TextDecoder("utf-8").decode(content);
    return entries[name];
  }

  await loadText("xl/workbook.xml");
  await loadText("xl/_rels/workbook.xml.rels");

  const workbook = parseWorkbook(entries);
  const needed = new Set(["xl/workbook.xml", "xl/_rels/workbook.xml.rels"]);
  if (catalog.has("xl/sharedStrings.xml")) needed.add("xl/sharedStrings.xml");

  const stockSheet = workbook.sheets.find((sheet) => normalize(sheet.name) === "STOCK");
  if (!stockSheet) throw new Error("No encontré una pestaña llamada STOCK.");
  needed.add(stockSheet.path);

  const pendingSheet = workbook.sheets.find((sheet) => normalize(sheet.name) === "PENDIENTES");
  if (pendingSheet) needed.add(pendingSheet.path);

  for (const path of needed) await loadText(path);
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
  const baseCols = BASE_COLUMNS.map((base) => ({ base: base.base, colIndex: findHeaderIndex(header, base.col) }));

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
  const rows = [];
  const lineas = new Set();

  for (let i = 1; i < stockRows.length; i++) {
    const row = stockRows[i];
    if (isTotalRow(row)) continue;
    while (row.length < header.length) row.push("");

    const concatValue = concatCol >= 0 ? normalize(row[concatCol]) : "";
    const builtValue = normalize(
      [row[skuCol], row[colorCol], row[talleCol]].filter(Boolean).join(" ")
    );
    const key = concatValue || builtValue;
    if (!key) continue;

    const bases = {};
    for (const { base, colIndex } of baseCols) {
      bases[base] = colIndex >= 0 ? asNumber(row[colIndex]) : 0;
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
    if (totalRow && totalRow[colIndex] !== undefined && totalRow[colIndex] !== "") totals[base.base] = asNumber(totalRow[colIndex]);
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
    pendingByConcat,
    totals,
    lineas: Array.from(lineas).sort(),
    productCount: rows.length
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
