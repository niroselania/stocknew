window.__stockAppLoaded = true;

/* =========================================================
   Configuración
   ========================================================= */
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
const PENDING_BASE = "PENDIENTES";
const ROW_H = 34;
const OVERSCAN = 10;
const SUGGEST_LIMIT = 8;
const RECENT_MAX = 8;
const COMPARE_PAGE = 400;

const state = {
  header: [],
  rows: [],
  index: {},
  totals: {},
  lineas: [],
  meta: null,
  history: [],
  serverMode: location.protocol === "http:" || location.protocol === "https:",
  serverVersion: null,
  cols: { concat: -1, sku: -1, nombre: -1, linea: -1, bases: [] },
  rowInfo: [],
  searchList: [],
  bySku: new Map(),
  baseStats: {},
  view: [],
  sort: { col: -1, dir: 1 },
  filter: { text: "", linea: "", base: "", mode: "", negBase: "" },
  lastStart: -1,
  current: null,
  suggestions: [],
  activeSuggestion: -1,
  comparison: null,
  compareLimit: COMPARE_PAGE,
  compareFilter: { text: "", base: "", type: "" },
  pendingQuery: new URLSearchParams(location.search).get("q")
};

const $ = (id) => document.getElementById(id);
const el = {
  fileInput: $("fileInput"), uploadBtn: $("uploadBtn"), searchInput: $("searchInput"), suggest: $("suggest"),
  scanBtn: $("scanBtn"), themeBtn: $("themeBtn"), exportCopyBtn: $("exportCopyBtn"), exportCsvBtn: $("exportCsvBtn"),
  shareBtn: $("shareBtn"), compareBtn: $("compareBtn"), exportCompareBtn: $("exportCompareBtn"), status: $("status"),
  progress: $("progress"), clock: $("clock"), fileName: $("fileName"), updatedAt: $("updatedAt"), serverInfo: $("serverInfo"),
  serverLed: $("serverLed"), dataLed: $("dataLed"), downloadCurrent: $("downloadCurrent"), recent: $("recent"),
  kpiProducts: $("kpiProducts"), kpiUnits: $("kpiUnits"), kpiPending: $("kpiPending"), kpiBases: $("kpiBases"),
  kpiBasesSub: $("kpiBasesSub"), kpiNeg: $("kpiNeg"), kpiNegBtn: $("kpiNegBtn"),
  resultBox: $("resultBox"), queryLabel: $("queryLabel"), totalsBox: $("totalsBox"), qualityBox: $("qualityBox"),
  stockSection: $("stockSection"), stockTable: $("stockTable"), stockInfo: $("stockInfo"), stockFoot: $("stockFoot"),
  stockFilterText: $("stockFilterText"), stockFilterLinea: $("stockFilterLinea"), stockFilterBase: $("stockFilterBase"),
  stockFilterMode: $("stockFilterMode"), activeFilter: $("activeFilter"), clearFiltersBtn: $("clearFiltersBtn"),
  exportStockBtn: $("exportStockBtn"), historyBox: $("historyBox"), historyInfo: $("historyInfo"),
  compareBox: $("compareBox"), compareLabel: $("compareLabel"), scannerModal: $("scannerModal"),
  scannerVideo: $("scannerVideo"), scannerClose: $("scannerClose"), dropOverlay: $("dropOverlay")
};

let xlsxWorker = null;
let scannerStream = null;
let scannerLoop = null;
let suggestTimer = null;
let filterTimer = null;
let scrollRaf = 0;

/* =========================================================
   Utilidades
   ========================================================= */
function normalize(value) {
  return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toUpperCase();
}

function asNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value ?? "").replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

const numberFmt = new Intl.NumberFormat("es-AR", { maximumFractionDigits: 2 });
const fmt = (n) => numberFmt.format(n || 0);
const signed = (n) => (n > 0 ? "+" : "") + fmt(n);

function formatDate(ts) {
  if (!ts) return "Sin fecha";
  return new Intl.DateTimeFormat("es-AR", { dateStyle: "short", timeStyle: "short" }).format(new Date(ts * 1000));
}

function timeAgo(ts) {
  if (!ts) return "";
  const s = Math.max(0, Date.now() / 1000 - ts);
  if (s < 60) return "hace instantes";
  if (s < 3600) return `hace ${Math.round(s / 60)} min`;
  if (s < 86400) return `hace ${Math.round(s / 3600)} h`;
  const d = Math.round(s / 86400);
  return `hace ${d} día${d === 1 ? "" : "s"}`;
}

function formatBytes(bytes) {
  if (!bytes) return "";
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function findHeaderIndex(header, name) {
  const wanted = normalize(name);
  return header.findIndex((cell) => normalize(cell) === wanted);
}

function icon(name, extra = "") {
  return `<svg class="i" ${extra}><use href="#i-${name}"/></svg>`;
}

function setStatus(text, kind = "") {
  el.status.textContent = text;
  el.status.className = kind;
  el.status.title = text;
}

function setProgress(value) {
  // value: null = ocultar, "busy" = indeterminado, número 0-100
  if (value === null) { el.progress.classList.add("hidden"); return; }
  el.progress.classList.remove("hidden");
  el.progress.classList.toggle("indeterminate", value === "busy");
  el.progress.firstElementChild.style.width = value === "busy" ? "" : `${value}%`;
}

function countUp(node, target) {
  const start = performance.now();
  const from = 0;
  const duration = 650;
  const step = (now) => {
    const t = Math.min(1, (now - start) / duration);
    const eased = 1 - Math.pow(1 - t, 3);
    node.textContent = fmt(Math.round(from + (target - from) * eased));
    if (t < 1) requestAnimationFrame(step);
    else node.textContent = fmt(target);
  };
  requestAnimationFrame(step);
}

function downloadCsv(filename, rows) {
  const csv = rows.map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(";")).join("\r\n");
  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function emptyBox(text, iconName = "box") {
  return `<div class="empty">${icon(iconName)}<div>${text}</div></div>`;
}

function productTotal(bases) {
  let total = 0;
  for (const [base, qty] of Object.entries(bases || {})) {
    if (base !== PENDING_BASE && qty > 0) total += qty;
  }
  return total;
}

function getToken() { return localStorage.getItem("stock-upload-token") || ""; }

/* =========================================================
   Carga de datos
   ========================================================= */
function getWorker() {
  if (!xlsxWorker) xlsxWorker = new Worker("/xlsx-worker.js");
  return xlsxWorker;
}

function parseWithWorker(buffer) {
  return new Promise((resolve, reject) => {
    const worker = getWorker();
    const handler = (event) => {
      worker.removeEventListener("message", handler);
      if (event.data.ok) resolve(event.data.payload);
      else reject(new Error(event.data.error));
    };
    worker.addEventListener("message", handler);
    worker.postMessage({ buffer }, [buffer]);
  });
}

async function loadWorkbookLocally(file, note = "") {
  setStatus("Leyendo la planilla en este navegador…");
  setProgress("busy");
  const payload = await parseWithWorker(await file.arrayBuffer());
  applyPayload(payload, file.name, null);
  setProgress(null);
  setStatus(`Planilla cargada en este navegador.${note ? " " + note : ""}`, "ok");
}

async function handleFile(file) {
  if (!file) return;
  if (!/\.xlsx$/i.test(file.name)) {
    setStatus("El archivo tiene que ser una planilla .xlsx", "err");
    return;
  }
  try {
    if (state.serverMode && state.serverVersion) await uploadToServer(file);
    else await loadWorkbookLocally(file);
  } catch (error) {
    console.error(error);
    setProgress(null);
    setStatus(`No pude leer la planilla: ${error.message}`, "err");
  } finally {
    el.fileInput.value = "";
  }
}

function xhrUpload(file, token) {
  return new Promise((resolve, reject) => {
    const data = new FormData();
    data.append("stock", file, file.name);
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/upload");
    xhr.timeout = 300000;
    if (token) xhr.setRequestHeader("X-Upload-Token", token);
    xhr.upload.onprogress = (event) => {
      if (!event.lengthComputable) return;
      const pct = Math.round((event.loaded / event.total) * 100);
      setProgress(pct);
      setStatus(pct < 100 ? `Subiendo ${file.name} · ${pct}%` : "Procesando la planilla en el servidor…");
      if (pct >= 100) setProgress("busy");
    };
    xhr.onload = () => resolve({ status: xhr.status, text: xhr.responseText });
    xhr.onerror = () => reject(new Error("Sin conexión con el servidor"));
    xhr.ontimeout = () => reject(new Error("El servidor tardó demasiado"));
    xhr.send(data);
  });
}

async function uploadToServer(file) {
  setStatus(`Subiendo ${file.name}…`);
  setProgress(0);
  let response;
  try {
    response = await xhrUpload(file, getToken());
    if (response.status === 401) {
      const token = window.prompt("Esta instalación pide un código para cargar planillas. Ingresalo:");
      if (!token) { setProgress(null); setStatus("Carga cancelada: falta el código.", "err"); return; }
      localStorage.setItem("stock-upload-token", token.trim());
      setProgress(0);
      response = await xhrUpload(file, token.trim());
      if (response.status === 401) localStorage.removeItem("stock-upload-token");
    }
  } catch (error) {
    await loadWorkbookLocally(file, `(El servidor falló: ${error.message})`);
    return;
  }

  if (response.status === 404) {
    await loadWorkbookLocally(file, "El servidor no tiene /api/upload: redeploy con el Dockerfile nuevo.");
    return;
  }
  if (response.status < 200 || response.status >= 300) {
    setProgress(null);
    setStatus(response.text || `Error ${response.status}`, "err");
    return;
  }

  const meta = JSON.parse(response.text);
  setStatus("Descargando datos procesados…");
  const payload = await fetchStockData();
  applyPayload(payload, meta.name, meta);
  await loadHistory();
  setProgress(null);
  const secs = meta.processSeconds ? ` en ${String(meta.processSeconds).replace(".", ",")} s` : "";
  setStatus(`Planilla cargada: ${meta.name} · ${fmt(meta.productCount)} productos${secs}.`, "ok");
}

async function fetchStockData() {
  const response = await fetch("/api/stock-data", { cache: "no-cache" });
  if (!response.ok) throw new Error(await response.text());
  return response.json();
}

async function checkServerVersion() {
  try {
    const response = await fetch("/api/version", { cache: "no-store" });
    if (!response.ok) {
      el.serverInfo.textContent = "versión antigua";
      el.serverLed.className = "led warn";
      el.serverLed.title = "Servidor viejo: hacé rebuild en Portainer";
      return null;
    }
    const version = await response.json();
    el.serverInfo.textContent = `v${version.version}`;
    el.serverLed.className = "led ok";
    el.serverLed.title = `Servidor v${version.version} en línea`;
    return version;
  } catch {
    el.serverInfo.textContent = "sin conexión";
    el.serverLed.className = "led err";
    el.serverLed.title = "Sin servidor: la planilla se lee en este navegador";
    return null;
  }
}

async function init() {
  if (!state.serverMode) {
    setStatus("Modo local: elegí una planilla para empezar.");
    el.serverInfo.textContent = "modo local";
    renderAllEmpty();
    return;
  }
  setProgress("busy");
  state.serverVersion = await checkServerVersion();
  try {
    const info = await fetch("/api/current", { cache: "no-store" });
    if (info.status === 404) {
      renderAllEmpty();
      setProgress(null);
      setStatus(state.serverVersion ? "Servidor listo. Subí la planilla para empezar." : "Subí una planilla para empezar.");
      return;
    }
    if (!info.ok) throw new Error(await info.text());
    const meta = await info.json();
    setStatus(`Cargando ${meta.name}…`);
    applyPayload(await fetchStockData(), meta.name, meta);
    await loadHistory();
    setStatus(`Listo · ${meta.name} · ${fmt(meta.productCount)} productos.`, "ok");
  } catch (error) {
    renderAllEmpty();
    setStatus(`No pude cargar la planilla del servidor: ${error.message}`, "err");
  } finally {
    setProgress(null);
  }
}

/* =========================================================
   Preparación de datos
   ========================================================= */
function applyPayload(payload, displayName, meta) {
  state.header = payload.header || [];
  state.rows = payload.rows || [];
  state.index = payload.index || {};
  state.totals = payload.totals || {};
  state.lineas = payload.lineas || [];
  state.meta = meta;
  state.comparison = null;

  const h = state.header;
  state.cols = {
    concat: findHeaderIndex(h, "CONCAT"),
    sku: findHeaderIndex(h, "SKU"),
    color: findHeaderIndex(h, "COLOR"),
    talle: findHeaderIndex(h, "TALLE"),
    nombre: findHeaderIndex(h, "NOMBRE"),
    linea: findHeaderIndex(h, "LINEA"),
    bases: BASE_COLUMNS.map((b) => ({ base: b.base, idx: findHeaderIndex(h, b.col) })).filter((b) => b.idx >= 0)
  };
  state.baseColIdx = new Set(state.cols.bases.map((b) => b.idx));

  buildDerived();

  el.fileName.textContent = displayName;
  el.fileName.title = displayName;
  el.updatedAt.textContent = meta?.uploadedAt ? `${formatDate(meta.uploadedAt)} · ${timeAgo(meta.uploadedAt)}` : "en este navegador";
  el.dataLed.className = "led ok";
  el.downloadCurrent.classList.toggle("hidden", !meta);
  el.searchInput.disabled = false;
  el.compareBtn.disabled = !(state.serverMode && meta);
  el.exportStockBtn.disabled = false;

  populateFilters();
  renderKpis();
  renderTotals();
  renderQuality();
  renderHistory();
  renderComparisonEmpty();
  resetView();

  if (state.pendingQuery) {
    const q = state.pendingQuery;
    state.pendingQuery = null;
    el.searchInput.value = q;
    runSearch(q);
  } else if (state.current && state.index[state.current]) {
    showProduct(state.current, { silent: true });
  } else {
    renderResultIdle();
  }
}

function buildDerived() {
  const { cols, rows, index } = state;
  const rowInfo = new Array(rows.length);
  const stats = {};
  for (const b of cols.bases) stats[b.base] = { positive: 0, negativeCells: 0, negativeQty: 0 };

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    let hasStock = false;
    let hasNeg = false;
    for (const b of cols.bases) {
      const q = asNumber(row[b.idx]);
      if (q > 0) { hasStock = true; stats[b.base].positive++; }
      else if (q < 0) { hasNeg = true; stats[b.base].negativeCells++; stats[b.base].negativeQty += q; }
    }
    const concat = cols.concat >= 0 ? normalize(row[cols.concat]) : "";
    const key = concat || normalize([row[cols.sku], row[cols.color], row[cols.talle]].filter((v) => v !== undefined && v !== "").join(" "));
    rowInfo[i] = {
      key,
      hay: `${key} ${cols.sku >= 0 ? normalize(row[cols.sku]) : ""} ${cols.nombre >= 0 ? normalize(row[cols.nombre]) : ""}`,
      linea: cols.linea >= 0 ? normalize(row[cols.linea]) : "",
      hasStock,
      hasNeg
    };
  }
  state.rowInfo = rowInfo;
  state.baseStats = stats;

  const list = [];
  const bySku = new Map();
  for (const [key, product] of Object.entries(index)) {
    const total = productTotal(product.bases);
    list.push({ key, hay: `${key} ${normalize(product.nombre)}`, total, nombre: product.nombre });
    const sku = normalize(product.sku);
    if (sku) {
      if (!bySku.has(sku)) bySku.set(sku, []);
      bySku.get(sku).push({ key, total, color: product.color, talle: product.talle });
    }
  }
  state.searchList = list;
  state.bySku = bySku;
}

/* =========================================================
   KPIs, totales y calidad
   ========================================================= */
function renderKpis() {
  const totals = state.totals || {};
  let units = 0;
  for (const [base, qty] of Object.entries(totals)) if (base !== PENDING_BASE) units += qty;
  const stats = Object.entries(state.baseStats);
  const active = stats.filter(([base, s]) => base !== PENDING_BASE && s.positive > 0).length;
  const totalBases = stats.filter(([base]) => base !== PENDING_BASE).length;
  const neg = stats.reduce((sum, [, s]) => sum + s.negativeCells, 0);

  countUp(el.kpiProducts, state.rows.length);
  countUp(el.kpiUnits, units);
  countUp(el.kpiPending, totals[PENDING_BASE] || 0);
  countUp(el.kpiBases, active);
  el.kpiBasesSub.textContent = `de ${totalBases} con stock positivo`;
  countUp(el.kpiNeg, neg);
  el.kpiNegBtn.classList.toggle("alert", neg > 0);
}

function renderTotals() {
  const entries = Object.entries(state.totals || {});
  if (!entries.length) { el.totalsBox.innerHTML = emptyBox("Sin datos todavía.", "layers"); return; }
  const units = entries.filter(([b]) => b !== PENDING_BASE).reduce((s, [, q]) => s + Math.max(0, q), 0);
  const max = Math.max(1, ...entries.map(([, q]) => Math.abs(q)));
  entries.sort((a, b) => (a[0] === PENDING_BASE) - (b[0] === PENDING_BASE) || b[1] - a[1]);
  el.totalsBox.innerHTML = entries.map(([base, qty], i) => {
    const pend = base === PENDING_BASE;
    const cls = qty < 0 ? "neg" : qty === 0 ? "zero" : pend ? "pend" : "";
    const share = !pend && units > 0 && qty > 0 ? `<span class="b-share">${Math.round((qty / units) * 100)}%</span>` : "";
    return `<button class="bar-row ${cls}" type="button" data-base="${escapeHtml(base)}" title="Ver productos con stock en ${escapeHtml(base)}">
      <span class="b-name">${escapeHtml(base)}</span>
      <span class="b-track"><span class="b-fill" style="width:${(Math.abs(qty) / max) * 100}%;animation-delay:${i * 40}ms"></span></span>
      <span class="b-qty">${fmt(qty)}${share}</span>
    </button>`;
  }).join("");
}

function renderQuality() {
  const stats = Object.entries(state.baseStats);
  if (!state.rows.length) { el.qualityBox.innerHTML = emptyBox("Cargá una planilla para ver alertas y controles.", "alert"); return; }
  const withStock = state.rowInfo.filter((r) => r.hasStock).length;
  const noStock = state.rows.length - withStock;
  const negRows = state.rowInfo.filter((r) => r.hasNeg).length;
  const negBases = stats.filter(([, s]) => s.negativeCells > 0).sort((a, b) => b[1].negativeCells - a[1].negativeCells);

  el.qualityBox.innerHTML = `
    <div class="q-grid">
      <button class="q-tile" type="button" data-mode="stock"><div class="l">Con stock</div><div class="v">${fmt(withStock)}</div></button>
      <button class="q-tile" type="button" data-mode="nostock"><div class="l">Sin stock</div><div class="v">${fmt(noStock)}</div></button>
      <button class="q-tile ${negRows ? "bad" : ""}" type="button" data-mode="negative"><div class="l">Productos con negativos</div><div class="v">${fmt(negRows)}</div></button>
      <div class="q-tile"><div class="l">Cobertura</div><div class="v">${state.rows.length ? Math.round((withStock / state.rows.length) * 100) : 0}%</div></div>
    </div>
    ${negBases.length
      ? `<div class="neg-list">${negBases.map(([base, s]) => `
          <button class="neg-row" type="button" data-negative-base="${escapeHtml(base)}">
            <span class="who"><span class="led"></span>${escapeHtml(base)}</span>
            <span class="mono"><span class="down">${fmt(s.negativeCells)}</span> <span class="dim">reg · ${fmt(s.negativeQty)} u</span></span>
          </button>`).join("")}</div>`
      : `<div class="ok-line">${icon("check")}Sin stock negativo en ninguna base.</div>`}`;
}

/* =========================================================
   Búsqueda
   ========================================================= */
function findMatches(rawQuery, limit = SUGGEST_LIMIT) {
  const q = normalize(rawQuery);
  if (!q) return { items: [], total: 0 };
  const tokens = q.split(" ");
  const scored = [];
  for (const item of state.searchList) {
    let score;
    if (item.key === q) score = 0;
    else if (item.key.startsWith(q)) score = 1;
    else if (item.key.includes(q)) score = 2;
    else if (tokens.every((t) => item.hay.includes(t))) score = 3;
    else continue;
    scored.push({ ...item, score });
  }
  scored.sort((a, b) => a.score - b.score || a.key.localeCompare(b.key, "es", { numeric: true }));
  return { items: scored.slice(0, limit), total: scored.length, query: q };
}

function highlight(text, q) {
  const safe = escapeHtml(text);
  if (!q) return safe;
  const i = text.indexOf(q);
  if (i < 0) return safe;
  return escapeHtml(text.slice(0, i)) + "<mark>" + escapeHtml(text.slice(i, i + q.length)) + "</mark>" + escapeHtml(text.slice(i + q.length));
}

function updateSuggestions() {
  const value = el.searchInput.value;
  if (normalize(value).length < 2) { hideSuggestions(); return; }
  const { items, total, query } = findMatches(value);
  state.suggestions = items;
  state.activeSuggestion = items.length ? 0 : -1;
  if (!items.length) {
    el.suggest.innerHTML = `<div class="suggest-foot">Sin coincidencias para “${escapeHtml(value.trim())}”</div>`;
  } else {
    el.suggest.innerHTML = items.map((item, i) => `
      <button class="suggest-item ${i === 0 ? "active" : ""}" type="button" role="option" data-key="${escapeHtml(item.key)}">
        <span class="code">${highlight(item.key, query)}</span>
        <span class="name">${escapeHtml(item.nombre || "")}</span>
        <span class="qty ${item.total ? "" : "zero"}">${fmt(item.total)}</span>
      </button>`).join("") +
      `<div class="suggest-foot">${fmt(total)} coincidencia${total === 1 ? "" : "s"} · ↑↓ para moverte · Enter para abrir</div>`;
  }
  el.suggest.classList.remove("hidden");
}

function hideSuggestions() {
  el.suggest.classList.add("hidden");
  state.suggestions = [];
  state.activeSuggestion = -1;
}

function moveSuggestion(delta) {
  if (!state.suggestions.length) return;
  state.activeSuggestion = (state.activeSuggestion + delta + state.suggestions.length) % state.suggestions.length;
  el.suggest.querySelectorAll(".suggest-item").forEach((node, i) => node.classList.toggle("active", i === state.activeSuggestion));
}

function runSearch(rawQuery) {
  hideSuggestions();
  const query = String(rawQuery ?? "").trim();
  if (!query) { renderResultIdle(); return; }
  const { items } = findMatches(query, 1);
  if (!items.length) {
    state.current = null;
    toggleResultActions(false);
    el.queryLabel.textContent = query;
    el.resultBox.innerHTML = emptyBox(`No encontré <b class="mono">${escapeHtml(query)}</b> en la planilla.<br><span class="dim">Probá con menos datos, por ejemplo sólo el SKU.</span>`, "search");
    return;
  }
  showProduct(items[0].key);
}

function showProduct(key, { silent = false } = {}) {
  const product = state.index[key];
  if (!product) return;
  hideSuggestions();
  state.current = key;
  el.searchInput.value = key;
  el.queryLabel.textContent = "";
  toggleResultActions(true);
  if (!silent) {
    addRecent(key);
    const url = new URL(location.href);
    url.searchParams.set("q", key);
    history.replaceState(null, "", url);
  }

  const bases = Object.entries(product.bases || {});
  const total = productTotal(product.bases);
  const max = Math.max(1, ...bases.map(([, q]) => Math.abs(q)));
  const withStock = bases.filter(([b, q]) => b !== PENDING_BASE && q > 0).length;
  bases.sort((a, b) => (a[0] === PENDING_BASE) - (b[0] === PENDING_BASE) || b[1] - a[1]);

  const variants = (state.bySku.get(normalize(product.sku)) || [])
    .slice()
    .sort((a, b) => a.key.localeCompare(b.key, "es", { numeric: true }));

  el.resultBox.innerHTML = `
    <div class="product-head">
      <div>
        <div class="product-code">${escapeHtml(key)}</div>
        <div class="product-name">${escapeHtml(product.nombre || "Sin nombre")}</div>
        <div class="product-tags">
          ${product.sku !== "" ? `<span class="chip cyan">SKU ${escapeHtml(product.sku)}</span>` : ""}
          ${product.color !== "" ? `<span class="chip violet">COLOR ${escapeHtml(product.color)}</span>` : ""}
          ${product.talle !== "" ? `<span class="chip green">TALLE ${escapeHtml(product.talle)}</span>` : ""}
          ${product.linea ? `<span class="chip">${escapeHtml(product.linea)}</span>` : ""}
        </div>
      </div>
      <div class="product-total">
        <div class="n ${total ? "" : "zero"}">${fmt(total)}</div>
        <div class="l">unidades · ${withStock} base${withStock === 1 ? "" : "s"}</div>
      </div>
    </div>
    <div class="bars">
      ${bases.map(([base, qty], i) => {
        const cls = qty < 0 ? "neg" : qty === 0 ? "zero" : base === PENDING_BASE ? "pend" : "";
        return `<button class="bar-row ${cls}" type="button" data-base="${escapeHtml(base)}" title="Ver todo el stock de ${escapeHtml(base)}">
          <span class="b-name">${escapeHtml(base)}</span>
          <span class="b-track"><span class="b-fill" style="width:${(Math.abs(qty) / max) * 100}%;animation-delay:${i * 35}ms"></span></span>
          <span class="b-qty">${fmt(qty)}</span>
        </button>`;
      }).join("")}
    </div>
    ${variants.length > 1 ? `
      <div class="variants">
        <h3>Otras variantes del SKU ${escapeHtml(product.sku)} · ${variants.length}</h3>
        <div class="variant-grid">
          ${variants.map((v) => `<button class="variant ${v.key === key ? "current" : ""} ${v.total ? "" : "none"}" type="button" data-key="${escapeHtml(v.key)}" title="${escapeHtml(v.key)}">${escapeHtml([v.color, v.talle].filter((x) => x !== "").join(" · ") || v.key)} <b>${fmt(v.total)}</b></button>`).join("")}
        </div>
      </div>` : ""}`;
}

function renderResultIdle() {
  toggleResultActions(false);
  el.queryLabel.textContent = "";
  if (!state.rows.length) { renderWelcome(); return; }
  el.resultBox.innerHTML = emptyBox(`Escribí un código como <b class="mono">25551 NLSA M</b>, un SKU o parte del nombre.<br><span class="dim">Atajo: tecla <kbd>/</kbd> para buscar desde cualquier lugar.</span>`, "search");
}

function renderWelcome() {
  el.resultBox.innerHTML = `
    <div class="welcome">
      <div class="ring">${icon("upload")}</div>
      <h2>Cargá la planilla de stock</h2>
      <p>Arrastrá el archivo .xlsx a cualquier parte de la pantalla o usá el botón. Se procesa una sola vez y queda disponible para todos los equipos.</p>
      <button class="btn primary" type="button" data-action="upload">${icon("upload")}Elegir planilla</button>
    </div>`;
}

function toggleResultActions(on) {
  el.exportCopyBtn.disabled = !on;
  el.exportCsvBtn.disabled = !on;
  el.shareBtn.disabled = !on || !state.serverMode;
}

function addRecent(key) {
  const list = JSON.parse(localStorage.getItem("stock-recent") || "[]").filter((k) => k !== key);
  list.unshift(key);
  localStorage.setItem("stock-recent", JSON.stringify(list.slice(0, RECENT_MAX)));
  renderRecent();
}

function renderRecent() {
  const list = JSON.parse(localStorage.getItem("stock-recent") || "[]").filter((k) => state.index[k]);
  el.recent.innerHTML = list.length
    ? `<span class="label">Recientes</span>` + list.map((k) => `<button class="chip" type="button" data-key="${escapeHtml(k)}">${escapeHtml(k)}</button>`).join("")
    : "";
}

/* =========================================================
   Exportar consulta
   ========================================================= */
function currentProductLines() {
  const product = state.index[state.current];
  const lines = [
    `Producto: ${state.current}`,
    `SKU: ${product.sku}`,
    `Color: ${product.color}`,
    `Talle: ${product.talle}`,
    `Nombre: ${product.nombre}`,
    ""
  ];
  for (const [base, qty] of Object.entries(product.bases)) if (qty > 0) lines.push(`${base}: ${fmt(qty)}`);
  lines.push("", `Total (sin pendientes): ${fmt(productTotal(product.bases))}`);
  return lines;
}

async function copyText(text, okMessage) {
  try {
    await navigator.clipboard.writeText(text);
    setStatus(okMessage, "ok");
  } catch {
    setStatus("No pude copiar al portapapeles (el navegador lo bloqueó).", "err");
  }
}

function exportSearchCsv() {
  const product = state.index[state.current];
  if (!product) return;
  const rows = [
    ["campo", "valor"], ["consulta", state.current], ["sku", product.sku], ["color", product.color],
    ["talle", product.talle], ["nombre", product.nombre], ["linea", product.linea], [], ["base", "cantidad"]
  ];
  for (const [base, qty] of Object.entries(product.bases)) if (qty > 0) rows.push([base, qty]);
  downloadCsv(`stock-${state.current.replace(/\s+/g, "-")}.csv`, rows);
  setStatus("CSV descargado.", "ok");
}

/* =========================================================
   Tabla de stock completo (virtualizada)
   ========================================================= */
function populateFilters() {
  el.stockFilterLinea.innerHTML = '<option value="">Todas las líneas</option>' +
    state.lineas.map((l) => `<option value="${escapeHtml(l)}">${escapeHtml(l)}</option>`).join("");
  el.stockFilterLinea.value = state.lineas.includes(state.filter.linea) ? state.filter.linea : "";
  el.stockFilterBase.innerHTML = '<option value="">Todas las bases</option>' +
    state.cols.bases.map((b) => `<option value="${escapeHtml(b.base)}">Con stock en ${escapeHtml(b.base)}</option>`).join("");
  el.stockFilterBase.value = state.filter.base;
}

function computeView() {
  const { rows, rowInfo, cols, filter, sort } = state;
  const text = normalize(filter.text);
  const tokens = text ? text.split(" ") : [];
  const linea = normalize(filter.linea);
  const baseIdx = filter.base ? (cols.bases.find((b) => b.base === filter.base)?.idx ?? -1) : -1;
  const negIdx = filter.negBase ? (cols.bases.find((b) => b.base === filter.negBase)?.idx ?? -1) : -1;

  const view = [];
  for (let i = 0; i < rows.length; i++) {
    const info = rowInfo[i];
    if (filter.mode === "stock" && !info.hasStock) continue;
    if (filter.mode === "nostock" && info.hasStock) continue;
    if (filter.mode === "negative" && !info.hasNeg) continue;
    if (filter.base && (baseIdx < 0 || asNumber(rows[i][baseIdx]) <= 0)) continue;
    if (filter.negBase && (negIdx < 0 || asNumber(rows[i][negIdx]) >= 0)) continue;
    if (linea && info.linea !== linea) continue;
    if (tokens.length && !tokens.every((t) => info.hay.includes(t))) continue;
    view.push(i);
  }

  if (sort.col >= 0) {
    const c = sort.col;
    const numeric = state.baseColIdx.has(c) || view.slice(0, 50).every((i) => rows[i][c] === "" || typeof rows[i][c] === "number");
    view.sort((a, b) => {
      const va = rows[a][c], vb = rows[b][c];
      const r = numeric ? asNumber(va) - asNumber(vb) : String(va ?? "").localeCompare(String(vb ?? ""), "es", { numeric: true });
      return r * sort.dir;
    });
  }
  state.view = view;
}

function resetView() {
  computeView();
  state.lastStart = -1;
  el.stockTable.scrollTop = 0;
  renderStockTable();
  renderActiveFilter();
}

function renderActiveFilter() {
  const chips = [];
  if (state.filter.negBase) chips.push(`<span class="chip red">Negativos en ${escapeHtml(state.filter.negBase)} <button class="btn sm icon" style="height:20px;width:20px;border:0;background:none" data-clear="negBase" aria-label="Quitar">${icon("x", 'style="width:12px;height:12px"')}</button></span>`);
  el.activeFilter.innerHTML = chips.length ? `<span class="dim" style="font-size:12px">${icon("filter", 'style="width:14px;height:14px;vertical-align:-3px"')} Filtro activo:</span>${chips.join("")}` : "";
  el.activeFilter.classList.toggle("hidden", !chips.length);
}

function renderStockTable() {
  const { header, view, sort } = state;
  if (!header.length) {
    el.stockTable.innerHTML = `<div style="padding:18px">${emptyBox("La tabla aparece cuando cargás una planilla.", "table")}</div>`;
    el.stockInfo.textContent = "";
    el.stockFoot.textContent = "";
    return;
  }
  el.stockInfo.textContent = `${fmt(view.length)} de ${fmt(state.rows.length)} filas`;
  el.stockFoot.textContent = `${fmt(view.length)} filas visibles`;
  el.stockTable.innerHTML = `
    <table>
      <thead><tr>${header.map((cell, i) => `<th data-col="${i}" class="${sort.col === i ? "sorted" : ""}">${escapeHtml(cell)}${sort.col === i ? `<span class="arrow">${sort.dir > 0 ? "▲" : "▼"}</span>` : ""}</th>`).join("")}</tr></thead>
      <tbody></tbody>
    </table>
    ${view.length ? "" : `<div style="padding:18px">${emptyBox("Ningún producto coincide con los filtros.", "filter")}</div>`}`;
  state.lastStart = -1;
  renderVisibleRows();
}

function renderVisibleRows() {
  const tbody = el.stockTable.querySelector("tbody");
  if (!tbody) return;
  const { view, rows, header } = state;
  const viewport = el.stockTable.clientHeight || 520;
  const scrollTop = Math.max(0, el.stockTable.scrollTop - 38);
  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - OVERSCAN);
  if (start === state.lastStart) return;
  state.lastStart = start;
  const end = Math.min(view.length, start + Math.ceil(viewport / ROW_H) + OVERSCAN * 2);
  const cols = header.length;
  const baseCols = state.baseColIdx;

  let html = `<tr class="spacer"><td colspan="${cols}" style="height:${start * ROW_H}px"></td></tr>`;
  for (let v = start; v < end; v++) {
    const i = view[v];
    const row = rows[i];
    html += `<tr class="data" data-row="${i}">`;
    for (let c = 0; c < cols; c++) {
      const value = row[c] ?? "";
      let cls = "";
      if (baseCols.has(c)) {
        const q = asNumber(value);
        cls = `num ${q > 0 ? "pos" : q < 0 ? "neg" : "z"}`;
      } else if (typeof value === "number") cls = "num";
      html += `<td class="${cls}">${escapeHtml(baseCols.has(c) && typeof value === "number" ? fmt(value) : value)}</td>`;
    }
    html += "</tr>";
  }
  html += `<tr class="spacer"><td colspan="${cols}" style="height:${Math.max(0, (view.length - end) * ROW_H)}px"></td></tr>`;
  tbody.innerHTML = html;
}

function applyFilter(patch, scroll = false) {
  Object.assign(state.filter, patch);
  el.stockFilterText.value = state.filter.text;
  el.stockFilterLinea.value = state.filter.linea;
  el.stockFilterBase.value = state.filter.base;
  el.stockFilterMode.value = state.filter.mode;
  resetView();
  if (scroll) el.stockSection.scrollIntoView({ behavior: "smooth", block: "start" });
}

function exportStockView() {
  const rows = [state.header, ...state.view.map((i) => state.rows[i])];
  const stamp = new Date().toISOString().slice(0, 10);
  downloadCsv(`stock-vista-${stamp}.csv`, rows);
  setStatus(`Exportadas ${fmt(state.view.length)} filas.`, "ok");
}

/* =========================================================
   Historial
   ========================================================= */
async function loadHistory() {
  if (!state.serverMode) return;
  try {
    const response = await fetch("/api/history", { cache: "no-store" });
    if (!response.ok) return;
    state.history = await response.json();
  } catch {
    state.history = [];
  }
  renderHistory();
}

function renderHistory() {
  if (!state.serverMode) { el.historyBox.innerHTML = emptyBox("El historial está disponible cuando la app corre en el servidor.", "history"); return; }
  if (!state.history.length) { el.historyBox.innerHTML = emptyBox("Todavía no hay planillas guardadas.", "history"); el.historyInfo.textContent = ""; return; }
  const currentId = state.meta?.id ?? state.history[0]?.id;
  const selected = state.comparison?.previous?.id;
  el.historyInfo.textContent = `${state.history.length} carga${state.history.length === 1 ? "" : "s"}`;
  const deleteEnabled = (state.serverVersion?.features || []).includes("delete");
  const downloadEnabled = (state.serverVersion?.features || []).includes("download");
  el.historyBox.innerHTML = `<div class="timeline">${state.history.map((entry) => {
    const isCurrent = entry.id === currentId;
    return `<div class="t-item ${isCurrent ? "current" : ""} ${entry.id === selected ? "selected" : ""}">
      <span class="t-dot"></span>
      <div style="min-width:0">
        <div class="t-name" title="${escapeHtml(entry.name)}">${escapeHtml(entry.name)}</div>
        <div class="t-meta">${formatDate(entry.uploadedAt)} · ${fmt(entry.productCount)} productos${entry.size ? " · " + formatBytes(entry.size) : ""}</div>
      </div>
      <div class="t-actions">
        ${isCurrent ? '<span class="chip cyan">ACTUAL</span>' : `<button class="btn sm" type="button" data-compare-id="${escapeHtml(entry.id)}" title="Comparar con la actual">${icon("compare")}Comparar</button>`}
        ${downloadEnabled ? `<a class="btn sm icon" href="/api/history/download?id=${encodeURIComponent(entry.id)}" title="Descargar esta planilla">${icon("download")}</a>` : ""}
        ${!isCurrent && deleteEnabled ? `<button class="btn sm icon danger" type="button" data-delete-id="${escapeHtml(entry.id)}" title="Eliminar del historial">${icon("trash")}</button>` : ""}
      </div>
    </div>`;
  }).join("")}</div>`;
}

async function deleteHistory(id) {
  const entry = state.history.find((e) => e.id === id);
  if (!entry || !window.confirm(`¿Eliminar "${entry.name}" (${formatDate(entry.uploadedAt)}) del historial?`)) return;
  let response = await fetch(`/api/history?id=${encodeURIComponent(id)}`, { method: "DELETE", headers: { "X-Upload-Token": getToken() } });
  if (response.status === 401) {
    const token = window.prompt("Ingresá el código de carga para eliminar:");
    if (!token) return;
    localStorage.setItem("stock-upload-token", token.trim());
    response = await fetch(`/api/history?id=${encodeURIComponent(id)}`, { method: "DELETE", headers: { "X-Upload-Token": token.trim() } });
  }
  if (!response.ok) { setStatus(`No pude eliminar: ${await response.text()}`, "err"); return; }
  if (state.comparison?.previous?.id === id) { state.comparison = null; renderComparisonEmpty(); }
  await loadHistory();
  setStatus(`Eliminada del historial: ${entry.name}`, "ok");
}

/* =========================================================
   Comparación
   ========================================================= */
function renderComparisonEmpty() {
  el.exportCompareBtn.disabled = true;
  el.compareLabel.textContent = "";
  el.compareBox.innerHTML = emptyBox(
    state.serverMode
      ? "Compará la planilla actual con la anterior o elegí una del historial para ver qué cambió."
      : "La comparación está disponible cuando la app corre en el servidor.",
    "compare"
  );
}

async function compare(historyId = "") {
  try {
    setStatus("Comparando planillas…");
    setProgress("busy");
    const url = historyId ? `/api/compare?id=${encodeURIComponent(historyId)}` : "/api/compare";
    const response = await fetch(url, { cache: "no-store" });
    if (!response.ok) throw new Error(await response.text());
    state.comparison = await response.json();
    state.compareLimit = COMPARE_PAGE;
    state.compareFilter = { text: "", base: "", type: "" };
    renderComparison();
    renderHistory();
    $("comparePanel").scrollIntoView({ behavior: "smooth", block: "start" });
    setStatus(`Comparación lista contra ${state.comparison.previous.name}.`, "ok");
  } catch (error) {
    setStatus(`No pude comparar: ${error.message}`, "err");
  } finally {
    setProgress(null);
  }
}

function filteredChanges() {
  const { text, base, type } = state.compareFilter;
  const tokens = normalize(text).split(" ").filter(Boolean);
  return (state.comparison?.changes || []).filter((c) => {
    if (base && c.base !== base) return false;
    if (type === "up" && !(c.delta > 0)) return false;
    if (type === "down" && !(c.delta < 0)) return false;
    if ((type === "new" || type === "removed") && c.type !== type) return false;
    if (tokens.length) {
      const hay = `${c.concat} ${normalize(c.nombre)}`;
      if (!tokens.every((t) => hay.includes(t))) return false;
    }
    return true;
  });
}

function renderComparison() {
  const data = state.comparison;
  if (!data) { renderComparisonEmpty(); return; }
  const changes = data.changes || [];
  const s = data.summary || {};
  el.compareLabel.textContent = `vs ${data.previous.name} · ${formatDate(data.previous.uploadedAt)}`;
  el.exportCompareBtn.disabled = !changes.length;

  if (!changes.length) {
    el.compareBox.innerHTML = `<div class="ok-line" style="margin-top:0">${icon("check")}No hay diferencias de stock entre las dos planillas.</div>`;
    return;
  }

  const unitsIn = s.unitsIn ?? changes.filter((c) => c.delta > 0).reduce((a, c) => a + c.delta, 0);
  const unitsOut = s.unitsOut ?? changes.filter((c) => c.delta < 0).reduce((a, c) => a + c.delta, 0);
  const changedProducts = s.changedProducts ?? new Set(changes.map((c) => c.concat)).size;
  let byBase = s.byBase;
  if (!byBase) {
    byBase = {};
    for (const c of changes) {
      const e = (byBase[c.base] ||= { in: 0, out: 0, net: 0, lines: 0 });
      e.lines++; e.net += c.delta;
      if (c.delta > 0) e.in += c.delta; else e.out += c.delta;
    }
  }
  const baseEntries = Object.entries(byBase).sort((a, b) => Math.abs(b[1].net) - Math.abs(a[1].net));
  const maxAbs = Math.max(1, ...baseEntries.map(([, e]) => Math.max(e.in, Math.abs(e.out))));
  const bases = baseEntries.map(([b]) => b).sort();

  el.compareBox.innerHTML = `
    <div class="cmp-kpis">
      <div class="q-tile"><div class="l">Productos con cambios</div><div class="v">${fmt(changedProducts)}</div></div>
      <div class="q-tile"><div class="l">Nuevos / eliminados</div><div class="v" style="white-space:nowrap"><span class="up">${fmt(s.newProducts ?? 0)}</span> <span class="dim">/</span> <span class="down">${fmt(s.removedProducts ?? 0)}</span></div></div>
      <div class="q-tile"><div class="l">Unidades que suben</div><div class="v up">${signed(unitsIn)}</div></div>
      <div class="q-tile"><div class="l">Unidades que bajan</div><div class="v down">${fmt(unitsOut)}</div></div>
    </div>
    <div class="diverge">
      ${baseEntries.map(([base, e]) => `
        <div class="dv-row" title="${escapeHtml(base)}: +${fmt(e.in)} / ${fmt(e.out)} · ${fmt(e.lines)} cambios">
          <span class="n">${escapeHtml(base)}</span>
          <span class="l"><span style="width:${(Math.abs(e.out) / maxAbs) * 100}%"></span></span>
          <span class="r"><span style="width:${(e.in / maxAbs) * 100}%"></span></span>
          <span class="v ${e.net > 0 ? "up" : e.net < 0 ? "down" : "dim"}">${signed(e.net)}</span>
        </div>`).join("")}
    </div>
    <div class="cmp-filters">
      <input id="cmpText" class="field mono" type="text" placeholder="Filtrar producto…" value="${escapeHtml(state.compareFilter.text)}" autocomplete="off">
      <select id="cmpBase" class="field"><option value="">Todas las bases</option>${bases.map((b) => `<option ${state.compareFilter.base === b ? "selected" : ""}>${escapeHtml(b)}</option>`).join("")}</select>
      <select id="cmpType" class="field">
        ${[["", "Todos los cambios"], ["up", "Sólo subas"], ["down", "Sólo bajas"], ["new", "Productos nuevos"], ["removed", "Productos eliminados"]]
          .map(([v, l]) => `<option value="${v}" ${state.compareFilter.type === v ? "selected" : ""}>${l}</option>`).join("")}
      </select>
    </div>
    <div id="cmpTable"></div>`;
  renderComparisonTable();
}

function renderComparisonTable() {
  const box = $("cmpTable");
  if (!box) return;
  const list = filteredChanges();
  const limited = list.slice(0, state.compareLimit);
  const tag = { new: "nuevo", removed: "eliminado", changed: "cambio" };
  box.innerHTML = list.length ? `
    <div class="table-wrap" style="max-height:420px;border:1px solid var(--line);border-radius:12px">
      <table>
        <thead><tr><th>Producto</th><th>Base</th><th>Tipo</th><th style="text-align:right">Antes</th><th style="text-align:right">Ahora</th><th style="text-align:right">Δ</th></tr></thead>
        <tbody>${limited.map((c) => `
          <tr class="data" data-key="${escapeHtml(c.concat)}" title="${escapeHtml(c.nombre || "")}">
            <td>${escapeHtml(c.concat)}</td>
            <td>${escapeHtml(c.base)}</td>
            <td><span class="type-tag ${c.type}">${tag[c.type] || c.type}</span></td>
            <td class="num">${fmt(c.before)}</td>
            <td class="num">${fmt(c.after)}</td>
            <td class="num ${c.delta > 0 ? "pos" : c.delta < 0 ? "neg" : ""}">${signed(c.delta)}</td>
          </tr>`).join("")}
        </tbody>
      </table>
    </div>
    <div class="table-foot" style="border:0;padding:10px 2px 0">
      <span>Mostrando ${fmt(limited.length)} de ${fmt(list.length)} cambios</span>
      ${list.length > limited.length ? '<button class="btn sm" type="button" id="cmpMore">Mostrar más</button>' : ""}
    </div>` : emptyBox("Ningún cambio coincide con el filtro.", "filter");
}

function exportComparison() {
  const data = state.comparison;
  if (!data) return;
  const rows = [["producto", "nombre", "base", "tipo", "antes", "ahora", "delta"]];
  for (const c of filteredChanges()) rows.push([c.concat, c.nombre, c.base, c.type, c.before, c.after, c.delta]);
  downloadCsv(`comparacion-${data.previous.id}.csv`, rows);
  setStatus(`Exportados ${fmt(rows.length - 1)} cambios.`, "ok");
}

/* =========================================================
   Escáner
   ========================================================= */
async function startScanner() {
  if (!("BarcodeDetector" in window)) {
    setStatus("Este navegador no soporta escaneo por cámara. Usá Chrome en Android.", "err");
    return;
  }
  try {
    scannerStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: "environment" } }, audio: false });
    el.scannerVideo.srcObject = scannerStream;
    await el.scannerVideo.play();
    el.scannerModal.classList.remove("hidden");
    const detector = new BarcodeDetector({ formats: ["code_128", "code_39", "ean_13", "ean_8", "qr_code"] });
    scannerLoop = window.setInterval(async () => {
      try {
        const codes = await detector.detect(el.scannerVideo);
        const raw = codes[0]?.rawValue || "";
        if (!raw) return;
        stopScanner();
        if (navigator.vibrate) navigator.vibrate(60);
        el.searchInput.value = raw.trim();
        runSearch(raw.trim());
        setStatus(`Código escaneado: ${raw}`, "ok");
      } catch { /* errores intermitentes del detector */ }
    }, 300);
  } catch (error) {
    setStatus(`No pude abrir la cámara: ${error.message}`, "err");
  }
}

function stopScanner() {
  el.scannerModal.classList.add("hidden");
  if (scannerLoop) { clearInterval(scannerLoop); scannerLoop = null; }
  if (scannerStream) { scannerStream.getTracks().forEach((t) => t.stop()); scannerStream = null; }
  el.scannerVideo.srcObject = null;
}

/* =========================================================
   Tema, navegación, reloj
   ========================================================= */
function syncThemeIcon() {
  const dark = document.documentElement.dataset.theme !== "light";
  el.themeBtn.innerHTML = icon(dark ? "sun" : "moon");
  el.themeBtn.dataset.tip = dark ? "Modo claro" : "Modo oscuro";
  document.querySelector('meta[name="theme-color"]').content = dark ? "#060a13" : "#eef2f8";
}

function toggleTheme() {
  const next = document.documentElement.dataset.theme === "light" ? "dark" : "light";
  document.documentElement.dataset.theme = next;
  localStorage.setItem("stock-theme", next);
  syncThemeIcon();
}

function setupNavObserver() {
  const links = [...document.querySelectorAll("[data-nav]")];
  const targets = links.map((a) => document.querySelector(a.getAttribute("href"))).filter(Boolean);
  if (!("IntersectionObserver" in window)) return;
  const observer = new IntersectionObserver((entries) => {
    const visible = entries.filter((e) => e.isIntersecting).sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top)[0];
    if (!visible) return;
    links.forEach((a) => a.classList.toggle("active", a.getAttribute("href") === `#${visible.target.id}`));
  }, { rootMargin: "-20% 0px -60% 0px" });
  targets.forEach((t) => observer.observe(t));
}

function tickClock() {
  el.clock.textContent = new Intl.DateTimeFormat("es-AR", { hour: "2-digit", minute: "2-digit" }).format(new Date());
  if (state.meta?.uploadedAt) el.updatedAt.textContent = `${formatDate(state.meta.uploadedAt)} · ${timeAgo(state.meta.uploadedAt)}`;
}

function renderAllEmpty() {
  renderWelcome();
  renderTotals();
  renderQuality();
  renderHistory();
  renderComparisonEmpty();
  renderStockTable();
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator) || location.protocol === "file:") return;
  navigator.serviceWorker.register("/sw.js").catch(() => {});
}

/* =========================================================
   Eventos
   ========================================================= */
el.uploadBtn.addEventListener("click", () => el.fileInput.click());
el.fileInput.addEventListener("change", () => handleFile(el.fileInput.files?.[0]));
el.scanBtn.addEventListener("click", startScanner);
el.scannerClose.addEventListener("click", stopScanner);
el.themeBtn.addEventListener("click", toggleTheme);
el.exportCopyBtn.addEventListener("click", () => state.current && copyText(currentProductLines().join("\n"), "Resultado copiado al portapapeles."));
el.exportCsvBtn.addEventListener("click", exportSearchCsv);
el.shareBtn.addEventListener("click", () => state.current && copyText(location.href, "Enlace copiado: abre directamente esta consulta."));
el.compareBtn.addEventListener("click", () => compare());
el.exportCompareBtn.addEventListener("click", exportComparison);
el.exportStockBtn.addEventListener("click", exportStockView);
el.kpiNegBtn.addEventListener("click", () => state.rows.length && applyFilter({ mode: "negative", negBase: "", base: "" }, true));

el.searchInput.addEventListener("input", () => {
  clearTimeout(suggestTimer);
  suggestTimer = setTimeout(updateSuggestions, 70);
});
el.searchInput.addEventListener("keydown", (event) => {
  if (event.key === "ArrowDown") { event.preventDefault(); moveSuggestion(1); }
  else if (event.key === "ArrowUp") { event.preventDefault(); moveSuggestion(-1); }
  else if (event.key === "Enter") {
    event.preventDefault();
    clearTimeout(suggestTimer);
    const chosen = state.suggestions[state.activeSuggestion];
    if (chosen && !el.suggest.classList.contains("hidden")) showProduct(chosen.key);
    else runSearch(el.searchInput.value);
  } else if (event.key === "Escape") {
    if (!el.suggest.classList.contains("hidden")) hideSuggestions();
    else { el.searchInput.value = ""; el.searchInput.blur(); }
  }
});
el.searchInput.addEventListener("focus", () => { if (normalize(el.searchInput.value).length >= 2 && el.searchInput.value !== state.current) updateSuggestions(); });
el.searchInput.addEventListener("blur", () => setTimeout(hideSuggestions, 160));
el.suggest.addEventListener("mousedown", (event) => {
  const item = event.target.closest("[data-key]");
  if (item) { event.preventDefault(); showProduct(item.dataset.key); el.searchInput.blur(); }
});
el.recent.addEventListener("click", (event) => {
  const chip = event.target.closest("[data-key]");
  if (chip) showProduct(chip.dataset.key);
});

el.resultBox.addEventListener("click", (event) => {
  if (event.target.closest("[data-action='upload']")) { el.fileInput.click(); return; }
  const variant = event.target.closest(".variant[data-key]");
  if (variant) { showProduct(variant.dataset.key); return; }
  const bar = event.target.closest("[data-base]");
  if (bar) applyFilter({ base: bar.dataset.base === PENDING_BASE ? "" : bar.dataset.base, negBase: "", mode: "stock", text: "", linea: "" }, true);
});
el.totalsBox.addEventListener("click", (event) => {
  const bar = event.target.closest("[data-base]");
  if (bar && bar.dataset.base !== PENDING_BASE) applyFilter({ base: bar.dataset.base, negBase: "", mode: "stock", text: "", linea: "" }, true);
});
el.qualityBox.addEventListener("click", (event) => {
  const mode = event.target.closest("[data-mode]");
  if (mode) { applyFilter({ mode: mode.dataset.mode, base: "", negBase: "", text: "" }, true); return; }
  const neg = event.target.closest("[data-negative-base]");
  if (neg) applyFilter({ negBase: neg.dataset.negativeBase, base: "", mode: "", text: "", linea: "" }, true);
});

el.stockFilterText.addEventListener("input", () => {
  clearTimeout(filterTimer);
  filterTimer = setTimeout(() => { state.filter.text = el.stockFilterText.value; resetView(); }, 120);
});
el.stockFilterLinea.addEventListener("change", () => applyFilter({ linea: el.stockFilterLinea.value }));
el.stockFilterBase.addEventListener("change", () => applyFilter({ base: el.stockFilterBase.value, negBase: "" }));
el.stockFilterMode.addEventListener("change", () => applyFilter({ mode: el.stockFilterMode.value }));
el.clearFiltersBtn.addEventListener("click", () => { state.sort = { col: -1, dir: 1 }; applyFilter({ text: "", linea: "", base: "", mode: "", negBase: "" }); });
el.activeFilter.addEventListener("click", (event) => {
  const btn = event.target.closest("[data-clear]");
  if (btn) applyFilter({ [btn.dataset.clear]: "" });
});
el.stockTable.addEventListener("scroll", () => {
  if (scrollRaf) return;
  scrollRaf = requestAnimationFrame(() => { scrollRaf = 0; renderVisibleRows(); });
}, { passive: true });
el.stockTable.addEventListener("click", (event) => {
  const th = event.target.closest("th[data-col]");
  if (th) {
    const col = Number(th.dataset.col);
    state.sort = state.sort.col === col ? { col, dir: state.sort.dir === 1 ? -1 : 1 } : { col, dir: state.baseColIdx.has(col) ? -1 : 1 };
    resetView();
    return;
  }
  const tr = event.target.closest("tr[data-row]");
  if (tr) {
    const key = state.rowInfo[Number(tr.dataset.row)]?.key;
    if (key && state.index[key]) { showProduct(key); $("resultPanel").scrollIntoView({ behavior: "smooth", block: "start" }); }
  }
});

el.historyBox.addEventListener("click", (event) => {
  const cmp = event.target.closest("[data-compare-id]");
  if (cmp) { compare(cmp.dataset.compareId); return; }
  const del = event.target.closest("[data-delete-id]");
  if (del) deleteHistory(del.dataset.deleteId);
});
el.compareBox.addEventListener("input", (event) => {
  if (event.target.id !== "cmpText") return;
  clearTimeout(filterTimer);
  filterTimer = setTimeout(() => { state.compareFilter.text = event.target.value; state.compareLimit = COMPARE_PAGE; renderComparisonTable(); }, 150);
});
el.compareBox.addEventListener("change", (event) => {
  if (event.target.id === "cmpBase") state.compareFilter.base = event.target.value;
  else if (event.target.id === "cmpType") state.compareFilter.type = event.target.value;
  else return;
  state.compareLimit = COMPARE_PAGE;
  renderComparisonTable();
});
el.compareBox.addEventListener("click", (event) => {
  if (event.target.closest("#cmpMore")) { state.compareLimit += COMPARE_PAGE; renderComparisonTable(); return; }
  const tr = event.target.closest("tr[data-key]");
  if (tr && state.index[tr.dataset.key]) { showProduct(tr.dataset.key); $("resultPanel").scrollIntoView({ behavior: "smooth", block: "start" }); }
});

// Atajos de teclado
document.addEventListener("keydown", (event) => {
  const typing = /INPUT|SELECT|TEXTAREA/.test(document.activeElement?.tagName || "");
  if ((event.key === "/" && !typing) || ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k")) {
    if (el.searchInput.disabled) return;
    event.preventDefault();
    el.searchInput.focus();
    el.searchInput.select();
  } else if (event.key === "Escape" && !el.scannerModal.classList.contains("hidden")) {
    stopScanner();
  }
});

// Arrastrar y soltar la planilla
let dragDepth = 0;
const hasFiles = (event) => [...(event.dataTransfer?.types || [])].includes("Files");
window.addEventListener("dragenter", (event) => { if (!hasFiles(event)) return; dragDepth++; el.dropOverlay.classList.remove("hidden"); });
window.addEventListener("dragleave", () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) el.dropOverlay.classList.add("hidden"); });
window.addEventListener("dragover", (event) => { if (hasFiles(event)) event.preventDefault(); });
window.addEventListener("drop", (event) => {
  if (!hasFiles(event)) return;
  event.preventDefault();
  dragDepth = 0;
  el.dropOverlay.classList.add("hidden");
  handleFile(event.dataTransfer.files[0]);
});

/* =========================================================
   Inicio
   ========================================================= */
syncThemeIcon();
setupNavObserver();
renderRecent();
tickClock();
setInterval(tickClock, 30000);
registerServiceWorker();
init().then(renderRecent);
