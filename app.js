window.__stockAppLoaded = true;

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

const VIRTUAL_ROW_HEIGHT = 32;
const VIRTUAL_OVERSCAN = 8;

const state = {
  stockHeader: [],
  stockRows: [],
  index: {},
  inventory: {},
  totals: {},
  baseStats: {},
  lineas: [],
  pendingByConcat: {},
  meta: null,
  history: [],
  lastSearch: null,
  stockVisible: false,
  serverMode: location.protocol === "http:" || location.protocol === "https:",
  stockFilter: { text: "", linea: "", base: "", negativeBase: "", onlyStock: false },
  virtualScrollTop: 0
};

const fileInput = document.getElementById("fileInput");
const searchInput = document.getElementById("searchInput");
const searchBtn = document.getElementById("searchBtn");
const stockBtn = document.getElementById("stockBtn");
const scanBtn = document.getElementById("scanBtn");
const themeBtn = document.getElementById("themeBtn");
const exportCopyBtn = document.getElementById("exportCopyBtn");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const compareBtn = document.getElementById("compareBtn");
const statusBox = document.getElementById("status");
const fileName = document.getElementById("fileName");
const updatedAt = document.getElementById("updatedAt");
const serverInfo = document.getElementById("serverInfo");
const rowCount = document.getElementById("rowCount");
const resultBox = document.getElementById("resultBox");
const queryLabel = document.getElementById("queryLabel");
const totalsBox = document.getElementById("totalsBox");
const qualityBox = document.getElementById("qualityBox");
const stockSection = document.getElementById("stockSection");
const stockTable = document.getElementById("stockTable");
const stockInfo = document.getElementById("stockInfo");
const stockFilterText = document.getElementById("stockFilterText");
const stockFilterLinea = document.getElementById("stockFilterLinea");
const stockFilterBase = document.getElementById("stockFilterBase");
const stockFilterOnly = document.getElementById("stockFilterOnly");
const historyBox = document.getElementById("historyBox");
const compareBox = document.getElementById("compareBox");
const scannerModal = document.getElementById("scannerModal");
const scannerVideo = document.getElementById("scannerVideo");
const scannerClose = document.getElementById("scannerClose");

let xlsxWorker = null;
let scannerStream = null;
let scannerLoop = null;
let searchDebounce = null;

function getWorker() {
  if (!xlsxWorker) xlsxWorker = new Worker("/xlsx-worker.js");
  return xlsxWorker;
}

function setStatus(text) {
  statusBox.textContent = text;
}

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

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[char]));
}

function findHeaderIndex(header, name) {
  const wanted = normalize(name);
  return header.findIndex((cell) => normalize(cell) === wanted);
}

function formatQty(value) {
  return new Intl.NumberFormat("es-AR", { maximumFractionDigits: 2 }).format(value);
}

function formatDate(timestamp) {
  if (!timestamp) return "Sin fecha";
  return new Intl.DateTimeFormat("es-AR", {
    dateStyle: "short",
    timeStyle: "short"
  }).format(new Date(timestamp * 1000));
}

function applyPayload(payload, displayName, meta = null) {
  state.stockHeader = payload.header || [];
  state.stockRows = payload.rows || [];
  state.index = payload.index || {};
  state.inventory = payload.inventory || {};
  state.totals = payload.totals || {};
  state.baseStats = calculateBaseStats();
  state.lineas = payload.lineas || [];
  state.pendingByConcat = payload.pendingByConcat || {};
  state.meta = meta;

  fileName.textContent = displayName;
  updatedAt.textContent = meta?.uploadedAt
    ? `Actualizada: ${formatDate(meta.uploadedAt)}`
    : "Cargada en este navegador";
  rowCount.textContent = `${payload.productCount ?? state.stockRows.length} productos`;

  searchBtn.disabled = false;
  stockBtn.disabled = false;
  compareBtn.disabled = !state.serverMode;
  exportCopyBtn.disabled = false;
  exportCsvBtn.disabled = false;

  populateLineaFilter();
  populateBaseFilter();
  renderTotals();
  renderQuality();
  renderHistory();
  stockTable.innerHTML = "";
  state.stockVisible = false;
  stockSection.classList.add("hidden");
  stockBtn.textContent = "Stock completo";
  compareBox.innerHTML = "";
  resultBox.innerHTML = '<div class="empty">Escribí un código como 25551 NLSA M y tocá Buscar.</div>';
}

function populateLineaFilter() {
  const current = stockFilterLinea.value;
  stockFilterLinea.innerHTML = '<option value="">Todas las líneas</option>' +
    state.lineas.map((linea) => `<option value="${escapeHtml(linea)}">${escapeHtml(linea)}</option>`).join("");
  stockFilterLinea.value = current;
}

function populateBaseFilter() {
  const current = state.stockFilter.base;
  stockFilterBase.innerHTML = '<option value="">Todas las bases</option>' +
    BASE_COLUMNS.map((base) => `<option value="${escapeHtml(base.base)}">${escapeHtml(base.base)}</option>`).join("");
  stockFilterBase.value = current;
}

async function parseWithWorker(buffer) {
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

async function loadWorkbook() {
  const file = fileInput.files && fileInput.files[0];
  if (!file) return;

  searchBtn.disabled = true;
  stockBtn.disabled = true;
  resultBox.innerHTML = '<div class="empty">Leyendo la planilla...</div>';

  try {
    if (state.serverMode && location.protocol !== "file:") {
      await uploadAndLoadFromServer(file);
      return;
    }

    setStatus("Leyendo archivo Excel en segundo plano...");
    const buffer = await file.arrayBuffer();
    const payload = await parseWithWorker(buffer);
    applyPayload(payload, file.name);
    setStatus("Planilla cargada en el navegador.");
  } catch (error) {
    console.error(error);
    setStatus(`No pude leer la planilla: ${error.message}`);
    resultBox.innerHTML = '<div class="empty">Revisá que el archivo sea .xlsx y que tenga la pestaña STOCK.</div>';
  }
}

async function loadWorkbookLocally(file, note = "") {
  setStatus("Leyendo planilla en el navegador (puede tardar con archivos grandes)...");
  const buffer = await file.arrayBuffer();
  const payload = await parseWithWorker(buffer);
  applyPayload(payload, file.name);
  const suffix = note ? ` ${note}` : "";
  setStatus(`Planilla cargada en este navegador.${suffix}`);
}

async function uploadAndLoadFromServer(file) {
  const data = new FormData();
  data.append("stock", file, file.name);

  setStatus(`Subiendo ${file.name} al servidor (puede tardar 1-2 min, no cierres la página)...`);

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 180000);
    const upload = await fetch("/api/upload", {
      method: "POST",
      body: data,
      signal: controller.signal
    });
    clearTimeout(timeout);
    const uploadText = await upload.text();

    if (!upload.ok) {
      if (upload.status === 404) {
        await loadWorkbookLocally(file, "El servidor no tiene /api/upload. Redeploy con el Dockerfile nuevo.");
        return;
      }
      throw new Error(uploadText || `Error ${upload.status}`);
    }

    const meta = JSON.parse(uploadText);
    setStatus("Procesando planilla en el servidor...");
    const response = await fetch(`/api/stock-data?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) {
      await loadWorkbookLocally(file, "No hay datos procesados en el servidor.");
      return;
    }

    applyPayload(await response.json(), meta.name, meta);
    await loadHistory();
    setStatus(`Planilla cargada: ${meta.name} (${meta.productCount} productos).`);
  } catch (error) {
    resultBox.innerHTML = `<div class="empty">Error del servidor: ${escapeHtml(error.message)}</div>`;
    try {
      await loadWorkbookLocally(file, "Se usó lectura local porque el servidor falló.");
    } catch (localError) {
      throw new Error(`${error.message} | Local: ${localError.message}`);
    }
  }
}

async function checkServerVersion() {
  try {
    const response = await fetch("/api/version", { cache: "no-store" });
    if (!response.ok) {
      if (serverInfo) {
        serverInfo.textContent = "Servidor VIEJO: falta /api/version. En Portainer hacé rebuild con el Dockerfile Python.";
        serverInfo.style.color = "#b42318";
      }
      setStatus("Servidor antiguo detectado. La subida de planillas no va a funcionar hasta redeploy.");
      return null;
    }
    const version = await response.json();
    if (serverInfo) {
      serverInfo.textContent = `Servidor v${version.version} OK`;
      serverInfo.style.color = "";
    }
    return version;
  } catch {
    if (serverInfo) {
      serverInfo.textContent = "No hay servidor Python (¿abrís el HTML directo o Nginx viejo?)";
      serverInfo.style.color = "#b42318";
    }
    return null;
  }
}

async function loadServerWorkbook() {
  if (!state.serverMode || location.protocol === "file:") return false;

  const version = await checkServerVersion();

  try {
    const info = await fetch("/api/current", { cache: "no-store" });
    if (info.status === 404) {
      setStatus(version
        ? `Servidor v${version.version} listo. Subí la planilla para empezar.`
        : "No hay planilla guardada. Subí una para empezar.");
      return false;
    }
    if (!info.ok) throw new Error(await info.text());

    const meta = await info.json();
    setStatus(`Cargando datos procesados: ${meta.name}...`);
    const response = await fetch(`/api/stock-data?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error(await response.text());
    const payload = await response.json();
    applyPayload(payload, meta.name, meta);
    await loadHistory();
    setStatus(`Planilla del servidor cargada: ${meta.name}.`);
    return true;
  } catch (error) {
    setStatus(`No pude cargar la planilla del servidor: ${error.message}`);
    return false;
  }
}

async function loadHistory() {
  if (!state.serverMode) return;
  try {
    const response = await fetch("/api/history", { cache: "no-store" });
    if (!response.ok) return;
    state.history = await response.json();
    renderHistory();
  } catch {
    state.history = [];
  }
}

function renderHistory() {
  if (!state.serverMode) {
    historyBox.innerHTML = '<div class="empty">El historial está disponible cuando corrés la app en el servidor.</div>';
    return;
  }

  if (!state.history.length) {
    historyBox.innerHTML = '<div class="empty">Todavía no hay historial de planillas.</div>';
    return;
  }

  historyBox.innerHTML = state.history.map((entry, index) => `
    <div class="history-item">
      <div>
        <strong>${escapeHtml(entry.name)}</strong>
        <div class="muted">${formatDate(entry.uploadedAt)} · ${entry.productCount} productos</div>
      </div>
      ${index === 0 ? '<span class="pill">Actual</span>' : `<button class="link-btn" data-compare-id="${escapeHtml(entry.id)}">Comparar</button>`}
    </div>
  `).join("");

  historyBox.querySelectorAll("[data-compare-id]").forEach((button) => {
    button.addEventListener("click", () => compareWithHistory(button.dataset.compareId));
  });
}

async function compareWithHistory(historyId) {
  try {
    setStatus("Comparando planillas...");
    const response = await fetch(`/api/compare?id=${encodeURIComponent(historyId)}`, { cache: "no-store" });
    if (!response.ok) throw new Error(await response.text());
    const data = await response.json();
    renderComparison(data);
    setStatus(`Comparación lista contra ${data.previous.name}.`);
  } catch (error) {
    setStatus(`No pude comparar: ${error.message}`);
  }
}

async function compareWithPrevious() {
  try {
    setStatus("Comparando con la planilla anterior...");
    const response = await fetch("/api/compare", { cache: "no-store" });
    if (!response.ok) throw new Error(await response.text());
    const data = await response.json();
    renderComparison(data);
    setStatus(`Comparación lista contra ${data.previous.name}.`);
  } catch (error) {
    setStatus(`No pude comparar: ${error.message}`);
  }
}

function renderComparison(data) {
  const changes = data.changes || [];
  if (!changes.length) {
    compareBox.innerHTML = '<div class="empty">No hay diferencias de stock entre las dos planillas.</div>';
    return;
  }

  const limited = changes.slice(0, 300);
  compareBox.innerHTML = `
    <div class="compare-summary">
      <span class="pill">${data.summary.changedLines} cambios</span>
      <span class="pill up">+${data.summary.positive}</span>
      <span class="pill down">-${data.summary.negative}</span>
      <span class="file-name">Actual: ${escapeHtml(data.current.name)} vs ${escapeHtml(data.previous.name)}</span>
    </div>
    <div class="table-wrap compare-table">
      <table>
        <thead>
          <tr>
            <th>Producto</th>
            <th>Base</th>
            <th>Antes</th>
            <th>Ahora</th>
            <th>Delta</th>
          </tr>
        </thead>
        <tbody>
          ${limited.map((item) => `
            <tr>
              <td>${escapeHtml(item.concat)}</td>
              <td>${escapeHtml(item.base)}</td>
              <td class="num">${formatQty(item.before)}</td>
              <td class="num">${formatQty(item.after)}</td>
              <td class="num ${item.delta > 0 ? "up" : item.delta < 0 ? "down" : ""}">${formatQty(item.delta)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
    ${changes.length > limited.length ? `<p class="file-name">Mostrando ${limited.length} de ${changes.length} cambios.</p>` : ""}
  `;
}

function findSearchKey(query) {
  if (state.index[query]) return query;
  const partial = Object.keys(state.index).filter((key) => key.includes(query));
  if (partial.length === 1) return partial[0];
  if (partial.length > 1) {
    const exactStart = partial.find((key) => key.startsWith(query));
    return exactStart || partial[0];
  }
  return null;
}

function searchStock() {
  const query = normalize(searchInput.value);
  if (!query) {
    resultBox.innerHTML = '<div class="empty">Ingresá SKU color y talle en el casillero de búsqueda.</div>';
    return;
  }

  const key = findSearchKey(query);
  queryLabel.textContent = searchInput.value.trim();

  if (!key) {
    state.lastSearch = null;
    exportCopyBtn.disabled = true;
    exportCsvBtn.disabled = true;
    resultBox.innerHTML = '<div class="empty">No encontré ese SKU color talle en la planilla.</div>';
    return;
  }

  const product = state.index[key];
  const available = Object.entries(product.bases || {})
    .filter(([, qty]) => qty > 0)
    .map(([base, qty]) => ({ base, qty }));

  state.lastSearch = { key, product, available, query: searchInput.value.trim() };
  exportCopyBtn.disabled = false;
  exportCsvBtn.disabled = false;

  const productHtml = `
    <div class="product">
      <div class="kv"><span>SKU</span><strong>${escapeHtml(product.sku)}</strong></div>
      <div class="kv"><span>Color</span><strong>${escapeHtml(product.color)}</strong></div>
      <div class="kv"><span>Talle</span><strong>${escapeHtml(product.talle)}</strong></div>
      <div class="kv"><span>Nombre</span><strong>${escapeHtml(product.nombre)}</strong></div>
      <div class="kv"><span>Línea</span><strong>${escapeHtml(product.linea)}</strong></div>
    </div>`;

  if (!available.length) {
    resultBox.innerHTML = productHtml + '<div class="empty">El producto existe, pero todas las bases figuran con stock 0.</div>';
    return;
  }

  resultBox.innerHTML = productHtml + `
    <div class="cards">
      ${available.map((item) => `
        <button class="card card-button" type="button" data-base="${escapeHtml(item.base)}" title="Ver todos los productos con stock en ${escapeHtml(item.base)}">
          <div class="name">${escapeHtml(item.base)}</div>
          <div class="qty">${formatQty(item.qty)}</div>
        </button>
      `).join("")}
    </div>`;
}

function renderTotals() {
  const totals = Object.entries(state.totals || {});
  if (!totals.length) {
    totalsBox.innerHTML = "";
    return;
  }

  totalsBox.innerHTML = totals.map(([base, qty]) => `
    <button class="card card-button" type="button" data-base="${escapeHtml(base)}" title="Ver todos los productos con stock en ${escapeHtml(base)}">
      <div class="name">${escapeHtml(base)}</div>
      <div class="qty">${formatQty(qty)}</div>
    </button>
  `).join("");
}

function calculateBaseStats() {
  const stats = {};
  for (const base of BASE_COLUMNS) {
    const colIndex = findHeaderIndex(state.stockHeader, base.col);
    let positiveProducts = 0;
    let negativeCells = 0;
    let negativeQty = 0;
    for (const row of state.stockRows) {
      const qty = colIndex >= 0 ? asNumber(row[colIndex]) : 0;
      if (qty > 0) positiveProducts += 1;
      if (qty < 0) {
        negativeCells += 1;
        negativeQty += qty;
      }
    }
    stats[base.base] = { positiveProducts, negativeCells, negativeQty };
  }
  return stats;
}

function renderQuality() {
  const stats = Object.entries(state.baseStats || {});
  if (!stats.length) {
    qualityBox.innerHTML = '<div class="empty">Cargá una planilla para ver alertas y controles.</div>';
    return;
  }
  const negativeCells = stats.reduce((sum, [, stat]) => sum + stat.negativeCells, 0);
  const activeBases = stats.filter(([, stat]) => stat.positiveProducts > 0).length;
  const negativeBases = stats.filter(([, stat]) => stat.negativeCells > 0);
  qualityBox.innerHTML = `
    <div class="alert-grid">
      <div class="alert-card"><div class="label">Productos cargados</div><div class="value">${formatQty(state.stockRows.length)}</div></div>
      <div class="alert-card"><div class="label">Bases activas</div><div class="value">${formatQty(activeBases)}</div></div>
      <div class="alert-card ${negativeCells ? "danger" : ""}"><div class="label">Registros negativos</div><div class="value">${formatQty(negativeCells)}</div></div>
    </div>
    <div class="base-health" style="margin-top:16px">
      ${negativeBases.length
        ? negativeBases.map(([base, stat]) => `<button class="base-health-row link-btn" data-negative-base="${escapeHtml(base)}"><span>${escapeHtml(base)}: ${formatQty(stat.negativeCells)} registro(s) negativo(s)</span><span>Ver detalle</span></button>`).join("")
        : '<div class="empty">No hay stock negativo en las bases cargadas.</div>'}
    </div>`;
}

function toggleStock() {
  state.stockVisible = !state.stockVisible;
  if (state.stockVisible) {
    stockSection.classList.remove("hidden");
    stockBtn.textContent = "Ocultar stock";
    renderStockTable();
  } else {
    stockSection.classList.add("hidden");
    stockBtn.textContent = "Stock completo";
  }
}

function rowHasStock(row) {
  for (const base of BASE_COLUMNS) {
    const colIndex = findHeaderIndex(state.stockHeader, base.col);
    if (colIndex >= 0 && asNumber(row[colIndex]) > 0) return true;
  }
  return false;
}

function getFilteredRows() {
  const concatCol = findHeaderIndex(state.stockHeader, "CONCAT");
  const skuCol = findHeaderIndex(state.stockHeader, "SKU");
  const lineaCol = findHeaderIndex(state.stockHeader, "LINEA");
  const text = normalize(state.stockFilter.text);
  const linea = normalize(state.stockFilter.linea);
  const base = BASE_COLUMNS.find((item) => item.base === state.stockFilter.base);
  const baseCol = base ? findHeaderIndex(state.stockHeader, base.col) : -1;
  const negativeBase = BASE_COLUMNS.find((item) => item.base === state.stockFilter.negativeBase);
  const negativeCol = negativeBase ? findHeaderIndex(state.stockHeader, negativeBase.col) : -1;

  return state.stockRows.filter((row) => {
    if (state.stockFilter.onlyStock && !rowHasStock(row)) return false;
    if (base && (baseCol < 0 || asNumber(row[baseCol]) <= 0)) return false;
    if (negativeBase && (negativeCol < 0 || asNumber(row[negativeCol]) >= 0)) return false;
    if (linea && normalize(row[lineaCol]) !== linea) return false;
    if (!text) return true;
    const concat = concatCol >= 0 ? normalize(row[concatCol]) : "";
    const sku = skuCol >= 0 ? normalize(row[skuCol]) : "";
    return concat.includes(text) || sku.includes(text);
  });
}

function showBaseStock(base) {
  if (!BASE_COLUMNS.some((item) => item.base === base)) return;
  state.stockFilter.base = base;
  state.stockFilter.negativeBase = "";
  stockFilterBase.value = base;
  state.stockFilter.onlyStock = true;
  stockFilterOnly.checked = true;
  state.virtualScrollTop = 0;
  state.stockVisible = true;
  stockSection.classList.remove("hidden");
  stockBtn.textContent = "Ocultar stock";
  renderStockTable();
  stockSection.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showNegativeBaseStock(base) {
  const item = BASE_COLUMNS.find((entry) => entry.base === base);
  if (!item) return;
  const colIndex = findHeaderIndex(state.stockHeader, item.col);
  if (colIndex < 0) return;
  state.stockFilter.base = "";
  state.stockFilter.negativeBase = base;
  stockFilterBase.value = "";
  state.stockFilter.onlyStock = false;
  stockFilterOnly.checked = false;
  state.stockFilter.text = "";
  stockFilterText.value = "";
  state.stockVisible = true;
  stockSection.classList.remove("hidden");
  stockBtn.textContent = "Ocultar stock";
  state.virtualScrollTop = 0;
  renderStockTable();
  stockSection.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderStockTable() {
  const header = state.stockHeader;
  const rows = getFilteredRows();
  const negativeLabel = state.stockFilter.negativeBase ? ` · negativos en ${state.stockFilter.negativeBase}` : "";
  stockInfo.textContent = `${rows.length} filas visibles${negativeLabel}`;

  const viewportHeight = stockTable.clientHeight || 480;
  const visibleCount = Math.ceil(viewportHeight / VIRTUAL_ROW_HEIGHT) + VIRTUAL_OVERSCAN * 2;
  const start = Math.max(0, Math.floor(state.virtualScrollTop / VIRTUAL_ROW_HEIGHT) - VIRTUAL_OVERSCAN);
  const end = Math.min(rows.length, start + visibleCount);
  const topSpacer = start * VIRTUAL_ROW_HEIGHT;
  const bottomSpacer = Math.max(0, (rows.length - end) * VIRTUAL_ROW_HEIGHT);

  stockTable.innerHTML = `
    <table>
      <thead>
        <tr>${header.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr>
      </thead>
      <tbody>
        <tr class="spacer"><td colspan="${header.length}" style="height:${topSpacer}px;padding:0;border:0"></td></tr>
        ${rows.slice(start, end).map((row) => `
          <tr style="height:${VIRTUAL_ROW_HEIGHT}px">
            ${header.map((_, index) => {
              const value = row[index] ?? "";
              const numeric = typeof value === "number" || /^-?\d+(\.\d+)?$/.test(String(value));
              return `<td class="${numeric ? "num" : ""}">${escapeHtml(value)}</td>`;
            }).join("")}
          </tr>
        `).join("")}
        <tr class="spacer"><td colspan="${header.length}" style="height:${bottomSpacer}px;padding:0;border:0"></td></tr>
      </tbody>
    </table>
  `;
}

function exportSearchCopy() {
  if (!state.lastSearch) return;
  const lines = [
    `Producto: ${state.lastSearch.query}`,
    `SKU: ${state.lastSearch.product.sku}`,
    `Color: ${state.lastSearch.product.color}`,
    `Talle: ${state.lastSearch.product.talle}`,
    `Nombre: ${state.lastSearch.product.nombre}`,
    ""
  ];
  for (const item of state.lastSearch.available) {
    lines.push(`${item.base}: ${formatQty(item.qty)}`);
  }
  navigator.clipboard.writeText(lines.join("\n")).then(() => {
    setStatus("Resultado copiado al portapapeles.");
  }).catch(() => {
    setStatus("No pude copiar al portapapeles.");
  });
}

function exportSearchCsv() {
  if (!state.lastSearch) return;
  const rows = [
    ["campo", "valor"],
    ["consulta", state.lastSearch.query],
    ["sku", state.lastSearch.product.sku],
    ["color", state.lastSearch.product.color],
    ["talle", state.lastSearch.product.talle],
    ["nombre", state.lastSearch.product.nombre],
    ["linea", state.lastSearch.product.linea],
    [],
    ["base", "cantidad"]
  ];
  for (const item of state.lastSearch.available) {
    rows.push([item.base, item.qty]);
  }
  const csv = rows.map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `stock-${normalize(state.lastSearch.query).replace(/\s+/g, "-")}.csv`;
  link.click();
  URL.revokeObjectURL(url);
  setStatus("CSV descargado.");
}

function toggleTheme() {
  const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
  document.documentElement.dataset.theme = next;
  localStorage.setItem("stock-theme", next);
  themeBtn.textContent = next === "dark" ? "Modo claro" : "Modo oscuro";
}

function applySavedTheme() {
  const saved = localStorage.getItem("stock-theme");
  if (saved === "dark") {
    document.documentElement.dataset.theme = "dark";
    themeBtn.textContent = "Modo claro";
  }
}

async function startScanner() {
  if (!("BarcodeDetector" in window)) {
    setStatus("Tu navegador no soporta escaneo por cámara. Usá Chrome en el celular.");
    return;
  }

  try {
    scannerStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: "environment" } },
      audio: false
    });
    scannerVideo.srcObject = scannerStream;
    await scannerVideo.play();
    scannerModal.classList.remove("hidden");

    const detector = new BarcodeDetector({ formats: ["code_128", "code_39", "ean_13", "ean_8", "qr_code"] });
    scannerLoop = window.setInterval(async () => {
      try {
        const codes = await detector.detect(scannerVideo);
        if (!codes.length) return;
        const raw = codes[0].rawValue || "";
        if (!raw) return;
        searchInput.value = raw.trim();
        stopScanner();
        searchStock();
        setStatus(`Código escaneado: ${raw}`);
      } catch {
        // ignore intermittent detect errors
      }
    }, 350);
  } catch (error) {
    setStatus(`No pude abrir la cámara: ${error.message}`);
  }
}

function stopScanner() {
  scannerModal.classList.add("hidden");
  if (scannerLoop) {
    clearInterval(scannerLoop);
    scannerLoop = null;
  }
  if (scannerStream) {
    scannerStream.getTracks().forEach((track) => track.stop());
    scannerStream = null;
  }
  scannerVideo.srcObject = null;
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator) || location.protocol === "file:") return;
  navigator.serviceWorker.register("/sw.js").catch(() => {});
}

fileInput.addEventListener("change", loadWorkbook);
searchBtn.addEventListener("click", searchStock);
stockBtn.addEventListener("click", toggleStock);
scanBtn.addEventListener("click", startScanner);
scannerClose.addEventListener("click", stopScanner);
themeBtn.addEventListener("click", toggleTheme);
exportCopyBtn.addEventListener("click", exportSearchCopy);
exportCsvBtn.addEventListener("click", exportSearchCsv);
compareBtn.addEventListener("click", compareWithPrevious);
searchInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !searchBtn.disabled) searchStock();
});
searchInput.addEventListener("input", () => {
  clearTimeout(searchDebounce);
  searchDebounce = setTimeout(() => {
    if (!searchBtn.disabled && normalize(searchInput.value).length >= 3) searchStock();
  }, 300);
});
stockFilterText.addEventListener("input", () => {
  state.stockFilter.text = stockFilterText.value;
  if (state.stockVisible) renderStockTable();
});
stockFilterLinea.addEventListener("change", () => {
  state.stockFilter.linea = stockFilterLinea.value;
  if (state.stockVisible) renderStockTable();
});
stockFilterBase.addEventListener("change", () => {
  state.stockFilter.base = stockFilterBase.value;
  state.stockFilter.negativeBase = "";
  state.virtualScrollTop = 0;
  if (state.stockVisible) renderStockTable();
});
stockFilterOnly.addEventListener("change", () => {
  state.stockFilter.onlyStock = stockFilterOnly.checked;
  if (state.stockVisible) renderStockTable();
});
totalsBox.addEventListener("click", (event) => {
  const button = event.target.closest("[data-base]");
  if (button) showBaseStock(button.dataset.base);
});
resultBox.addEventListener("click", (event) => {
  const button = event.target.closest("[data-base]");
  if (button) showBaseStock(button.dataset.base);
});
qualityBox.addEventListener("click", (event) => {
  const button = event.target.closest("[data-negative-base]");
  if (button) showNegativeBaseStock(button.dataset.negativeBase);
});
stockTable.addEventListener("scroll", () => {
  state.virtualScrollTop = stockTable.scrollTop;
  if (state.stockVisible) renderStockTable();
});

applySavedTheme();
registerServiceWorker();
exportCopyBtn.disabled = true;
exportCsvBtn.disabled = true;
compareBtn.disabled = true;

if (!window.__stockAppLoaded) {
  document.getElementById("status").textContent =
    "No se cargó app.js. En Portainer tenés que usar el Dockerfile Python nuevo (no solo Nginx).";
} else {
  loadServerWorkbook().then((loaded) => {
    if (!loaded) checkServerVersion();
  });
}
