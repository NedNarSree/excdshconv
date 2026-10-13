const PREVIEW_LIMIT = 100;
const ACCENT = "#0f766e";
const SLICE_COLORS = [
  "#0f766e",
  "#14b8a6",
  "#0ea5e9",
  "#6366f1",
  "#f59e0b",
  "#ef4444",
  "#8b5cf6",
  "#84cc16",
];

const state = {
  workbook: null,
  fileName: "",
  columns: [],
  rows: [],
  numbers: [],
  categories: [],
  dates: [],
  title: "",
  filters: {},
};

const charts = { bar: null, doughnut: null, line: null };

const el = (id) => document.getElementById(id);

function showScreen(id) {
  ["screen-upload", "screen-setup", "screen-dashboard"].forEach((name) => {
    el(name).hidden = name !== id;
  });
}

function showUploadError() {
  el("upload-error").hidden = false;
  el("sheet-picker").hidden = true;
  el("file-name").hidden = true;
  el("file-input").value = "";
  state.workbook = null;
  showScreen("screen-upload");
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (value instanceof Date) return null;
  let text = String(value).trim();
  if (!text) return null;
  let negative = false;
  if (/^\(.*\)$/.test(text)) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  const cleaned = text.replace(/[$€£¥₹,\s%]/g, "");
  if (!/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(cleaned)) return null;
  const n = Number(cleaned);
  if (!Number.isFinite(n)) return null;
  return negative ? -Math.abs(n) : n;
}

function parseDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number") return null;
  const text = String(value).trim();
  if (!text || parseNumber(text) !== null) return null;

  const isoMatch = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[T\s].*)?$/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(year, month - 1, day);
    }
    return null;
  }

  const slashMatch = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:\s.*)?$/);
  if (slashMatch) {
    const a = Number(slashMatch[1]);
    const b = Number(slashMatch[2]);
    let year = Number(slashMatch[3]);
    if (year < 100) year += year >= 70 ? 1900 : 2000;
    const month = a <= 12 && b > 12 ? a : b <= 12 && a > 12 ? b : a;
    const day = month === a ? b : a;
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(year, month - 1, day);
    }
    return null;
  }

  if (
    /\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\b/i.test(text) &&
    /\b\d{4}\b/.test(text)
  ) {
    const parsed = Date.parse(text);
    if (!Number.isNaN(parsed)) {
      const d = new Date(parsed);
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
  }
  return null;
}

function looksNumeric(values) {
  let ok = 0;
  let seen = 0;
  for (const v of values) {
    if (v === null || v === undefined || String(v).trim() === "") continue;
    seen += 1;
    if (parseNumber(v) !== null && !(v instanceof Date)) ok += 1;
  }
  return seen > 0 && ok / seen >= 0.7;
}

function looksDate(values) {
  let ok = 0;
  let seen = 0;
  for (const v of values) {
    if (v === null || v === undefined || String(v).trim() === "") continue;
    seen += 1;
    if (parseDate(v)) ok += 1;
  }
  return seen > 0 && ok / seen >= 0.7;
}

function formatNumber(n) {
  if (n === null || n === undefined || Number.isNaN(n)) return "—";
  const options = Number.isInteger(n)
    ? { maximumFractionDigits: 0 }
    : { minimumFractionDigits: 0, maximumFractionDigits: 2 };
  return new Intl.NumberFormat(undefined, options).format(n);
}

function displayCell(value) {
  if (value === null || value === undefined || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return String(value);
}

function sheetToTable(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: true,
    dateNF: "yyyy-mm-dd",
  });
  if (!rows.length) return null;
  const columns = Object.keys(rows[0]).filter((name) => {
    const trimmed = String(name).trim();
    return trimmed !== "" && !/^__EMPTY/.test(trimmed);
  });
  if (!columns.length) return null;
  const hasValue = rows.some((row) =>
    columns.some((col) => {
      const v = row[col];
      return v !== null && v !== undefined && String(v).trim() !== "";
    })
  );
  if (!hasValue) return null;
  return { columns, rows };
}

function guessKinds(columns, rows) {
  const sample = rows.slice(0, 80);
  const kinds = {};
  for (const col of columns) {
    const values = sample.map((r) => r[col]);
    if (looksDate(values)) kinds[col] = "date";
    else if (looksNumeric(values)) kinds[col] = "number";
    else kinds[col] = "category";
  }
  return kinds;
}

function checkedValues(containerId) {
  return [...document.querySelectorAll("#" + containerId + " input:checked")].map(
    (input) => input.value
  );
}

function fillCheckList(containerId, columns, kinds, kind) {
  const box = el(containerId);
  box.innerHTML = columns
    .map((col) => {
      const checked = kinds[col] === kind ? "checked" : "";
      return `<label><input type="checkbox" value="${escapeAttr(
        col
      )}" ${checked} /><span>${escapeHtml(col)}</span></label>`;
    })
    .join("");
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeAttr(text) {
  return escapeHtml(text);
}

function fillSelect(select, options, extra) {
  const extras = extra ? extra : [];
  select.innerHTML = extras
    .concat(options)
    .map((opt) => `<option value="${escapeAttr(opt.value)}">${escapeHtml(opt.label)}</option>`)
    .join("");
}

function uniqueValues(rows, col) {
  const seen = new Set();
  const list = [];
  for (const row of rows) {
    const key = displayCell(row[col]);
    const store = key === "" ? "(blank)" : key;
    if (!seen.has(store)) {
      seen.add(store);
      list.push(store);
    }
  }
  list.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  return list;
}

function applyFilters(rows) {
  return rows.filter((row) => {
    for (const [col, wanted] of Object.entries(state.filters)) {
      if (!wanted) continue;
      const key = displayCell(row[col]);
      const store = key === "" ? "(blank)" : key;
      if (store !== wanted) return false;
    }
    return true;
  });
}

function sumBy(rows, groupCol, valueCol) {
  const totals = new Map();
  for (const row of rows) {
    const label = displayCell(row[groupCol]);
    const key = label === "" ? "(blank)" : label;
    const amount =
      valueCol === "__count" ? 1 : parseNumber(row[valueCol]) || 0;
    totals.set(key, (totals.get(key) || 0) + amount);
  }
  return [...totals.entries()].sort((a, b) => b[1] - a[1]);
}

function trendBy(rows, dateCol, valueCol) {
  const points = [];
  for (const row of rows) {
    const d = parseDate(row[dateCol]);
    if (!d) continue;
    points.push({
      t: d.getTime(),
      d,
      amount: valueCol === "__count" ? 1 : parseNumber(row[valueCol]) || 0,
    });
  }
  if (!points.length) return [];
  points.sort((a, b) => a.t - b.t);
  const span = points[points.length - 1].t - points[0].t;
  const useMonth = span > 1000 * 60 * 60 * 24 * 60;
  const totals = new Map();
  for (const p of points) {
    const key = useMonth
      ? `${p.d.getFullYear()}-${String(p.d.getMonth() + 1).padStart(2, "0")}`
      : `${p.d.getFullYear()}-${String(p.d.getMonth() + 1).padStart(2, "0")}-${String(
          p.d.getDate()
        ).padStart(2, "0")}`;
    totals.set(key, (totals.get(key) || 0) + p.amount);
  }
  return [...totals.entries()].map(([label, value]) => ({ label, value }));
}

function destroyChart(name) {
  if (charts[name]) {
    charts[name].destroy();
    charts[name] = null;
  }
}

function valueOptions() {
  const opts = state.numbers.map((col) => ({ value: col, label: col }));
  opts.push({ value: "__count", label: "Number of rows" });
  return opts;
}

function valueLabel(value) {
  return value === "__count" ? "Number of rows" : "Total of " + value;
}

function renderKpis(rows) {
  const box = el("kpi-row");
  if (!state.numbers.length) {
    box.hidden = true;
    box.innerHTML = "";
    return;
  }
  box.hidden = false;
  box.innerHTML = state.numbers
    .flatMap((col) => {
      const nums = rows
        .map((r) => parseNumber(r[col]))
        .filter((n) => n !== null);
      const total = nums.reduce((a, b) => a + b, 0);
      const avg = nums.length ? total / nums.length : 0;
      return [
        kpiCard("Total of " + col, formatNumber(total)),
        kpiCard("Average of " + col, formatNumber(avg)),
      ];
    })
    .join("");
}

function kpiCard(label, value) {
  return `<div class="kpi"><div class="label">${escapeHtml(
    label
  )}</div><div class="value">${escapeHtml(value)}</div></div>`;
}

function renderFilters() {
  const wrap = el("filter-row");
  const controls = el("filter-controls");
  if (!state.categories.length) {
    wrap.hidden = true;
    controls.innerHTML = "";
    state.filters = {};
    return;
  }
  wrap.hidden = false;
  controls.innerHTML = state.categories
    .map((col) => {
      const options = uniqueValues(state.rows, col)
        .map(
          (v) =>
            `<option value="${escapeAttr(v)}"${
              state.filters[col] === v ? " selected" : ""
            }>${escapeHtml(v)}</option>`
        )
        .join("");
      return `<label>${escapeHtml(
        col
      )}<select data-filter="${escapeAttr(
        col
      )}"><option value="">All</option>${options}</select></label>`;
    })
    .join("");
}

function renderTable(rows) {
  const table = el("preview-table");
  const caption = el("table-caption");
  const cols = state.columns;
  const shown = rows.slice(0, PREVIEW_LIMIT);
  caption.textContent =
    rows.length > PREVIEW_LIMIT
      ? `Showing ${PREVIEW_LIMIT} of ${rows.length} rows`
      : `${rows.length} row${rows.length === 1 ? "" : "s"}`;
  if (!cols.length) {
    table.innerHTML = "";
    return;
  }
  const head = `<thead><tr>${cols
    .map((c) => `<th>${escapeHtml(c)}</th>`)
    .join("")}</tr></thead>`;
  const body = shown
    .map(
      (row) =>
        `<tr>${cols
          .map((c) => `<td>${escapeHtml(displayCell(row[c]))}</td>`)
          .join("")}</tr>`
    )
    .join("");
  table.innerHTML = head + `<tbody>${body}</tbody>`;
}

function renderBar(rows) {
  const canShow = state.categories.length > 0;
  el("card-bar").hidden = !canShow;
  if (!canShow) {
    destroyChart("bar");
    return;
  }
  const group = el("bar-group").value;
  const value = el("bar-value").value;
  el("bar-subtitle").textContent = `${valueLabel(value)} by ${group}`;
  const data = sumBy(rows, group, value).slice(0, 20);
  destroyChart("bar");
  charts.bar = new Chart(el("chart-bar"), {
    type: "bar",
    data: {
      labels: data.map((d) => d[0]),
      datasets: [
        {
          label: valueLabel(value),
          data: data.map((d) => d[1]),
          backgroundColor: ACCENT,
          borderRadius: 6,
        },
      ],
    },
    options: chartOptions(),
  });
}

function renderDoughnut(rows) {
  const canShow = state.categories.length > 0;
  el("card-doughnut").hidden = !canShow;
  if (!canShow) {
    destroyChart("doughnut");
    return;
  }
  const group = el("doughnut-group").value;
  const value = el("doughnut-value").value;
  el("doughnut-subtitle").textContent = `${valueLabel(value)} by ${group}`;
  const all = sumBy(rows, group, value).filter((d) => d[1] > 0);
  let data = all;
  if (all.length > 8) {
    const top = all.slice(0, 7);
    const otherTotal = all.slice(7).reduce((sum, d) => sum + d[1], 0);
    data = [...top, ["Other", otherTotal]];
  }
  destroyChart("doughnut");
  charts.doughnut = new Chart(el("chart-doughnut"), {
    type: "doughnut",
    data: {
      labels: data.map((d) => d[0]),
      datasets: [
        {
          data: data.map((d) => d[1]),
          backgroundColor: data.map((_, i) => SLICE_COLORS[i % SLICE_COLORS.length]),
        },
      ],
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom" } },
    },
  });
}

function renderLine(rows) {
  const canShow = state.dates.length > 0;
  el("card-line").hidden = !canShow;
  if (!canShow) {
    destroyChart("line");
    return;
  }
  const dateCol = el("line-date").value;
  const value = el("line-value").value;
  el("line-subtitle").textContent = `${valueLabel(value)} over ${dateCol}`;
  const data = trendBy(rows, dateCol, value);
  destroyChart("line");
  charts.line = new Chart(el("chart-line"), {
    type: "line",
    data: {
      labels: data.map((d) => d.label),
      datasets: [
        {
          label: valueLabel(value),
          data: data.map((d) => d.value),
          borderColor: ACCENT,
          backgroundColor: "rgba(15, 118, 110, 0.12)",
          fill: true,
          tension: 0.25,
          pointRadius: 3,
        },
      ],
    },
    options: chartOptions(),
  });
}

function chartOptions() {
  return {
    animation: false,
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      x: { grid: { display: false } },
      y: { beginAtZero: true, ticks: { precision: 0 } },
    },
  };
}

function renderDashboard() {
  const rows = applyFilters(state.rows);
  const title = state.title.trim();
  el("dashboard-title").hidden = !title;
  el("dashboard-title").textContent = title;
  el("charts-wrap").hidden = !state.categories.length && !state.dates.length;
  renderKpis(rows);
  renderTable(rows);
  renderBar(rows);
  renderDoughnut(rows);
  renderLine(rows);
}

function openSetup(table) {
  state.columns = table.columns;
  state.rows = table.rows;
  const kinds = guessKinds(table.columns, table.rows);
  fillCheckList("cols-numbers", table.columns, kinds, "number");
  fillCheckList("cols-categories", table.columns, kinds, "category");
  fillCheckList("cols-dates", table.columns, kinds, "date");
  el("setup-error").hidden = true;
  showScreen("screen-setup");
}

function openDashboard() {
  const numbers = checkedValues("cols-numbers");
  const categories = checkedValues("cols-categories");
  const dates = checkedValues("cols-dates");
  if (!numbers.length && !categories.length && !dates.length) {
    el("setup-error").hidden = false;
    return;
  }
  el("setup-error").hidden = true;
  state.numbers = numbers;
  state.categories = categories;
  state.dates = dates;
  state.title = el("dash-title").value;
  state.filters = {};

  const values = valueOptions();
  fillSelect(el("bar-group"), categories.map((c) => ({ value: c, label: c })));
  fillSelect(
    el("doughnut-group"),
    categories.map((c) => ({ value: c, label: c }))
  );
  if (categories.length > 1) {
    el("doughnut-group").value = categories[1];
  }
  fillSelect(el("bar-value"), values);
  fillSelect(el("doughnut-value"), values);
  fillSelect(el("line-date"), dates.map((c) => ({ value: c, label: c })));
  fillSelect(el("line-value"), values);

  renderFilters();
  showScreen("screen-dashboard");
  renderDashboard();
}

async function readWorkbook(file) {
  if (!file || file.size === 0) {
    throw new Error("Empty file");
  }
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".csv")) {
    const text = await file.text();
    if (!text.trim() || text.includes("\0")) {
      throw new Error("Invalid CSV");
    }
    return XLSX.read(text, { type: "string", raw: true });
  }
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error("Invalid XLSX");
  }
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

el("file-input").addEventListener("change", async (event) => {
  const file = event.target.files && event.target.files[0];
  el("upload-error").hidden = true;
  el("sheet-picker").hidden = true;
  if (!file) return;
  const lower = file.name.toLowerCase();
  if (!lower.endsWith(".xlsx") && !lower.endsWith(".csv")) {
    showUploadError();
    return;
  }
  try {
    const workbook = await readWorkbook(file);
    const sheets = workbook.SheetNames || [];
    if (!sheets.length) {
      showUploadError();
      return;
    }
    state.workbook = workbook;
    state.fileName = file.name;
    if (sheets.length === 1) {
      const table = sheetToTable(workbook, sheets[0]);
      if (!table) {
        showUploadError();
        return;
      }
      el("file-name").hidden = false;
      el("file-name").textContent = file.name;
      openSetup(table);
      return;
    }
    const hasReadableSheet = sheets.some((name) => sheetToTable(workbook, name) !== null);
    if (!hasReadableSheet) {
      showUploadError();
      return;
    }
    el("file-name").hidden = false;
    el("file-name").textContent = file.name;
    const select = el("sheet-select");
    select.innerHTML = sheets
      .map((name) => `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`)
      .join("");
    el("sheet-picker").hidden = false;
  } catch (err) {
    showUploadError();
  }
});

el("sheet-continue").addEventListener("click", () => {
  if (!state.workbook) {
    showUploadError();
    return;
  }
  const table = sheetToTable(state.workbook, el("sheet-select").value);
  if (!table) {
    showUploadError();
    return;
  }
  openSetup(table);
});

el("setup-back").addEventListener("click", () => {
  el("file-input").value = "";
  el("file-name").hidden = true;
  el("sheet-picker").hidden = true;
  el("upload-error").hidden = true;
  state.workbook = null;
  showScreen("screen-upload");
});

el("setup-continue").addEventListener("click", openDashboard);

el("start-over").addEventListener("click", () => {
  el("file-input").value = "";
  el("file-name").hidden = true;
  el("sheet-picker").hidden = true;
  el("upload-error").hidden = true;
  el("dash-title").value = "";
  state.workbook = null;
  destroyChart("bar");
  destroyChart("doughnut");
  destroyChart("line");
  showScreen("screen-upload");
});

el("filter-controls").addEventListener("change", (event) => {
  const select = event.target.closest("select[data-filter]");
  if (!select) return;
  state.filters[select.getAttribute("data-filter")] = select.value;
  renderDashboard();
});

["bar-group", "bar-value", "doughnut-group", "doughnut-value", "line-date", "line-value"].forEach(
  (id) => {
    el(id).addEventListener("change", renderDashboard);
  }
);

el("export-pdf").addEventListener("click", async () => {
  const button = el("export-pdf");
  button.disabled = true;
  document.body.classList.add("exporting");
  try {
    const node = el("dashboard-export");
    const canvas = await html2canvas(node, {
      scale: 2,
      backgroundColor: "#f3f5f7",
      useCORS: true,
    });
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 24;
    const usableWidth = pageWidth - margin * 2;
    const usableHeight = pageHeight - margin * 2;
    const scale = Math.min(usableWidth / canvas.width, usableHeight / canvas.height);
    const imgWidth = canvas.width * scale;
    const imgHeight = canvas.height * scale;
    const x = margin + (usableWidth - imgWidth) / 2;
    const y = margin;
    const img = canvas.toDataURL("image/png");
    pdf.addImage(img, "PNG", x, y, imgWidth, imgHeight);
    const name = (state.title.trim() || "dashboard").replace(/[^\w\-]+/g, "_");
    pdf.save(name + ".pdf");
  } finally {
    document.body.classList.remove("exporting");
    button.disabled = false;
  }
});
