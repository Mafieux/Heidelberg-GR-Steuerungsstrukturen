"""
build_webapp.py
================
Generator für den Heidelberg-Haushaltsdaten-Browser (Single-HTML-App).

Liest:  C:\\Users\\D061012\\Desktop\\HD Haushalt\\Daten\\01_HD_Haushaltsdaten.xlsx
Schreibt: C:\\Users\\D061012\\Desktop\\HD Haushalt\\WebApp\\index.html

Aufruf: python WebApp/build_webapp.py
"""

from pathlib import Path
from openpyxl import load_workbook
from datetime import datetime
import json
import sys

# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------

BASE      = Path(r"C:\Users\D061012\Desktop\HD Haushalt")
XLSX      = BASE / "Daten" / "01_HD_Haushaltsdaten.xlsx"
OUT_HTML  = BASE / "WebApp" / "index.html"

POPULATED_SHEETS = [
    "00_README",
    "01_Eckdaten",
    "02_Teilhaushalte",
    "08_Schulden_Liquiditaet",
    "09_Kennzahlen",
    "13_Mapping_TH_PB",
    "14_Stammdaten_TH",
]

SHEET_META = {
    "00_README":              {"label": "00 · README",            "icon": "📖"},
    "01_Eckdaten":            {"label": "01 · Eckdaten",          "icon": "📊"},
    "02_Teilhaushalte":       {"label": "02 · Teilhaushalte",     "icon": "🏛️"},
    "08_Schulden_Liquiditaet":{"label": "08 · Schulden",          "icon": "💰"},
    "09_Kennzahlen":          {"label": "09 · Kennzahlen",        "icon": "📐"},
    "13_Mapping_TH_PB":       {"label": "13 · Mapping TH⇄PB",     "icon": "🔗"},
    "14_Stammdaten_TH":       {"label": "14 · Stammdaten TH",     "icon": "🗂️"},
}

# ---------------------------------------------------------------------------
# Excel-Reader
# ---------------------------------------------------------------------------

def read_sheet(ws):
    """Liest ein Worksheet zu (headers, rows) als list/list-of-list."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    headers = [str(h) if h is not None else "" for h in rows[0]]
    data_rows = []
    for r in rows[1:]:
        # Komplett leere Zeilen überspringen
        if all(cell is None or (isinstance(cell, str) and cell.strip() == "") for cell in r):
            continue
        data_rows.append(list(r))
    return headers, data_rows


def read_readme(ws):
    """00_README ist Key-Value, kein Tabellenformat."""
    pairs = []
    for r in ws.iter_rows(values_only=True):
        if r and r[0]:
            label = str(r[0])
            value = "" if (len(r) < 2 or r[1] is None) else str(r[1])
            pairs.append({"label": label, "value": value})
    return pairs


def read_kennzahlen(ws):
    """09_Kennzahlen: nur Zeilen mit 4-stelliger Jahreszahl als Daten,
    Rest als Hinweise."""
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h) if h is not None else "" for h in rows[0]]
    data_rows = []
    hinweise = []
    for r in rows[1:]:
        if r and r[0] is not None:
            # Versuche, Jahr zu erkennen
            try:
                yr = int(r[0])
                if 2000 <= yr <= 2050:
                    data_rows.append(list(r))
                    continue
            except (ValueError, TypeError):
                pass
            # Sonst: Text-Zeile als Hinweis sammeln
            text = " ".join(str(c) for c in r if c is not None).strip()
            if text:
                hinweise.append(text)
    return headers, data_rows, hinweise


# ---------------------------------------------------------------------------
# Hauptlogik: Daten aufbereiten
# ---------------------------------------------------------------------------

def build_data():
    if not XLSX.exists():
        sys.exit(f"[FEHLER] Excel-Datei nicht gefunden: {XLSX}")
    wb = load_workbook(XLSX, data_only=True)

    sheets = {}
    for name in POPULATED_SHEETS:
        if name not in wb.sheetnames:
            print(f"[WARNUNG] Sheet '{name}' nicht im Workbook — übersprungen.")
            continue
        ws = wb[name]
        if name == "00_README":
            sheets[name] = {
                "type": "kv",
                "pairs": read_readme(ws),
                "label": SHEET_META[name]["label"],
                "icon": SHEET_META[name]["icon"],
            }
        elif name == "09_Kennzahlen":
            headers, data_rows, hinweise = read_kennzahlen(ws)
            sheets[name] = {
                "type": "table",
                "headers": headers,
                "rows": data_rows,
                "hinweise": hinweise,
                "label": SHEET_META[name]["label"],
                "icon": SHEET_META[name]["icon"],
            }
        else:
            headers, data_rows = read_sheet(ws)
            sheets[name] = {
                "type": "table",
                "headers": headers,
                "rows": data_rows,
                "label": SHEET_META[name]["label"],
                "icon": SHEET_META[name]["icon"],
            }

    return {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "sheets": sheets,
    }


# ---------------------------------------------------------------------------
# HTML-Template
# ---------------------------------------------------------------------------

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Stadt Heidelberg — Haushaltsdaten 2017–2029</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
:root {
  --sap-blue: #2F5496;
  --sap-blue-dark: #1F3864;
  --subheader: #D9E1F2;
  --neg: #C00000;
  --pos: #2E7D32;
  --bg: #FFFFFF;
  --bg-soft: #F7F8FA;
  --muted: #6B7280;
  --border: #E5E7EB;
  --text: #1F2937;
}
* { box-sizing: border-box; }
html, body {
  margin: 0; padding: 0;
  font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  font-size: 14px; line-height: 1.5;
  color: var(--text); background: var(--bg-soft);
}
header.app-header {
  position: sticky; top: 0; z-index: 50;
  background: var(--sap-blue); color: white;
  padding: 0.75rem 1.5rem;
  display: flex; align-items: center; justify-content: space-between;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
header.app-header h1 { margin: 0; font-size: 1.1rem; font-weight: 600; }
header.app-header .meta { font-size: 0.85rem; opacity: 0.9; display: flex; gap: 1.5rem; align-items: center; }
header.app-header label { cursor: pointer; user-select: none; }
header.app-header input[type=checkbox] { margin-right: 0.3rem; vertical-align: middle; }

.layout { display: flex; min-height: calc(100vh - 50px); }

nav.sidebar {
  width: 240px; flex-shrink: 0;
  background: white; border-right: 1px solid var(--border);
  padding: 1rem 0;
  position: sticky; top: 50px; height: calc(100vh - 50px);
  overflow-y: auto;
}
nav.sidebar a {
  display: block;
  padding: 0.6rem 1.2rem;
  color: var(--text); text-decoration: none;
  border-left: 3px solid transparent;
  font-size: 0.92rem;
}
nav.sidebar a:hover { background: var(--bg-soft); }
nav.sidebar a.active { background: var(--subheader); border-left-color: var(--sap-blue); color: var(--sap-blue-dark); font-weight: 600; }
nav.sidebar .nav-section { font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); padding: 1rem 1.2rem 0.3rem; }

main {
  flex: 1; padding: 2rem 2.5rem;
  max-width: 1400px;
}

section.view { display: none; animation: fadeIn 0.2s ease-in; }
section.view.active { display: block; }
@keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }

.section-pill {
  display: inline-block;
  background: var(--sap-blue); color: white;
  padding: 0.4rem 1rem;
  border-radius: 999px;
  font-size: 0.85rem; font-weight: 600;
  margin-bottom: 0.5rem;
  letter-spacing: 0.02em;
}
section.view h2 {
  font-size: 1.6rem; margin: 0.5rem 0 0.5rem;
  color: var(--sap-blue-dark);
  border-bottom: 3px solid var(--sap-blue);
  padding-bottom: 0.5rem;
}
section.view .descr { color: var(--muted); margin-bottom: 1.5rem; }

/* Dashboard cards */
.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 1.25rem; margin-bottom: 2rem; }
.card {
  background: white; border-radius: 8px;
  padding: 1.25rem; box-shadow: 0 2px 6px rgba(0,0,0,0.06);
  border-top: 4px solid var(--sap-blue);
}
.card .kpi { font-size: 1.9rem; font-weight: 700; color: var(--sap-blue-dark); margin: 0.3rem 0; }
.card .kpi.neg { color: var(--neg); }
.card h3 { margin: 0; font-size: 0.95rem; color: var(--muted); font-weight: 500; }
.card p { margin: 0.6rem 0 0; font-size: 0.88rem; line-height: 1.4; }

/* Chart panels */
.chart-panel {
  background: white; border-radius: 8px;
  padding: 1.25rem; margin-bottom: 1.5rem;
  box-shadow: 0 2px 6px rgba(0,0,0,0.06);
}
.chart-panel h3 { margin: 0 0 1rem; font-size: 1rem; color: var(--sap-blue-dark); }
.chart-container { position: relative; height: 320px; }
.chart-container.tall { height: 450px; }

/* Table panel */
.table-panel {
  background: white; border-radius: 8px;
  padding: 1.25rem; margin-bottom: 1.5rem;
  box-shadow: 0 2px 6px rgba(0,0,0,0.06);
}
.table-panel h3 { margin: 0 0 1rem; font-size: 1rem; color: var(--sap-blue-dark); }
.filters { display: flex; flex-wrap: wrap; gap: 0.75rem; margin-bottom: 1rem; align-items: center; }
.filters label { font-size: 0.85rem; color: var(--muted); }
.filters select, .filters input {
  padding: 0.4rem 0.6rem; border: 1px solid var(--border); border-radius: 4px;
  font-size: 0.9rem; background: white;
}
.filters input { min-width: 200px; }

.tbl-wrap { overflow-x: auto; }
table.data {
  width: 100%; border-collapse: collapse; font-size: 0.88rem;
}
table.data th, table.data td {
  padding: 0.55rem 0.8rem;
  text-align: left;
  border-bottom: 1px solid var(--border);
}
table.data th {
  background: var(--subheader); color: var(--sap-blue-dark);
  font-weight: 600; cursor: pointer; user-select: none;
  position: sticky; top: 0;
  white-space: nowrap;
}
table.data th:hover { background: #C8D5EC; }
table.data th .arr { color: var(--sap-blue); margin-left: 0.3rem; }
table.data tr:nth-child(even) td { background: #FAFBFC; }
table.data tr:hover td { background: #EEF2F8; }
table.data td.num { text-align: right; font-variant-numeric: tabular-nums; }
table.data td.neg { color: var(--neg); font-weight: 600; }
table.data td.muted { color: var(--muted); font-size: 0.82rem; }

.pagination { display: flex; justify-content: center; gap: 0.4rem; margin-top: 1rem; align-items: center; font-size: 0.85rem; }
.pagination button {
  border: 1px solid var(--border); background: white;
  padding: 0.3rem 0.7rem; border-radius: 4px; cursor: pointer;
  font-size: 0.85rem;
}
.pagination button:hover:not(:disabled) { background: var(--bg-soft); }
.pagination button:disabled { opacity: 0.4; cursor: not-allowed; }
.pagination .info { color: var(--muted); margin: 0 0.5rem; }

/* Accordion (mapping) */
.accordion-item {
  background: white; border-radius: 6px; margin-bottom: 0.5rem;
  box-shadow: 0 1px 3px rgba(0,0,0,0.04);
  overflow: hidden;
}
.accordion-head {
  padding: 0.7rem 1rem; cursor: pointer;
  display: flex; justify-content: space-between; align-items: center;
  background: var(--subheader); font-weight: 600;
  border-left: 4px solid var(--sap-blue);
}
.accordion-head:hover { background: #C8D5EC; }
.accordion-head .chevron { transition: transform 0.2s; }
.accordion-item.open .chevron { transform: rotate(90deg); }
.accordion-body { display: none; padding: 0.5rem 1rem 1rem; }
.accordion-item.open .accordion-body { display: block; }

/* Key-value list (README) */
.kv-list { background: white; border-radius: 8px; padding: 1.5rem; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }
.kv-list .kv-row { display: grid; grid-template-columns: 200px 1fr; gap: 1rem; padding: 0.4rem 0; border-bottom: 1px solid var(--border); }
.kv-list .kv-row:last-child { border-bottom: none; }
.kv-list .kv-label { font-weight: 600; color: var(--sap-blue-dark); }
.kv-list .kv-row.heading { background: var(--subheader); padding: 0.6rem 0.8rem; border-radius: 4px; grid-template-columns: 1fr; font-weight: 700; color: var(--sap-blue-dark); margin: 0.8rem 0 0.4rem; border-bottom: none; }
.kv-list .kv-row.heading .kv-label { color: var(--sap-blue-dark); }

/* Warning banner */
.banner {
  background: #FFF4E5; border-left: 4px solid #F59E0B;
  padding: 0.8rem 1rem; border-radius: 4px;
  margin-bottom: 1.25rem; font-size: 0.9rem;
}
.banner strong { color: #92400E; }

footer.app-footer {
  background: white; padding: 1.5rem 2.5rem;
  border-top: 1px solid var(--border);
  font-size: 0.82rem; color: var(--muted);
  margin-left: 240px;
}
footer.app-footer p { margin: 0.3rem 0; }

@media (max-width: 900px) {
  nav.sidebar { display: none; }
  main { padding: 1.5rem 1rem; }
  footer.app-footer { margin-left: 0; }
  header.app-header { flex-direction: column; align-items: flex-start; gap: 0.4rem; }
  header.app-header .meta { font-size: 0.8rem; }
}
</style>
</head>
<body>

<header class="app-header">
  <h1>🏛️ Stadt Heidelberg — Haushaltsdaten 2017–2029</h1>
  <div class="meta">
    <span>Stand: <strong id="generated-date">__GEN__</strong></span>
    <label><input type="checkbox" id="th-fw-toggle"> TH FW einblenden</label>
  </div>
</header>

<div class="layout">
  <nav class="sidebar" id="sidebar">
    <a href="#dashboard" class="nav-link active" data-target="dashboard">🏠 Hauptsicht</a>
    <div class="nav-section">Sheets</div>
    <!-- nav links injected -->
  </nav>

  <main id="main">
    <!-- Dashboard -->
    <section class="view active" id="dashboard">
      <span class="section-pill">Hauptsicht</span>
      <h2>Drei Kernaussagen aus Stufe 1</h2>
      <p class="descr">Datenfundierte Beobachtungen aus den Haushaltsplänen 2019/20 bis Nachtrag 2026. Klick links auf einzelne Sheets für Details.</p>

      <div class="cards">
        <div class="card">
          <h3>1. Defizit ist strukturell</h3>
          <div class="kpi neg" id="kpi-defizit-2025">–</div>
          <p>Mittelfristplan 2027–2029 zeigt anhaltend –49 bis –52 Mio. €. Auch nach Nachtrag bleibt das Problem.</p>
        </div>
        <div class="card">
          <h3>2. Hebel konzentriert</h3>
          <div class="kpi">TH 51 · 20 · 50</div>
          <p>Drei Teilhaushalte verursachen 60 % des Defizits. TH 20 und 50 sind rechtlich gebunden — Hebel liegt außerhalb Top-3 oder gezielt in TH 51 (Jugendamt).</p>
        </div>
        <div class="card">
          <h3>3. Schulden steigen schneller als Defizit</h3>
          <div class="kpi neg" id="kpi-schulden-2029">–</div>
          <p>+413 Mio. € geplanter Schuldenaufbau bis 2029. Kombination aus Defizit- und Investitionsfinanzierung über Kredit ist der Treiber.</p>
        </div>
      </div>

      <div class="chart-panel">
        <h3>Ordentliches Ergebnis 2017–2029 (Mio. €)</h3>
        <div class="chart-container"><canvas id="chart-defizit"></canvas></div>
      </div>

      <div class="chart-panel">
        <h3>Schuldenstand Kernhaushalt 2018–2029 (Mio. €)</h3>
        <div class="chart-container"><canvas id="chart-schulden"></canvas></div>
      </div>

      <div class="chart-panel">
        <h3>Top-10 Teilhaushalte nach Defizit 2025 (Mio. €) — TH FW ausgeblendet</h3>
        <div class="chart-container tall"><canvas id="chart-th-top10"></canvas></div>
      </div>
    </section>

    <!-- Sheet sections (injected by JS) -->
    <section class="view" id="view-00_README"></section>
    <section class="view" id="view-01_Eckdaten"></section>
    <section class="view" id="view-02_Teilhaushalte"></section>
    <section class="view" id="view-08_Schulden_Liquiditaet"></section>
    <section class="view" id="view-09_Kennzahlen"></section>
    <section class="view" id="view-13_Mapping_TH_PB"></section>
    <section class="view" id="view-14_Stammdaten_TH"></section>
  </main>
</div>

<footer class="app-footer">
  <p><strong>Datenbasis:</strong> <code>Daten\01_HD_Haushaltsdaten.xlsx</code> (Stufe 1 MVP, 7 von 15 Sheets befüllt).</p>
  <p><strong>Quellen:</strong> Q1 = Haushaltsplan 2019/20 · Q2 = 2021/22 · Q3 = 2023/24 · Q4 = 2025/26 (Hauptquelle) · Q5 = Nachtragshaushalt 2026.</p>
  <p><strong>Konvention:</strong> Negative Werte rot. TH FW (Allgemeine Finanzwirtschaft) bei TH-Rankings standardmäßig ausgeblendet — zentrale Steuer-/Umlagenposition verzerrt sonst das Bild.</p>
  <p><strong>Regenerieren:</strong> <code>python WebApp/build_webapp.py</code></p>
</footer>

<script>
// ===========================================================================
// DATA
// ===========================================================================
const DATA = __DATA__;
const SHEET_ORDER = __SHEET_ORDER__;

document.getElementById("generated-date").textContent = DATA.generated;

// ===========================================================================
// Helpers
// ===========================================================================
function fmtNum(v, einheit) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v !== "number") return String(v);
  if (einheit === "EUR") {
    return v.toLocaleString("de-DE", {maximumFractionDigits: 0});
  }
  if (einheit === "%") {
    return v.toLocaleString("de-DE", {minimumFractionDigits: 1, maximumFractionDigits: 2}) + " %";
  }
  return v.toLocaleString("de-DE", {minimumFractionDigits: 1, maximumFractionDigits: 1});
}
function isNumeric(v) { return typeof v === "number" && !isNaN(v); }

function cellClass(v, header) {
  if (!isNumeric(v)) return "";
  let cls = "num";
  if (v < 0) cls += " neg";
  return cls;
}

function cellRender(v, header, einheit) {
  if (v === null || v === undefined) return "";
  if (isNumeric(v)) return fmtNum(v, einheit);
  return String(v);
}

// ===========================================================================
// Navigation
// ===========================================================================
function buildSidebar() {
  const sb = document.getElementById("sidebar");
  for (const name of SHEET_ORDER) {
    const meta = DATA.sheets[name];
    if (!meta) continue;
    const a = document.createElement("a");
    a.href = "#" + name;
    a.className = "nav-link";
    a.dataset.target = "view-" + name;
    a.textContent = meta.icon + " " + meta.label;
    sb.appendChild(a);
  }
}

function showView(id) {
  document.querySelectorAll("section.view").forEach(s => s.classList.remove("active"));
  document.querySelectorAll(".nav-link").forEach(a => a.classList.remove("active"));
  const v = document.getElementById(id);
  if (v) v.classList.add("active");
  const link = document.querySelector(`.nav-link[data-target="${id}"]`);
  if (link) link.classList.add("active");
  window.scrollTo(0, 0);
}

function setupRouting() {
  document.querySelectorAll(".nav-link").forEach(a => {
    a.addEventListener("click", e => {
      e.preventDefault();
      const tgt = a.dataset.target;
      location.hash = a.getAttribute("href").substring(1);
      showView(tgt);
    });
  });
  // Initial route
  const hash = location.hash.substring(1);
  if (hash === "dashboard" || !hash) {
    showView("dashboard");
  } else {
    showView("view-" + hash);
  }
}

// ===========================================================================
// Generic table renderer
// ===========================================================================
function makeTablePanel(sheetName, opts) {
  opts = opts || {};
  const sheet = DATA.sheets[sheetName];
  const headers = sheet.headers;
  const allRows = sheet.rows;

  // Build container
  const container = document.getElementById("view-" + sheetName);
  container.innerHTML = "";

  const pill = document.createElement("span");
  pill.className = "section-pill";
  pill.textContent = "Sheet " + sheetName.split("_")[0];
  container.appendChild(pill);

  const h2 = document.createElement("h2");
  h2.textContent = sheet.icon + " " + sheet.label;
  container.appendChild(h2);

  if (opts.descr) {
    const p = document.createElement("p");
    p.className = "descr";
    p.innerHTML = opts.descr;
    container.appendChild(p);
  }
  if (opts.banner) {
    const b = document.createElement("div");
    b.className = "banner";
    b.innerHTML = opts.banner;
    container.appendChild(b);
  }

  // Chart placeholder
  let chartPanel = null;
  if (opts.chart) {
    chartPanel = document.createElement("div");
    chartPanel.className = "chart-panel";
    chartPanel.innerHTML = `<h3>${opts.chart.title}</h3><div class="chart-container ${opts.chart.tall ? 'tall' : ''}"><canvas id="canvas-${sheetName}"></canvas></div>`;
    container.appendChild(chartPanel);
  }

  // Table panel
  const panel = document.createElement("div");
  panel.className = "table-panel";

  // Filters
  const filterDiv = document.createElement("div");
  filterDiv.className = "filters";

  const filterState = {};
  const filterDefs = opts.filters || [];

  // Free-text filter
  const fTxt = document.createElement("input");
  fTxt.type = "text";
  fTxt.placeholder = "Suche…";
  filterDiv.appendChild(Object.assign(document.createElement("label"), {textContent: "🔎 "}));
  filterDiv.appendChild(fTxt);

  // Column dropdowns
  for (const fd of filterDefs) {
    const lbl = document.createElement("label");
    lbl.textContent = fd.label + ": ";
    const sel = document.createElement("select");
    sel.dataset.col = fd.col;
    const optAll = document.createElement("option");
    optAll.value = "__ALL__";
    optAll.textContent = "Alle";
    sel.appendChild(optAll);
    const colIdx = headers.indexOf(fd.col);
    const vals = [...new Set(allRows.map(r => r[colIdx]).filter(v => v !== null && v !== undefined))].sort();
    for (const v of vals) {
      const o = document.createElement("option");
      o.value = String(v);
      o.textContent = String(v);
      sel.appendChild(o);
    }
    if (fd.default) sel.value = fd.default;
    filterDiv.appendChild(lbl);
    filterDiv.appendChild(sel);
    filterState[fd.col] = sel;
  }

  panel.appendChild(filterDiv);

  // Table
  const wrap = document.createElement("div");
  wrap.className = "tbl-wrap";
  const tbl = document.createElement("table");
  tbl.className = "data";
  const thead = document.createElement("thead");
  const trH = document.createElement("tr");
  headers.forEach((h, i) => {
    const th = document.createElement("th");
    th.textContent = h;
    th.dataset.col = i;
    trH.appendChild(th);
  });
  thead.appendChild(trH);
  tbl.appendChild(thead);
  const tbody = document.createElement("tbody");
  tbl.appendChild(tbody);
  wrap.appendChild(tbl);
  panel.appendChild(wrap);

  // Pagination
  const pagDiv = document.createElement("div");
  pagDiv.className = "pagination";
  panel.appendChild(pagDiv);

  container.appendChild(panel);

  // State
  const state = {
    sortCol: opts.defaultSortCol !== undefined ? opts.defaultSortCol : -1,
    sortDir: opts.defaultSortDir || "asc",
    page: 1,
    perPage: opts.perPage || 50,
  };

  // Find einheit-column index for formatting
  const einheitCol = headers.indexOf("Einheit");

  function getFiltered() {
    let rows = allRows.slice();
    // Apply TH FW toggle (only sheets with TH_Code column)
    const thCol = headers.indexOf("TH_Code");
    const hideFW = !document.getElementById("th-fw-toggle").checked;
    if (thCol >= 0 && hideFW) {
      rows = rows.filter(r => r[thCol] !== "FW");
    }
    // Apply column filters
    for (const [col, sel] of Object.entries(filterState)) {
      const v = sel.value;
      if (v !== "__ALL__") {
        const idx = headers.indexOf(col);
        rows = rows.filter(r => String(r[idx]) === v);
      }
    }
    // Apply text filter
    const txt = fTxt.value.toLowerCase().trim();
    if (txt) {
      rows = rows.filter(r => r.some(c => c !== null && c !== undefined && String(c).toLowerCase().includes(txt)));
    }
    return rows;
  }

  function sortRows(rows) {
    if (state.sortCol < 0) return rows;
    const dir = state.sortDir === "asc" ? 1 : -1;
    return rows.slice().sort((a, b) => {
      const av = a[state.sortCol], bv = b[state.sortCol];
      if (av === null || av === undefined) return 1;
      if (bv === null || bv === undefined) return -1;
      if (typeof av === "number" && typeof bv === "number") return (av - bv) * dir;
      return String(av).localeCompare(String(bv), "de") * dir;
    });
  }

  function render() {
    let rows = getFiltered();
    rows = sortRows(rows);
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / state.perPage));
    if (state.page > totalPages) state.page = totalPages;
    const start = (state.page - 1) * state.perPage;
    const pageRows = rows.slice(start, start + state.perPage);

    tbody.innerHTML = "";
    for (const r of pageRows) {
      const tr = document.createElement("tr");
      r.forEach((v, i) => {
        const td = document.createElement("td");
        const einheit = einheitCol >= 0 ? r[einheitCol] : null;
        td.className = cellClass(v, headers[i]);
        td.textContent = cellRender(v, headers[i], einheit);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }

    // Update sort arrows
    thead.querySelectorAll("th").forEach((th, i) => {
      const existing = th.querySelector(".arr");
      if (existing) existing.remove();
      if (i === state.sortCol) {
        const a = document.createElement("span");
        a.className = "arr";
        a.textContent = state.sortDir === "asc" ? "▲" : "▼";
        th.appendChild(a);
      }
    });

    // Pagination
    pagDiv.innerHTML = "";
    const prev = document.createElement("button");
    prev.textContent = "◀ Zurück";
    prev.disabled = state.page <= 1;
    prev.onclick = () => { state.page--; render(); };
    const next = document.createElement("button");
    next.textContent = "Weiter ▶";
    next.disabled = state.page >= totalPages;
    next.onclick = () => { state.page++; render(); };
    const info = document.createElement("span");
    info.className = "info";
    info.textContent = `Seite ${state.page} / ${totalPages}  ·  ${total} Zeilen${total !== allRows.length ? " (gefiltert)" : ""}`;
    pagDiv.appendChild(prev);
    pagDiv.appendChild(info);
    pagDiv.appendChild(next);

    // Update chart if applicable
    if (opts.chart && opts.chart.render) {
      opts.chart.render(rows, "canvas-" + sheetName);
    }
  }

  // Wire up sort
  thead.querySelectorAll("th").forEach((th, i) => {
    th.addEventListener("click", () => {
      if (state.sortCol === i) {
        state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
      } else {
        state.sortCol = i;
        state.sortDir = "asc";
      }
      render();
    });
  });

  // Wire up filters
  fTxt.addEventListener("input", () => { state.page = 1; render(); });
  for (const sel of Object.values(filterState)) {
    sel.addEventListener("change", () => { state.page = 1; render(); });
  }

  return { render, getFiltered, state, allRows, headers };
}

// ===========================================================================
// Sheet 00 README — Key-Value
// ===========================================================================
function renderReadme() {
  const container = document.getElementById("view-00_README");
  const sheet = DATA.sheets["00_README"];
  container.innerHTML = "";
  const pill = document.createElement("span");
  pill.className = "section-pill"; pill.textContent = "Sheet 00";
  container.appendChild(pill);
  const h2 = document.createElement("h2");
  h2.textContent = sheet.icon + " " + sheet.label;
  container.appendChild(h2);

  const list = document.createElement("div");
  list.className = "kv-list";
  for (const p of sheet.pairs) {
    const row = document.createElement("div");
    // Heuristik: Leere Value = Heading
    if (!p.value || p.value === "None") {
      row.className = "kv-row heading";
      const span = document.createElement("span");
      span.className = "kv-label";
      span.textContent = p.label;
      row.appendChild(span);
    } else {
      row.className = "kv-row";
      const lab = document.createElement("span");
      lab.className = "kv-label";
      lab.textContent = p.label;
      const val = document.createElement("span");
      val.textContent = p.value;
      row.appendChild(lab);
      row.appendChild(val);
    }
    list.appendChild(row);
  }
  container.appendChild(list);
}

// ===========================================================================
// Sheet 13 Mapping — Accordion
// ===========================================================================
function renderMapping() {
  const container = document.getElementById("view-13_Mapping_TH_PB");
  const sheet = DATA.sheets["13_Mapping_TH_PB"];
  container.innerHTML = "";

  const pill = document.createElement("span");
  pill.className = "section-pill"; pill.textContent = "Sheet 13";
  container.appendChild(pill);
  const h2 = document.createElement("h2");
  h2.textContent = sheet.icon + " " + sheet.label;
  container.appendChild(h2);
  const p = document.createElement("p");
  p.className = "descr";
  p.innerHTML = "Zuordnung Teilhaushalte ⇄ Produktbereiche/-gruppen (n:m). Klick auf einen Produktbereich, um seine Produktgruppen aufzuklappen.";
  container.appendChild(p);

  const headers = sheet.headers;
  const rows = sheet.rows;
  const pbIdx = headers.indexOf("PB_Code");
  const pbNameIdx = headers.indexOf("PB_Bezeichnung");
  const pgIdx = headers.indexOf("PG_Code");
  const pgNameIdx = headers.indexOf("PG_Bezeichnung");
  const thIdx = headers.indexOf("TH_Codes");

  // Group by PB
  const groups = {};
  for (const r of rows) {
    const pb = r[pbIdx];
    if (!groups[pb]) groups[pb] = { name: r[pbNameIdx], rows: [] };
    groups[pb].rows.push(r);
  }

  const wrap = document.createElement("div");
  for (const pb of Object.keys(groups).sort()) {
    const g = groups[pb];
    const item = document.createElement("div");
    item.className = "accordion-item";
    const head = document.createElement("div");
    head.className = "accordion-head";
    head.innerHTML = `<span><strong>PB ${pb}</strong> · ${g.name}  <span style="color:var(--muted);font-weight:normal">(${g.rows.length} Produktgruppen)</span></span><span class="chevron">▶</span>`;
    head.addEventListener("click", () => item.classList.toggle("open"));
    item.appendChild(head);

    const body = document.createElement("div");
    body.className = "accordion-body";
    const tbl = document.createElement("table");
    tbl.className = "data";
    tbl.innerHTML = `<thead><tr><th style="width:80px">PG</th><th>Bezeichnung</th><th style="width:240px">TH_Codes</th></tr></thead>`;
    const tbody = document.createElement("tbody");
    for (const r of g.rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${r[pgIdx] || ""}</td><td>${r[pgNameIdx] || ""}</td><td>${r[thIdx] || ""}</td>`;
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody);
    body.appendChild(tbl);
    item.appendChild(body);
    wrap.appendChild(item);
  }
  container.appendChild(wrap);
}

// ===========================================================================
// Dashboard charts
// ===========================================================================
const charts = {};

function dashCharts() {
  // 1. Defizit-Trend
  const eck = DATA.sheets["01_Eckdaten"];
  const eckH = eck.headers;
  const jahrI = eckH.indexOf("Jahr");
  const kennI = eckH.indexOf("Kennzahl");
  const wertI = eckH.indexOf("Wert");
  const defizitRows = eck.rows
    .filter(r => r[kennI] === "Ordentliches Ergebnis")
    .sort((a, b) => a[jahrI] - b[jahrI]);
  // Pro Jahr nur 1 Wert: priorisiere Plan, sonst nimm den ersten
  const byYear = {};
  for (const r of defizitRows) {
    const yr = r[jahrI];
    if (!byYear[yr]) byYear[yr] = r[wertI];
  }
  const dYears = Object.keys(byYear).sort();
  const dVals = dYears.map(y => byYear[y]);

  charts.defizit = new Chart(document.getElementById("chart-defizit"), {
    type: "line",
    data: {
      labels: dYears,
      datasets: [{
        label: "Ord. Ergebnis (Mio. €)",
        data: dVals,
        borderColor: "#2F5496",
        backgroundColor: "rgba(47,84,150,0.15)",
        fill: true,
        tension: 0.25,
        pointRadius: 5,
        pointBackgroundColor: dVals.map(v => v < 0 ? "#C00000" : "#2E7D32"),
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => `${ctx.parsed.y.toFixed(1)} Mio. €` } },
      },
      scales: {
        y: { ticks: { callback: v => v + " Mio." } },
      },
    },
  });

  // KPI 2025
  const v2025 = byYear["2025"];
  if (v2025 !== undefined) {
    document.getElementById("kpi-defizit-2025").textContent = v2025.toFixed(1) + " Mio. € (2025)";
  }

  // 2. Schulden-Trend
  const sch = DATA.sheets["08_Schulden_Liquiditaet"];
  const sH = sch.headers;
  const sJahr = sH.indexOf("Jahr"), sKenn = sH.indexOf("Kennzahl"), sWert = sH.indexOf("Wert");
  const schRows = sch.rows
    .filter(r => typeof r[sKenn] === "string" && r[sKenn].startsWith("Schuldenstand"))
    .sort((a, b) => a[sJahr] - b[sJahr]);
  const sByYear = {};
  for (const r of schRows) {
    const yr = r[sJahr];
    if (!sByYear[yr]) sByYear[yr] = r[sWert];
  }
  const sYears = Object.keys(sByYear).sort();
  const sVals = sYears.map(y => sByYear[y]);

  charts.schulden = new Chart(document.getElementById("chart-schulden"), {
    type: "line",
    data: {
      labels: sYears,
      datasets: [{
        label: "Schuldenstand (Mio. €)",
        data: sVals,
        borderColor: "#C00000",
        backgroundColor: "rgba(192,0,0,0.12)",
        fill: true, tension: 0.25,
        pointRadius: 4,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${ctx.parsed.y.toFixed(1)} Mio. €` } } },
      scales: { y: { ticks: { callback: v => v + " Mio." } } },
    },
  });

  const v2029 = sByYear["2029"];
  if (v2029 !== undefined) {
    document.getElementById("kpi-schulden-2029").textContent = v2029.toFixed(0) + " Mio. € (2029)";
  }

  // 3. TH-Top-10 Defizit 2025
  const th = DATA.sheets["02_Teilhaushalte"];
  const tH = th.headers;
  const tJahr = tH.indexOf("Jahr"), tCode = tH.indexOf("TH_Code"), tBez = tH.indexOf("TH_Bezeichnung"),
        tKenn = tH.indexOf("Kennzahl"), tWert = tH.indexOf("Wert");
  // Heuristik: Kennzahl-Name für ord. Ergebnis kann "Ord_Ergebnis" oder "Ordentliches Ergebnis" sein
  const ergRows = th.rows.filter(r =>
    r[tJahr] === 2025 &&
    r[tCode] !== "FW" &&
    typeof r[tKenn] === "string" &&
    (r[tKenn].toLowerCase().includes("ergebnis"))
  );
  ergRows.sort((a, b) => a[tWert] - b[tWert]);
  const top10 = ergRows.slice(0, 10);
  const lbls = top10.map(r => `${r[tCode]} ${r[tBez]}`);
  const vals = top10.map(r => r[tWert] / 1_000_000); // EUR → Mio. €

  charts.top10 = new Chart(document.getElementById("chart-th-top10"), {
    type: "bar",
    data: {
      labels: lbls,
      datasets: [{
        label: "Defizit 2025 (Mio. €)",
        data: vals,
        backgroundColor: "#C00000",
      }],
    },
    options: {
      indexAxis: "y",
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${ctx.parsed.x.toFixed(1)} Mio. €` } } },
      scales: { x: { ticks: { callback: v => v + " Mio." } } },
    },
  });
}

// ===========================================================================
// Per-sheet chart renderers
// ===========================================================================

function chartEckdatenLine(rows, canvasId) {
  // Wenn nur 1 Kennzahl gefiltert: zeige Verlauf
  const eck = DATA.sheets["01_Eckdaten"];
  const h = eck.headers;
  const jI = h.indexOf("Jahr"), kI = h.indexOf("Kennzahl"), wI = h.indexOf("Wert"), eI = h.indexOf("Einheit");
  const kennSet = [...new Set(rows.map(r => r[kI]))];

  if (charts["eck"]) { charts["eck"].destroy(); }
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  // Gruppiere pro Kennzahl, max 5 Linien
  const top5 = kennSet.slice(0, 5);
  const allYears = [...new Set(rows.map(r => r[jI]))].sort();
  const palette = ["#2F5496","#C00000","#2E7D32","#9333EA","#EA580C"];
  const datasets = top5.map((k, idx) => {
    const series = allYears.map(y => {
      const m = rows.find(r => r[jI] === y && r[kI] === k);
      return m ? m[wI] : null;
    });
    return {
      label: k,
      data: series,
      borderColor: palette[idx % palette.length],
      backgroundColor: palette[idx % palette.length] + "22",
      tension: 0.25,
      spanGaps: true,
    };
  });

  charts["eck"] = new Chart(ctx, {
    type: "line",
    data: { labels: allYears, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: "bottom" } },
    },
  });
}

function chartTeilhBar(rows, canvasId) {
  const th = DATA.sheets["02_Teilhaushalte"];
  const h = th.headers;
  const cI = h.indexOf("TH_Code"), bI = h.indexOf("TH_Bezeichnung"), wI = h.indexOf("Wert");

  if (charts["th"]) { charts["th"].destroy(); }
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  // Sortiert nach Wert ASC
  const sorted = rows.slice().sort((a, b) => a[wI] - b[wI]);
  const lbls = sorted.map(r => `${r[cI]} ${r[bI]}`);
  const vals = sorted.map(r => r[wI] / 1_000_000);
  const colors = vals.map(v => v < 0 ? "#C00000" : "#2F5496");

  charts["th"] = new Chart(ctx, {
    type: "bar",
    data: { labels: lbls, datasets: [{ label: "Wert (Mio. €)", data: vals, backgroundColor: colors }] },
    options: {
      indexAxis: "y",
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${ctx.parsed.x.toFixed(1)} Mio. €` } } },
      scales: { x: { ticks: { callback: v => v + " Mio." } } },
    },
  });
}

function chartSchuldenLine(rows, canvasId) {
  const h = DATA.sheets["08_Schulden_Liquiditaet"].headers;
  const jI = h.indexOf("Jahr"), kI = h.indexOf("Kennzahl"), wI = h.indexOf("Wert");
  const schRows = rows.filter(r => typeof r[kI] === "string" && r[kI].toLowerCase().includes("schuldenstand")).sort((a,b)=>a[jI]-b[jI]);

  if (charts["sch"]) { charts["sch"].destroy(); }
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  charts["sch"] = new Chart(ctx, {
    type: "line",
    data: { labels: schRows.map(r => r[jI]), datasets: [{ label: "Schuldenstand (Mio. €)", data: schRows.map(r => r[wI]), borderColor: "#C00000", backgroundColor: "rgba(192,0,0,0.12)", fill: true, tension: 0.25 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } },
  });
}

// ===========================================================================
// Bootstrap
// ===========================================================================

const renderers = {};

function buildAllSheets() {
  renderReadme();

  renderers["01_Eckdaten"] = makeTablePanel("01_Eckdaten", {
    descr: "Gesamthaushalt 2017–2029 im Long-Format. Filter setzen, dann zeigt der Chart oben die Zeitreihe.",
    filters: [
      { col: "Kennzahl", label: "Kennzahl" },
      { col: "Plan_Typ", label: "Plan-Typ" },
    ],
    chart: { title: "Zeitreihe (gefiltert)", render: chartEckdatenLine },
    defaultSortCol: 0,
    defaultSortDir: "asc",
    perPage: 50,
  });

  renderers["02_Teilhaushalte"] = makeTablePanel("02_Teilhaushalte", {
    descr: "41 Teilhaushalte × 7 Kennzahlen × 2 Jahre = 574 Datenpunkte (in EUR). Filter Jahr+Kennzahl setzen für ein TH-Ranking.",
    filters: [
      { col: "Jahr", label: "Jahr", default: "2025" },
      { col: "Kennzahl", label: "Kennzahl", default: "Ord_Ergebnis" },
    ],
    chart: { title: "TH-Ranking (gefiltert)", tall: true, render: chartTeilhBar },
    defaultSortCol: 4,
    defaultSortDir: "asc",
    perPage: 50,
  });

  renderers["08_Schulden_Liquiditaet"] = makeTablePanel("08_Schulden_Liquiditaet", {
    descr: "Schulden-Trajektorie 2018–2029 + Kennzahlen wie Aufwandsdeckungsgrad.",
    chart: { title: "Schuldenstand-Verlauf (Mio. €)", render: chartSchuldenLine },
    defaultSortCol: 0,
    defaultSortDir: "asc",
    perPage: 50,
  });

  renderers["09_Kennzahlen"] = makeTablePanel("09_Kennzahlen", {
    descr: "Kennzahlen zur finanziellen Leistungsfähigkeit.",
    banner: "<strong>⚠ Hinweis:</strong> Die Kennzahlentabellen in den PDFs sind um 90° gedreht — automatische Extraktion liefert verstümmelte Werte. Nur teilbefüllt; vollständige Aufnahme manuell erforderlich (Stufe 2).",
    defaultSortCol: 0,
    perPage: 50,
  });

  renderMapping();

  renderers["14_Stammdaten_TH"] = makeTablePanel("14_Stammdaten_TH", {
    descr: "Stammdaten aller Teilhaushalte mit Bezeichnungen über vier Plan-Generationen. Suche unten findet TH über alle Namensvarianten.",
    defaultSortCol: 0,
    perPage: 50,
  });

  // Initial render
  for (const r of Object.values(renderers)) r.render();
}

function setupGlobalToggle() {
  document.getElementById("th-fw-toggle").addEventListener("change", () => {
    for (const r of Object.values(renderers)) r.render();
  });
}

// ============ Go ============
buildSidebar();
setupRouting();
buildAllSheets();
dashCharts();
setupGlobalToggle();

// React to hash changes (back button)
window.addEventListener("hashchange", () => {
  const h = location.hash.substring(1);
  if (h === "" || h === "dashboard") showView("dashboard");
  else showView("view-" + h);
});
</script>

</body>
</html>
"""


# ---------------------------------------------------------------------------
# Render & write
# ---------------------------------------------------------------------------

def main():
    OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    print(f"Lese: {XLSX}")
    data = build_data()
    print(f"Sheets geladen: {list(data['sheets'].keys())}")
    for name, sh in data["sheets"].items():
        if sh["type"] == "kv":
            print(f"  {name}: {len(sh['pairs'])} Key-Value-Paare")
        else:
            print(f"  {name}: {len(sh['rows'])} Datenzeilen, {len(sh['headers'])} Spalten")

    html = HTML_TEMPLATE.replace(
        "__GEN__", data["generated"]
    ).replace(
        "__DATA__", json.dumps(data, ensure_ascii=False, default=str)
    ).replace(
        "__SHEET_ORDER__", json.dumps(POPULATED_SHEETS)
    )

    OUT_HTML.write_text(html, encoding="utf-8")
    size_kb = OUT_HTML.stat().st_size / 1024
    print(f"\n[OK] index.html geschrieben: {OUT_HTML}")
    print(f"     Größe: {size_kb:.1f} KB")


if __name__ == "__main__":
    main()
