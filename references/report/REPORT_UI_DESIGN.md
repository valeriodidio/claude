# Design Grafico Report — Sistema UI Dark Theme

Tutti i report dell'admin Yeppon usano lo stesso design system dark.
Questo documento contiene il CSS completo, i componenti e le funzioni JS di formattazione.

---

## Palette colori — CSS custom properties

```css
:root {
    --bg:          #0f1624;   /* sfondo pagina principale */
    --bg-elev:     #182235;   /* card, filter bar, tabelle */
    --bg-elev2:    #1f2b41;   /* input, select, pulsanti secondari */
    --bg-row-a:    #141d2e;   /* righe zebra dispari */
    --bg-row-b:    #18223a;   /* righe zebra pari */
    --bg-expanded: #35527f;   /* riga espansa (drill aperto) */
    --bg-drill1:   #40608f;   /* drill level 1 */
    --bg-drill2:   #55769f;   /* drill level 2 */
    --bg-drill3:   #6e8cb5;   /* drill level 3 */
    --bg-hover:    #2e4370;   /* hover riga tabella */

    --border:      #2c3a55;   /* bordo principale */
    --border-soft: #243149;   /* bordo zebra righe */

    --text:        #e5eaf3;   /* testo principale */
    --text-mute:   #9aa7bd;   /* testo secondario / label */
    --text-dim:    #6f7d96;   /* testo disabilitato / zero */

    --accent:      #4f8cff;   /* blu primario (bottoni, tab attivo, bordi focus) */
    --accent-2:    #6aa1ff;   /* blu hover */

    --warn-bg:     #3a2f12;   /* sfondo banner warning */
    --warn-bd:     #6b5116;   /* bordo banner warning */
    --warn-fg:     #f5d27a;   /* testo banner warning */

    --neg:         #ff6b6b;   /* delta negativo / perdita */
    --pos:         #4ade80;   /* delta positivo / guadagno */
}
```

---

## Reset e base

```css
body {
    margin: 0;
    font-family: Arial, sans-serif;
    font-size: 13px;
    background: var(--bg);
    color: var(--text);
}
#Content { padding: 12px; }
h1 { margin: 0 0 12px 0; color: var(--text); }
h3 { color: var(--text); }
```

---

## Filter Bar

```css
.filter-bar {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    align-items: flex-end;
    background: var(--bg-elev);
    padding: 10px;
    border-radius: 6px;
    margin-bottom: 12px;
    border: 1px solid var(--border-soft);
}
.filter-bar label {
    display: block;
    font-size: 11px;
    color: var(--text-mute);
    margin-bottom: 2px;
}
.filter-bar .field { min-width: 140px; }
.filter-bar input[type=text],
.filter-bar select {
    padding: 5px 7px;
    font-size: 13px;
    background: var(--bg-elev2);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 4px;
}
.filter-bar input[type=text]:focus,
.filter-bar select:focus {
    outline: none;
    border-color: var(--accent);
}
.filter-bar button {
    padding: 6px 14px;
    background: var(--accent);
    color: #fff;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-weight: 600;
}
.filter-bar button:hover { background: var(--accent-2); }
.filter-bar button.secondary { background: #485874; }
.filter-bar button.secondary:hover { background: #5a6c8c; }
```

---

## KPI Cards

```css
.kpi-row {
    display: flex;
    gap: 10px;
    margin-bottom: 10px;
    flex-wrap: wrap;
}
.kpi {
    background: var(--bg-elev);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 10px 14px;
    min-width: 140px;
}
.kpi .lbl {
    font-size: 11px;
    color: var(--text-mute);
    text-transform: uppercase;
    letter-spacing: .5px;
}
.kpi .val {
    font-size: 18px;
    font-weight: bold;
    margin-top: 4px;
    color: var(--text);
}
```

HTML di esempio:
```html
<div class="kpi-row">
  <div class="kpi"><div class="lbl">Fatturato</div><div class="val" id="kpi-turnover">—</div></div>
  <div class="kpi"><div class="lbl">Ordini</div><div class="val" id="kpi-orders">—</div></div>
  <div class="kpi"><div class="lbl">Qtà</div><div class="val" id="kpi-qty">—</div></div>
</div>
```

---

## Tabs

```css
.tabs {
    display: flex;
    flex-wrap: wrap;
    border-bottom: 2px solid var(--accent);
    margin-bottom: 8px;
    gap: 2px;
}
.tabs a {
    padding: 7px 14px;
    cursor: pointer;
    background: var(--bg-elev);
    border: 1px solid var(--border);
    border-bottom: none;
    border-radius: 4px 4px 0 0;
    color: var(--text-mute);
    text-decoration: none;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    max-width: 340px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.tabs a:hover  { color: var(--text); background: var(--bg-elev2); }
.tabs a.active { background: var(--accent); color: #fff; border-color: var(--accent); }
.tab-panel        { display: none; }
.tab-panel.active { display: block; }
```

---

## Tabella report principale (table.report)

```css
table.report {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 12px;
    background: var(--bg-elev);
    border: 1px solid var(--border);
    border-radius: 6px;
    overflow: hidden;
}
table.report th, table.report td {
    border-bottom: 1px solid var(--border-soft);
    padding: 7px 10px;
    text-align: right;
}
table.report th {
    background: #1a2b4a;
    color: var(--text);
    text-align: center;
    cursor: pointer;
    border-bottom: 2px solid var(--accent);
    font-weight: 600;
    letter-spacing: .3px;
    position: sticky;
    top: 0;
    z-index: 1;
}
table.report th:hover { background: #223763; }
table.report td.text  { text-align: left; }

/* Zebra */
table.report tbody tr:nth-child(odd)  td { background: var(--bg-row-a); }
table.report tbody tr:nth-child(even) td { background: var(--bg-row-b); }
table.report tbody tr:hover td { background: var(--bg-hover); }

/* Riga totale */
table.report tr.total td {
    background: #1a2b4a !important;
    color: var(--text);
    font-weight: bold;
    border-top: 2px solid var(--accent);
}

/* Drill level 1 */
table.report tr.drill td {
    padding-left: 28px;
    background: var(--bg-drill1) !important;
    color: #f1f5ff;
    font-style: italic;
}
table.report tr.drill td:first-child { border-left: 4px solid var(--accent); }
table.report tr.drill:hover td { background: #4b6da0 !important; }

/* Drill level 2 */
table.report tr.drill2 td {
    padding-left: 48px;
    background: var(--bg-drill2) !important;
    color: #e8f0ff;
    font-style: italic;
}
table.report tr.drill2 td:first-child { border-left: 4px solid var(--accent-2); }

/* Drill level 3 */
table.report tr.drill3 td {
    padding-left: 68px;
    background: var(--bg-drill3) !important;
    color: #ddeaff;
}
table.report tr.drill3 td:first-child { border-left: 4px solid #8eb3d0; }

/* Riga espansa (padre con drill aperto) */
table.report tr.expanded td { background: var(--bg-expanded) !important; }
```

---

## Tabella prodotti (table.pd-table)

```css
table.pd-table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 12px;
    background: var(--bg-elev);
    border: 1px solid var(--border);
    border-radius: 6px;
    overflow: hidden;
}
table.pd-table th, table.pd-table td {
    padding: 6px 8px;
    border-bottom: 1px solid var(--border-soft);
    text-align: right;
    white-space: nowrap;
}
table.pd-table th {
    background: #1a2b4a;
    color: var(--text);
    text-align: center;
    border-bottom: 2px solid var(--accent);
    position: sticky;
    top: 0;
    z-index: 1;
    cursor: pointer;
}
table.pd-table th:hover { background: #223763; }
table.pd-table th.sorted-asc::after  { content: " \25B2"; font-size: 9px; }
table.pd-table th.sorted-desc::after { content: " \25BC"; font-size: 9px; }
table.pd-table td.text { text-align: left; white-space: normal; }
table.pd-table tbody tr:nth-child(odd)  td { background: var(--bg-row-a); }
table.pd-table tbody tr:nth-child(even) td { background: var(--bg-row-b); }
table.pd-table tbody tr:hover td { background: var(--bg-hover); }
.pd-table .mono {
    font-family: Consolas, 'Courier New', monospace;
    font-size: 11.5px;
    color: #c9d6ee;
}
```

---

## Barra Split + pulsante Expand All

```css
.split-bar {
    display: flex;
    gap: 10px;
    align-items: center;
    margin: 2px 0 10px 0;
}
.split-bar .split-metric {
    color: var(--text-mute);
    font-size: 12px;
    display: inline-flex;
    align-items: center;
    gap: 6px;
}
.split-bar .split-metric select {
    background: var(--bg-elev2);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 3px 6px;
    font-size: 12px;
}
button.split-toggle {
    background: var(--bg-elev2);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 5px 12px;
    cursor: pointer;
    font-size: 12px;
    display: inline-flex;
    align-items: center;
    gap: 6px;
}
button.split-toggle::before  { content: "\25A3"; font-size: 14px; color: var(--text-mute); }
button.split-toggle:hover    { border-color: var(--accent); color: #fff; }
button.split-toggle.active   { background: var(--accent); color: #fff; border-color: var(--accent); }
button.split-toggle.active::before { color: #fff; }
```

---

## Tabella Pivot

```css
table.report.pivot th, table.report.pivot td { padding: 5px 6px; font-size: 11.5px; }
table.report.pivot th.col-mktp { min-width: 80px; }
table.report.pivot th.col-total,
table.report.pivot td.col-total {
    background: #213557 !important;
    font-weight: 600;
    border-left: 2px solid var(--accent);
}
table.report.pivot tr.total td.col-total { background: #1a2b4a !important; }
table.report.pivot .cell-zero { color: var(--text-dim); }
```

---

## Icona lente (drill)

```css
.row-lens {
    display: inline-block;
    margin-left: 6px;
    cursor: pointer;
    font-size: 12px;
    opacity: .55;
    transition: opacity .15s, transform .15s;
    padding: 1px 4px;
    border-radius: 3px;
    user-select: none;
}
.row-lens:hover { opacity: 1; transform: scale(1.15); background: rgba(255,255,255,0.08); }
```

---

## Paginazione (pd-pager)

```css
.pd-pager {
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 10px 0;
    color: var(--text-mute);
    font-size: 12px;
}
.pd-pager button {
    background: var(--bg-elev2);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 4px 10px;
    cursor: pointer;
}
.pd-pager button:disabled { opacity: .4; cursor: not-allowed; }
.pd-pager button:not(:disabled):hover { border-color: var(--accent); color: #fff; }
.pd-pager input {
    background: var(--bg-elev2);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 3px 6px;
    width: 60px;
    text-align: center;
}
```

---

## Chart.js — tema dark

```javascript
Chart.defaults.color = '#9aa7bd';
Chart.defaults.borderColor = '#2c3a55';

// Opzioni comuni per tutti i grafici
const CHART_OPTIONS = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: { labels: { color: '#e5eaf3', boxWidth: 14, padding: 12 } },
        tooltip: {
            backgroundColor: '#182235',
            borderColor: '#2c3a55',
            borderWidth: 1,
            titleColor: '#e5eaf3',
            bodyColor: '#9aa7bd',
        }
    },
    scales: {
        x: {
            ticks: { color: '#9aa7bd' },
            grid:  { color: '#243149' }
        },
        y: {
            ticks: { color: '#9aa7bd' },
            grid:  { color: '#243149' }
        }
    }
};
```

---

## Librerie CDN (ordine di caricamento)

```html
<!-- jQuery -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>

<!-- Moment.js + DateRangePicker -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/daterangepicker@3.1.0/daterangepicker.min.js"></script>
<link  rel="stylesheet" href="https://cdn.jsdelivr.net/npm/daterangepicker@3.1.0/daterangepicker.css">

<!-- Select2 (multi-select filtri) -->
<link  rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.1.0-rc.0/css/select2.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.1.0-rc.0/js/select2.min.js"></script>

<!-- Chart.js -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
```

---

## Funzioni JS di formattazione numeri

```javascript
// Formatta € con separatore migliaia e 2 decimali
function fmtEur(v) {
    if (v == null || isNaN(v)) return '—';
    return '€ ' + Number(v).toLocaleString('it-IT', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

// Formatta numero intero con separatore migliaia
function fmt(v) {
    if (v == null || isNaN(v)) return '—';
    return Number(v).toLocaleString('it-IT');
}

// Formatta percentuale (es. 12.3%)
function fmtPct(v) {
    if (v == null || isNaN(v)) return '—';
    return Number(v).toFixed(1) + '%';
}

// Delta colorato (+/-)
function fmtDelta(v) {
    if (v == null || isNaN(v)) return '—';
    const s = (v >= 0 ? '+' : '') + Number(v).toLocaleString('it-IT', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
    const cls = v >= 0 ? 'pos' : 'neg';
    return `<span style="color:var(--${cls})">${s}%</span>`;
}
```

---

## Responsive (mobile)

```css
@media (max-width: 700px) {
    .filter-bar { flex-direction: column; }
    .filter-bar .field { width: 100%; }
    .kpi-row { flex-direction: column; }
    .kpi { min-width: unset; }
    table.report, table.pd-table {
        font-size: 11px;
        display: block;
        overflow-x: auto;
    }
    table.report th, table.report td,
    table.pd-table th, table.pd-table td { padding: 5px 6px; }
}
```

---

## Struttura HTML pagina tipo

```html
<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<title>Nome Report</title>
<!-- CDN -->
<style>/* :root + tutti i CSS sopra */</style>
</head>
<body>
<div id="Content">
  <h1>Nome Report</h1>

  <!-- Filter bar -->
  <div class="filter-bar">
    <div class="field">
      <label>Periodo</label>
      <input type="text" id="daterange">
    </div>
    <div class="field">
      <label>Marketplace</label>
      <select id="sel-mktp" multiple style="width:200px"></select>
    </div>
    <button id="btn-load">Carica</button>
    <button class="secondary" id="btn-export">⬇ Excel</button>
  </div>

  <!-- KPI -->
  <div class="kpi-row">
    <div class="kpi"><div class="lbl">Fatturato</div><div class="val" id="kpi-val">—</div></div>
  </div>

  <!-- Tabs -->
  <div class="tabs">
    <a class="active" data-tab="tab-brand">Brand</a>
    <a data-tab="tab-cat">Categoria</a>
  </div>
  <div id="tab-brand" class="tab-panel active">
    <table class="report"><thead>...</thead><tbody></tbody></table>
  </div>
  <div id="tab-cat" class="tab-panel">
    ...
  </div>
</div>

<script>/* JS */</script>
</body>
</html>
```
