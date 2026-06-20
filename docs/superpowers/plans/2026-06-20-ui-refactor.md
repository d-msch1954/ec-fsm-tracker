# UI Refactor — Salesforce-Style CRM Shell Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Decompose `EC_FieldService_Tracker.html` into `src/` files, write a new visual shell matching the UX mockup, and reassemble via `build.js` into a single self-contained deploy artifact.

**Architecture:** Extract SheetJS blob and existing JS/CSS from the monolithic HTML into `vendor/`, `src/core.js`, `src/ui.js`, `src/styles.css`, and `src/app.template.html`. Write new CSS and HTML structure; keep data logic; full-rewrite UI layer. `build.js` inlines everything back into `EC_FieldService_Tracker.html` + `docs/index.html`.

**Tech Stack:** Vanilla JS (ES5-compatible IIFE/UMD, no modules at runtime), SheetJS (inlined), hand-built SVG charts, CSS custom properties, Node.js ESM for build scripts only.

## Global Constraints
- Single self-contained HTML output — zero external requests at runtime (Google Fonts `<link>` is the only exception, degrades gracefully)
- No npm, no bundler, no framework
- `node build.js` is the only build command
- Primary color: `#86BC25` (Deloitte Green). Never use it for non-brand purposes
- Font stack: `'Open Sans', 'Segoe UI', system-ui, sans-serif`
- All JS wrapped in IIFE or UMD — no top-level `let`/`const` leaking to global scope (except `EC` and `XLSX`)
- `core.js` — pure functions, no DOM, no `XLSX` calls at module level
- Working dir: `C:\Custom Projects\ec-fsm-oppty-tracker`

---

## File Map

| File | Action | Purpose |
|---|---|---|
| `scripts/extract.js` | Create | One-time: splits existing HTML into src files |
| `vendor/xlsx.full.min.js` | Create (extracted) | SheetJS blob |
| `src/app.template.html` | Create | HTML skeleton with `<!--INLINE:*-->` markers |
| `src/styles.css` | Create | Complete visual system |
| `src/core.js` | Create | Pure data logic — parse, filter, KPIs, sort |
| `src/ui.js` | Create | DOM rendering + event wiring |
| `build.js` | Modify | Output to both root + docs/ |
| `EC_FieldService_Tracker.html` | Rebuilt output | — |
| `docs/index.html` | Rebuilt output | GitHub Pages artifact |

---

## Task 0: Extraction Script

**Files:**
- Create: `scripts/extract.js`
- Creates: `vendor/xlsx.full.min.js`, `src/core.ui.extracted.js`, `src/styles.css.bak`

- [ ] **Create `scripts/extract.js`**

```javascript
// scripts/extract.js  — run once: node scripts/extract.js
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

const root = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(root, '..', 'EC_FieldService_Tracker.html'), 'utf8');

const scriptRe = /<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi;
const scripts = [];
let m;
while ((m = scriptRe.exec(html)) !== null) scripts.push(m[1].trim());

const styleRe = /<style(?:\s[^>]*)?>([\s\S]*?)<\/style>/gi;
const styles = [];
while ((m = styleRe.exec(html)) !== null) styles.push(m[1].trim());

// SheetJS is the script block containing the XLSX global
const xlsxScript = scripts.find(s =>
  s.includes('var XLSX') || s.includes('XLSX.version') || s.includes('"SheetJS"')
);
const appScripts = scripts.filter(s => s !== xlsxScript);

mkdirSync(join(root, '..', 'vendor'), { recursive: true });
mkdirSync(join(root, '..', 'src'), { recursive: true });

if (!xlsxScript) { console.error('ERROR: SheetJS not found — check the HTML file'); process.exit(1); }

writeFileSync(join(root, '..', 'vendor', 'xlsx.full.min.js'), xlsxScript, 'utf8');
writeFileSync(join(root, '..', 'src', 'core.ui.extracted.js'),
  appScripts.join('\n\n// ─── next script block ───\n\n'), 'utf8');
writeFileSync(join(root, '..', 'src', 'styles.css.bak'),
  styles.join('\n\n/* ─── next style block ─── */\n\n'), 'utf8');

console.log(`Scripts: ${scripts.length} blocks | SheetJS: ${Math.round(xlsxScript.length/1024)}KB | App JS: ${Math.round(appScripts.reduce((a,s)=>a+s.length,0)/1024)}KB`);
console.log(`Styles: ${styles.length} blocks`);
console.log('Done. Review src/core.ui.extracted.js then split into core.js + ui.js');
```

- [ ] **Run extraction**

```
node scripts/extract.js
```

Expected output (approximate):
```
Scripts: 2 blocks | SheetJS: ~2000KB | App JS: ~50KB
Styles: 1 blocks
Done. Review src/core.ui.extracted.js then split into core.js + ui.js
```

- [ ] **Verify `vendor/xlsx.full.min.js` exists and is > 500KB**

```powershell
(Get-Item vendor/xlsx.full.min.js).Length / 1KB
```

Expected: > 500 (KB)

- [ ] **Commit**

```
git add scripts/extract.js vendor/xlsx.full.min.js src/core.ui.extracted.js src/styles.css.bak
git commit -m "feat: extract SheetJS and app JS/CSS from monolithic HTML"
```

---

## Task 1: Update `build.js`

**Files:**
- Modify: `build.js`

- [ ] **Replace `build.js` entirely**

```javascript
// build.js
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const root = dirname(fileURLToPath(import.meta.url));

const template = readFileSync(join(root, 'src', 'app.template.html'), 'utf8');
const styles   = readFileSync(join(root, 'src', 'styles.css'), 'utf8');
const xlsx     = readFileSync(join(root, 'vendor', 'xlsx.full.min.js'), 'utf8');
const core     = readFileSync(join(root, 'src', 'core.js'), 'utf8');
const ui       = readFileSync(join(root, 'src', 'ui.js'), 'utf8');

const out = template
  .replace('<!--INLINE:styles-->', `<style>\n${styles}\n</style>`)
  .replace('<!--INLINE:xlsx-->',   `<script>\n${xlsx}\n</script>`)
  .replace('<!--INLINE:core-->',   `<script>\n${core}\n</script>`)
  .replace('<!--INLINE:ui-->',     `<script>\n${ui}\n</script>`);

mkdirSync(join(root, 'docs'), { recursive: true });
writeFileSync(join(root, 'EC_FieldService_Tracker.html'), out, 'utf8');
writeFileSync(join(root, 'docs', 'index.html'), out, 'utf8');
console.log(`Built → ${Math.round(out.length / 1024)}KB`);
```

- [ ] **Commit**

```
git add build.js
git commit -m "feat: update build.js to inline 4 src files into dual outputs"
```

---

## Task 2: HTML Template

**Files:**
- Create: `src/app.template.html`

- [ ] **Create `src/app.template.html`**

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>E&C Field Service — Account Tracker</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
  <!--INLINE:styles-->
</head>
<body>
<div class="app" id="app">

  <aside class="sidebar">
    <div class="sidebar-logo">Deloitte<span class="logo-dot">.</span></div>
    <nav class="sidebar-nav">
      <a class="nav-item nav-item--disabled" data-view="home" href="#">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
        Home
      </a>
      <a class="nav-item nav-item--disabled" data-view="accounts" href="#">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>
        Accounts
      </a>
      <a class="nav-item nav-item--disabled" data-view="alliances" href="#">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v4l3 3"/></svg>
        Alliances
      </a>
    </nav>
    <div class="sidebar-import">
      <button class="btn-import" id="importBtn">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
        Import Excel
      </button>
      <div class="import-filename" id="importFilename"></div>
      <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none">
    </div>
  </aside>

  <main class="main" id="main">

    <div class="empty-state" id="emptyState">
      <div class="drop-zone" id="dropZone">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#9CA3AF" stroke-width="1.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
        <p class="drop-zone-title">Drop your Excel here</p>
        <p class="drop-zone-sub">or click <strong>Import Excel</strong> in the sidebar</p>
        <p class="drop-zone-hint">Expects a workbook with "Lean View" and optionally "Alliances" sheets</p>
      </div>
    </div>

    <div class="view" id="view-home" style="display:none">
      <h1 class="view-title">Home</h1>
      <div class="kpi-grid" id="kpiGrid"></div>
      <div class="charts-row" id="chartsRow"></div>
      <div class="home-section">
        <div class="home-section-header">
          <h2 class="section-title">High Priority Accounts</h2>
          <a href="#" class="view-all-link" id="viewAllLink">View all accounts →</a>
        </div>
        <div id="homeTable"></div>
      </div>
    </div>

    <div class="view" id="view-accounts" style="display:none">
      <div class="accounts-header">
        <h1 class="view-title" id="accountsTitle">Accounts</h1>
        <span class="record-count" id="recordCount"></span>
      </div>
      <div class="toolbar">
        <input class="search-input" id="searchInput" type="text" placeholder="Search accounts, POCs, status…">
        <div class="filter-chips" id="filterChips"></div>
        <button class="btn-clear-filters" id="clearFilters" style="display:none">✕ Clear filters</button>
      </div>
      <div id="accountsTable"></div>
    </div>

    <div class="view" id="view-alliances" style="display:none">
      <h1 class="view-title">Alliances</h1>
      <div class="alliance-cards" id="allianceCards"></div>
      <div class="alliance-detail" id="allianceDetail" style="display:none">
        <button class="btn-back" id="allianceBack">← Back to Alliances</button>
        <div id="allianceDetailContent"></div>
      </div>
    </div>

  </main>

  <aside class="detail-panel" id="detailPanel">
    <div class="panel-header">
      <button class="panel-close" id="panelClose" title="Close">✕</button>
    </div>
    <div class="panel-body" id="panelBody"></div>
  </aside>

</div>
<div class="toast" id="toast"></div>
<!--INLINE:xlsx-->
<!--INLINE:core-->
<!--INLINE:ui-->
</body>
</html>
```

- [ ] **Commit**

```
git add src/app.template.html
git commit -m "feat: add HTML template with sidebar layout and view scaffolding"
```

---

## Task 3: CSS — Layout Shell

**Files:**
- Create: `src/styles.css`

- [ ] **Create `src/styles.css` with the full design system**

```css
/* ── Reset & custom properties ─────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --green:      #86BC25;
  --green-tint: #F0F7E1;
  --amber:      #EF9F27;
  --amber-tint: #FEF3E0;
  --bg:         #F3F3F3;
  --surface:    #FFFFFF;
  --border:     #E5E7EB;
  --text:       #1A1A1A;
  --muted:      #6B7280;
  --row-hover:  #F9FAFB;
  --sidebar-w:  220px;
  --panel-w:    420px;
  --radius:     8px;
}

body {
  font-family: 'Open Sans', 'Segoe UI', system-ui, sans-serif;
  font-size: 14px;
  line-height: 1.5;
  background: var(--bg);
  color: var(--text);
  height: 100vh;
  overflow: hidden;
}

/* ── App shell ──────────────────────────────────────────────────────── */
.app { display: flex; height: 100vh; }

/* ── Sidebar ────────────────────────────────────────────────────────── */
.sidebar {
  width: var(--sidebar-w);
  min-width: var(--sidebar-w);
  background: var(--surface);
  border-right: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  height: 100vh;
  position: fixed;
  top: 0; left: 0;
  z-index: 100;
}

.sidebar-logo {
  padding: 24px 20px 20px;
  font-size: 20px;
  font-weight: 700;
  letter-spacing: -0.5px;
  color: var(--text);
  border-bottom: 1px solid var(--border);
}
.logo-dot { color: var(--green); }

.sidebar-nav {
  flex: 1;
  padding: 12px 8px;
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.nav-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 500;
  color: var(--text);
  text-decoration: none;
  cursor: pointer;
  transition: background 0.15s, color 0.15s;
  position: relative;
}
.nav-item:hover:not(.nav-item--disabled) { background: var(--bg); }
.nav-item--active {
  background: var(--green-tint);
  color: var(--green);
  font-weight: 600;
}
.nav-item--active::before {
  content: '';
  position: absolute;
  left: 0; top: 6px; bottom: 6px;
  width: 3px;
  background: var(--green);
  border-radius: 0 2px 2px 0;
}
.nav-item--disabled {
  color: var(--muted);
  cursor: default;
  opacity: 0.5;
}

.sidebar-import {
  padding: 16px;
  border-top: 1px solid var(--border);
}

.btn-import {
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 9px 16px;
  background: var(--green);
  color: #fff;
  border: none;
  border-radius: 6px;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  transition: opacity 0.15s;
}
.btn-import:hover { opacity: 0.88; }

.import-filename {
  margin-top: 8px;
  font-size: 11px;
  color: var(--muted);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

/* ── Main content ───────────────────────────────────────────────────── */
.main {
  margin-left: var(--sidebar-w);
  flex: 1;
  height: 100vh;
  overflow-y: auto;
  padding: 32px 36px;
  transition: margin-right 0.2s ease;
}
.app.panel-open .main { margin-right: var(--panel-w); }

.view-title {
  font-size: 22px;
  font-weight: 700;
  color: var(--text);
  margin-bottom: 24px;
}

/* ── Empty / import state ───────────────────────────────────────────── */
.empty-state {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
}
.drop-zone {
  text-align: center;
  padding: 64px 48px;
  border: 2px dashed var(--border);
  border-radius: var(--radius);
  background: var(--surface);
  max-width: 420px;
}
.drop-zone svg { margin-bottom: 16px; }
.drop-zone-title { font-size: 18px; font-weight: 600; margin-bottom: 8px; }
.drop-zone-sub   { color: var(--muted); margin-bottom: 6px; }
.drop-zone-hint  { font-size: 12px; color: var(--muted); }
.main.drag-over .drop-zone { border-color: var(--green); background: var(--green-tint); }

/* ── KPI cards ──────────────────────────────────────────────────────── */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 16px;
  margin-bottom: 24px;
}
.kpi-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 20px;
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.kpi-icon { color: var(--green); margin-bottom: 4px; }
.kpi-value { font-size: 28px; font-weight: 700; line-height: 1; }
.kpi-label { font-size: 13px; color: var(--muted); font-weight: 500; }
.kpi-sub   { font-size: 12px; color: var(--muted); }

/* ── Charts row ─────────────────────────────────────────────────────── */
.charts-row {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
  margin-bottom: 32px;
}
.chart-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 20px;
}
.chart-title { font-size: 13px; font-weight: 600; color: var(--muted); margin-bottom: 14px; text-transform: uppercase; letter-spacing: 0.04em; }
.chart-bar-row { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; font-size: 12px; cursor: pointer; }
.chart-bar-row:hover .chart-bar-fill { filter: brightness(0.92); }
.chart-bar-label { width: 110px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text); }
.chart-bar-track { flex: 1; height: 8px; background: var(--bg); border-radius: 4px; overflow: hidden; }
.chart-bar-fill  { height: 100%; border-radius: 4px; background: var(--green); transition: width 0.3s ease; }
.chart-bar-count { width: 24px; text-align: right; color: var(--muted); }

/* Priority stacked bar */
.stacked-bar { height: 16px; border-radius: 6px; overflow: hidden; display: flex; margin-top: 8px; }
.stacked-seg { height: 100%; transition: width 0.3s ease; }

/* ── Home section ───────────────────────────────────────────────────── */
.home-section { margin-top: 8px; }
.home-section-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
.section-title { font-size: 16px; font-weight: 600; }
.view-all-link { font-size: 13px; color: var(--green); text-decoration: none; font-weight: 500; }
.view-all-link:hover { text-decoration: underline; }

/* ── Accounts header / toolbar ──────────────────────────────────────── */
.accounts-header { display: flex; align-items: baseline; gap: 12px; margin-bottom: 16px; }
.record-count { font-size: 13px; color: var(--muted); }

.toolbar {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 16px;
  flex-wrap: wrap;
}
.search-input {
  padding: 8px 12px;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 13px;
  width: 260px;
  outline: none;
  background: var(--surface);
}
.search-input:focus { border-color: var(--green); box-shadow: 0 0 0 2px var(--green-tint); }

.filter-chips { display: flex; gap: 6px; flex-wrap: wrap; }
.chip {
  padding: 5px 12px;
  border: 1px solid var(--border);
  border-radius: 20px;
  font-size: 12px;
  cursor: pointer;
  background: var(--surface);
  color: var(--text);
  user-select: none;
  transition: border-color 0.15s, background 0.15s;
}
.chip:hover { border-color: var(--green); }
.chip--active { background: var(--green-tint); border-color: var(--green); color: var(--green); font-weight: 600; }

.btn-clear-filters {
  padding: 5px 12px;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 12px;
  cursor: pointer;
  background: transparent;
  color: var(--muted);
}
.btn-clear-filters:hover { border-color: var(--text); color: var(--text); }

/* ── Table ──────────────────────────────────────────────────────────── */
.data-table {
  width: 100%;
  border-collapse: collapse;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  overflow: hidden;
  font-size: 13px;
}
.data-table th {
  padding: 11px 14px;
  text-align: left;
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--muted);
  background: var(--bg);
  border-bottom: 1px solid var(--border);
  white-space: nowrap;
  cursor: pointer;
  user-select: none;
}
.data-table th:hover { color: var(--text); }
.data-table th.sort-asc::after  { content: ' ↑'; color: var(--green); }
.data-table th.sort-desc::after { content: ' ↓'; color: var(--green); }
.data-table td {
  padding: 12px 14px;
  border-bottom: 1px solid var(--border);
  vertical-align: middle;
}
.data-table tbody tr { cursor: pointer; transition: background 0.12s; }
.data-table tbody tr:hover { background: var(--row-hover); }
.data-table tbody tr:last-child td { border-bottom: none; }
.data-table tbody tr.row--selected { background: var(--green-tint); }
.data-table tbody tr.row--selected td:first-child {
  box-shadow: inset 3px 0 0 var(--green);
}

/* ── Pills & chips ──────────────────────────────────────────────────── */
.pill {
  display: inline-block;
  padding: 3px 10px;
  border-radius: 12px;
  font-size: 11px;
  font-weight: 600;
  white-space: nowrap;
}
.pill--HIGH    { background: var(--green-tint); color: #4a6e0a; }
.pill--MEDIUM  { background: var(--amber-tint);  color: #8a5600; }
.pill--LOW     { background: #F3F4F6; color: var(--muted); }
.pill--UNSET   { background: #F3F4F6; color: #9CA3AF; }

.tag {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 11px;
  font-weight: 500;
  margin-right: 4px;
}
/* Classification tag colors */
.tag--acp        { background: #EEF2FF; color: #3730A3; }
.tag--upstream   { background: #E0F2FE; color: #0369A1; }
.tag--mid        { background: #CCFBF1; color: #0F766E; }
.tag--ets        { background: #F3E8FF; color: #7E22CE; }
.tag--mining     { background: #FFF7ED; color: #C2410C; }
.tag--default    { background: #F3F4F6; color: var(--muted); }

.interacted-badge { font-size: 13px; }

/* ── Detail panel ───────────────────────────────────────────────────── */
.detail-panel {
  width: var(--panel-w);
  position: fixed;
  right: 0; top: 0;
  height: 100vh;
  background: var(--surface);
  border-left: 1px solid var(--border);
  box-shadow: -4px 0 16px rgba(0,0,0,0.10);
  transform: translateX(100%);
  transition: transform 0.2s ease;
  z-index: 200;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}
.detail-panel.open { transform: translateX(0); }

.panel-header {
  display: flex;
  justify-content: flex-end;
  padding: 14px 16px;
  border-bottom: 1px solid var(--border);
  flex-shrink: 0;
}
.panel-close {
  background: none; border: none; font-size: 18px; cursor: pointer;
  color: var(--muted); padding: 4px 8px; border-radius: 4px;
}
.panel-close:hover { background: var(--bg); color: var(--text); }

.panel-body {
  flex: 1;
  overflow-y: auto;
  padding: 20px;
}

.panel-company { font-size: 18px; font-weight: 700; margin-bottom: 6px; }
.panel-badges  { display: flex; gap: 6px; align-items: center; margin-bottom: 20px; flex-wrap: wrap; }
.panel-cols    { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }

.panel-section { margin-bottom: 20px; }
.panel-section-title {
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--muted);
  margin-bottom: 8px;
}
.panel-field { margin-bottom: 12px; }
.panel-field-label { font-size: 11px; color: var(--muted); margin-bottom: 2px; }
.panel-field-value { font-size: 13px; }
.panel-field-value a { color: var(--green); text-decoration: none; }
.panel-field-value a:hover { text-decoration: underline; }

.panel-bullets { list-style: none; padding: 0; }
.panel-bullets li { padding: 4px 0; font-size: 13px; padding-left: 14px; position: relative; }
.panel-bullets li::before { content: '•'; position: absolute; left: 0; color: var(--green); }

.panel-action-item { display: flex; gap: 10px; padding: 6px 0; border-bottom: 1px solid var(--border); font-size: 12px; }
.panel-action-date { color: var(--muted); white-space: nowrap; min-width: 40px; }

.panel-footer {
  padding: 14px 20px;
  border-top: 1px solid var(--border);
  font-size: 11px;
  color: var(--muted);
  flex-shrink: 0;
}

/* ── Alliance cards ─────────────────────────────────────────────────── */
.alliance-cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }
.alliance-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 24px;
  cursor: pointer;
  transition: box-shadow 0.15s, border-color 0.15s;
}
.alliance-card:hover { border-color: var(--green); box-shadow: 0 2px 12px rgba(134,188,37,0.12); }
.alliance-card-name { font-size: 16px; font-weight: 700; margin-bottom: 12px; }
.alliance-card-stat { font-size: 13px; color: var(--muted); margin-bottom: 4px; }
.alliance-card-action { font-size: 12px; color: var(--muted); margin-top: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

.btn-back {
  background: none; border: none; color: var(--green); font-size: 14px;
  font-weight: 600; cursor: pointer; padding: 0 0 20px; display: block;
}
.btn-back:hover { text-decoration: underline; }

/* ── Toast ──────────────────────────────────────────────────────────── */
.toast {
  position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
  background: var(--text); color: #fff; padding: 10px 20px;
  border-radius: 6px; font-size: 13px; z-index: 9999;
  opacity: 0; transition: opacity 0.2s; pointer-events: none;
}
.toast--visible { opacity: 1; }
.toast--error { background: #DC2626; }

/* ── Empty states ───────────────────────────────────────────────────── */
.empty-rows td {
  text-align: center; color: var(--muted); padding: 40px;
  font-size: 14px;
}

/* ── Print ──────────────────────────────────────────────────────────── */
@media print {
  .sidebar, .detail-panel, .toast, .toolbar, .btn-import { display: none !important; }
  .main { margin: 0; padding: 16px; }
  .app.panel-open .main { margin-right: 0; }
  .data-table { font-size: 11px; }
}
```

- [ ] **Commit**

```
git add src/styles.css
git commit -m "feat: add complete CSS design system — sidebar, KPI cards, table, pills, panel"
```

---

## Task 4: `core.js` — Pure Data Logic

**Files:**
- Create: `src/core.js`

**Interfaces — Produces (used by `ui.js`):**
- `EC.parseWorkbook(wb)` → `{ accounts: Account[], alliances: Alliance[] }`
- `EC.computeKPIs(accounts)` → `{ total, highPriority, highPct, engaged, engagedPct, meetingsIn30, nextMeeting }`
- `EC.groupBy(accounts, field)` → `{ [value]: count }`
- `EC.groupByAlliance(accounts)` → `{ [allianceName]: count }`
- `EC.filterAccounts(accounts, filters)` → `Account[]`
- `EC.sortAccounts(accounts, field, dir)` → `Account[]`
- `EC.parsePipelineBullets(text)` → `string[]`
- `EC.parseActionItems(text)` → `{ date: string|null, text: string }[]`
- `EC.extractEmails(text)` → `string[]`
- `EC.formatDate(date)` → `string`

**Account shape:**
```
{ company, classification, priority, status, interacted (bool),
  nextMeeting (Date|null), clusterPOC, keyPOC, clientPOC,
  existingVendors, existingWork, pipeline, action,
  lastUpdatedDate (Date|null), lastUpdatedBy, notes, alliance }
```

- [ ] **Create `src/core.js`**

```javascript
(function(root, factory) {
  if (typeof module !== 'undefined') module.exports = factory();
  else root.EC = factory();
})(typeof self !== 'undefined' ? self : this, function() {

  var STALENESS_DAYS = 30;
  var MEETING_WINDOW_DAYS = 30;

  var HEADER_MAP = {
    'company name':           'company',
    'account classification':  'classification',
    'priority focus':          'priority',
    'status / next steps':     'status',
    'interacted with?':        'interacted',
    'next meeting date':       'nextMeeting',
    'cluster poc':             'clusterPOC',
    'key account poc':         'keyPOC',
    'client poc':              'clientPOC',
    'existing vendors':        'existingVendors',
    'existing work done':      'existingWork',
    'vendor pipeline status':  'pipeline',
    'action items':            'action',
    'last updated date':       'lastUpdatedDate',
    'last updated by':         'lastUpdatedBy',
    'notes':                   'notes',
    'alliance':                'alliance',
  };

  var KNOWN_HEADERS = Object.keys(HEADER_MAP);

  // ── Parsing ───────────────────────────────────────────────────────────────

  function parseWorkbook(wb) {
    return {
      accounts:  parseAccountsSheet(wb),
      alliances: parseAlliancesSheet(wb),
    };
  }

  function parseAccountsSheet(wb) {
    var sheetName = wb.SheetNames.find(function(n) {
      return n.trim().toLowerCase() === 'lean view';
    }) || wb.SheetNames[0];
    var ws = wb.Sheets[sheetName];
    if (!ws) return [];

    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find header row by scanning first 6 rows for ≥2 known headers
    var headerIdx = -1, colMap = {};
    for (var i = 0; i < Math.min(6, rows.length); i++) {
      var cells = rows[i].map(function(c) { return String(c).toLowerCase().trim(); });
      var hits = KNOWN_HEADERS.filter(function(h) { return cells.indexOf(h) !== -1; }).length;
      if (hits >= 2) {
        headerIdx = i;
        cells.forEach(function(cell, j) { colMap[cell] = j; });
        break;
      }
    }
    if (headerIdx === -1) return [];

    function get(row, header) {
      var idx = colMap[header.toLowerCase()];
      return idx !== undefined ? String(row[idx] || '').trim() : '';
    }

    return rows.slice(headerIdx + 1)
      .filter(function(row) { return row.some(function(c) { return c !== ''; }); })
      .map(function(row) {
        var pri = get(row, 'priority focus').toUpperCase();
        return {
          company:        get(row, 'company name'),
          classification: get(row, 'account classification'),
          priority:       ['HIGH','MEDIUM','LOW'].indexOf(pri) !== -1 ? pri : 'UNSET',
          status:         get(row, 'status / next steps'),
          interacted:     get(row, 'interacted with?').toLowerCase() === 'yes',
          nextMeeting:    parseDate(get(row, 'next meeting date')),
          clusterPOC:     get(row, 'cluster poc'),
          keyPOC:         get(row, 'key account poc'),
          clientPOC:      get(row, 'client poc'),
          existingVendors:get(row, 'existing vendors'),
          existingWork:   get(row, 'existing work done'),
          pipeline:       get(row, 'vendor pipeline status'),
          action:         get(row, 'action items'),
          lastUpdatedDate:parseDate(get(row, 'last updated date')),
          lastUpdatedBy:  get(row, 'last updated by'),
          notes:          get(row, 'notes'),
          alliance:       get(row, 'alliance'),
        };
      })
      .filter(function(a) { return a.company; });
  }

  function parseAlliancesSheet(wb) {
    var sheetName = wb.SheetNames.find(function(n) {
      return n.trim().toLowerCase() === 'alliances';
    });
    if (!sheetName) return [];
    var ws = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return rows.map(function(r) {
      return {
        vendor:      String(r['Vendor'] || '').trim(),
        internalPOC: String(r['Internal POC'] || '').trim(),
        vendorPOC:   String(r['Vendor POC'] || '').trim(),
        comment:     String(r['Comment'] || '').trim(),
        nextAction:  String(r['Next Action'] || '').trim(),
      };
    }).filter(function(a) { return a.vendor; });
  }

  function parseDate(val) {
    if (!val || val === '') return null;
    var n = Number(val);
    if (!isNaN(n) && n > 1000) {        // Excel serial
      var d = new Date((n - 25569) * 86400000);
      if (!isNaN(d.getTime())) return d;
    }
    var d2 = new Date(val);
    return isNaN(d2.getTime()) ? null : d2;
  }

  // ── KPIs ──────────────────────────────────────────────────────────────────

  function computeKPIs(accounts) {
    var today = new Date(); today.setHours(0,0,0,0);
    var in30 = new Date(today); in30.setDate(today.getDate() + MEETING_WINDOW_DAYS);

    var total       = accounts.length;
    var highPriority= accounts.filter(function(a) { return a.priority === 'HIGH'; }).length;
    var engaged     = accounts.filter(function(a) { return a.interacted; }).length;
    var upcoming    = accounts.filter(function(a) {
      return a.nextMeeting && a.nextMeeting >= today && a.nextMeeting <= in30;
    });
    upcoming.sort(function(a,b) { return a.nextMeeting - b.nextMeeting; });

    return {
      total:        total,
      highPriority: highPriority,
      highPct:      total ? Math.round(highPriority / total * 100) : 0,
      engaged:      engaged,
      engagedPct:   total ? Math.round(engaged / total * 100) : 0,
      meetingsIn30: upcoming.length,
      nextMeeting:  upcoming.length ? upcoming[0].nextMeeting : null,
    };
  }

  // ── Grouping ──────────────────────────────────────────────────────────────

  function groupBy(accounts, field) {
    return accounts.reduce(function(acc, a) {
      var key = a[field] || 'Unset';
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});
  }

  function groupByAlliance(accounts) {
    var result = {};
    accounts.forEach(function(a) {
      if (!a.alliance) return;
      a.alliance.split(/[,;]/).forEach(function(s) {
        var al = s.trim();
        if (al) result[al] = (result[al] || 0) + 1;
      });
    });
    return result;
  }

  // ── Filtering ─────────────────────────────────────────────────────────────

  function filterAccounts(accounts, filters) {
    var search          = (filters.search || '').toLowerCase();
    var classifications = filters.classifications || [];
    var priorities      = filters.priorities || [];
    var alliances       = filters.alliances || [];
    var interacted      = filters.interacted !== undefined ? filters.interacted : null;
    var SEARCH_FIELDS   = ['company','classification','status','clusterPOC','keyPOC','clientPOC','pipeline','notes'];

    return accounts.filter(function(a) {
      if (search && !SEARCH_FIELDS.some(function(f) {
        return (a[f] || '').toLowerCase().indexOf(search) !== -1;
      })) return false;
      if (classifications.length && classifications.indexOf(a.classification) === -1) return false;
      if (priorities.length && priorities.indexOf(a.priority) === -1) return false;
      if (alliances.length) {
        var tags = (a.alliance || '').split(/[,;]/).map(function(s){ return s.trim(); }).filter(Boolean);
        if (!alliances.some(function(al) { return tags.indexOf(al) !== -1; })) return false;
      }
      if (interacted !== null && a.interacted !== interacted) return false;
      return true;
    });
  }

  function sortAccounts(accounts, field, dir) {
    var d = dir === 'desc' ? -1 : 1;
    return accounts.slice().sort(function(a, b) {
      var av = a[field], bv = b[field];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (av instanceof Date && bv instanceof Date) return d * (av - bv);
      return d * String(av).localeCompare(String(bv));
    });
  }

  // ── Text helpers ──────────────────────────────────────────────────────────

  function parsePipelineBullets(text) {
    if (!text) return [];
    var idx = text.toLowerCase().indexOf('pipeline:');
    var src = idx >= 0 ? text.slice(idx + 9) : text;
    return src.split(/[\n\r]+/)
      .map(function(l) { return l.replace(/^[-•*]\s*/, '').trim(); })
      .filter(Boolean);
  }

  function parseActionItems(text) {
    if (!text) return [];
    var re = /\[(\d{2}\/\d{2})\]\s*([^\[]+)/g, items = [], m;
    while ((m = re.exec(text)) !== null) {
      items.push({ date: m[1], text: m[2].trim() });
    }
    if (items.length) return items;
    return text.split(/[\n\r]+/)
      .map(function(l) { return l.trim(); })
      .filter(Boolean)
      .map(function(t) { return { date: null, text: t }; });
  }

  function extractEmails(text) {
    if (!text) return [];
    return (text.match(/[\w.+\-]+@[\w.\-]+\.\w+/g) || []);
  }

  function formatDate(d) {
    if (!d) return '—';
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  return {
    parseWorkbook: parseWorkbook,
    computeKPIs:   computeKPIs,
    groupBy:       groupBy,
    groupByAlliance: groupByAlliance,
    filterAccounts: filterAccounts,
    sortAccounts:  sortAccounts,
    parsePipelineBullets: parsePipelineBullets,
    parseActionItems: parseActionItems,
    extractEmails: extractEmails,
    formatDate:    formatDate,
    STALENESS_DAYS: STALENESS_DAYS,
    MEETING_WINDOW_DAYS: MEETING_WINDOW_DAYS,
  };
});
```

- [ ] **Commit**

```
git add src/core.js
git commit -m "feat: add core.js — pure data logic (parse, KPIs, filter, sort, text helpers)"
```

---

## Task 5: `ui.js` — Import, Routing, Toast, Panel

**Files:**
- Create: `src/ui.js` (skeleton — extended in Tasks 6–8)

- [ ] **Create `src/ui.js`** with state, init, file handling, nav routing, panel, toast:

```javascript
(function() {
  'use strict';

  // ── State ────────────────────────────────────────────────────────────────
  var state = {
    accounts: [],
    alliances: [],
    activeView: null,
    selectedAccount: null,
    filters: { search: '', classifications: [], priorities: [], alliances: [], interacted: null },
    sort: { field: 'company', dir: 'asc' },
  };

  // ── Helpers ──────────────────────────────────────────────────────────────
  function esc(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function classTag(classification) {
    var map = {
      'acp accounts': 'acp', 'upstream': 'upstream',
      'mid / downstream': 'mid', 'ets / ofs': 'ets', 'mining': 'mining'
    };
    var key = (classification || '').toLowerCase();
    return 'tag tag--' + (map[key] || 'default');
  }

  function priorityPill(p) {
    var label = p === 'UNSET' ? 'No Priority' : p;
    return '<span class="pill pill--' + esc(p) + '">' + esc(label) + '</span>';
  }

  function showToast(msg, type) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast toast--' + (type || 'success') + ' toast--visible';
    setTimeout(function() { t.className = 'toast'; }, 3500);
  }

  // ── Navigation ───────────────────────────────────────────────────────────
  function enableNav() {
    document.querySelectorAll('.nav-item').forEach(function(el) {
      el.classList.remove('nav-item--disabled');
    });
    document.getElementById('emptyState').style.display = 'none';
  }

  function navigate(view) {
    state.activeView = view;
    document.querySelectorAll('.view').forEach(function(el) { el.style.display = 'none'; });
    document.querySelectorAll('.nav-item').forEach(function(el) {
      el.classList.toggle('nav-item--active', el.dataset.view === view);
    });
    document.getElementById('view-' + view).style.display = '';
    closePanel();
    if (view === 'home')      renderHome();
    else if (view === 'accounts')  renderAccounts();
    else if (view === 'alliances') renderAlliances();
  }

  // ── File handling ────────────────────────────────────────────────────────
  function handleFile(file) {
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) {
      showToast('Please upload an .xlsx file', 'error'); return;
    }
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        var parsed = EC.parseWorkbook(wb);
        if (!parsed.accounts.length) {
          showToast('No accounts found — check the "Lean View" sheet exists', 'error'); return;
        }
        state.accounts = parsed.accounts;
        state.alliances = parsed.alliances;
        localStorage.setItem('ec_last_file', file.name);
        document.getElementById('importFilename').textContent = file.name;
        enableNav();
        showToast('Imported ' + parsed.accounts.length + ' accounts, ' + parsed.alliances.length + ' alliances');
        navigate('home');
      } catch(err) {
        showToast('Failed to read file: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ── Panel ────────────────────────────────────────────────────────────────
  function openPanel(account) {
    state.selectedAccount = account;
    renderPanel(account);
    document.getElementById('detailPanel').classList.add('open');
    document.getElementById('app').classList.add('panel-open');
  }

  function closePanel() {
    state.selectedAccount = null;
    document.getElementById('detailPanel').classList.remove('open');
    document.getElementById('app').classList.remove('panel-open');
    document.querySelectorAll('.row--selected').forEach(function(el) {
      el.classList.remove('row--selected');
    });
  }

  function renderPanel(account) {
    var emails = function(text) {
      return EC.extractEmails(text).map(function(e) {
        return '<a href="mailto:' + esc(e) + '">' + esc(e) + '</a>';
      }).join(', ') || esc(text) || '—';
    };
    var pocField = function(label, val) {
      if (!val) return '';
      return '<div class="panel-field"><div class="panel-field-label">' + esc(label) + '</div>' +
             '<div class="panel-field-value">' + emails(val) + '</div></div>';
    };
    var field = function(label, val) {
      if (!val) return '';
      return '<div class="panel-field"><div class="panel-field-label">' + esc(label) + '</div>' +
             '<div class="panel-field-value">' + esc(val) + '</div></div>';
    };

    var bullets = EC.parsePipelineBullets(account.pipeline);
    var actions = EC.parseActionItems(account.action);

    var allianceTags = (account.alliance || '').split(/[,;]/).map(function(s){ return s.trim(); }).filter(Boolean)
      .map(function(al) { return '<span class="tag tag--default">' + esc(al) + '</span>'; }).join(' ');

    var html = '<div class="panel-company">' + esc(account.company) + '</div>' +
      '<div class="panel-badges">' +
        '<span class="' + classTag(account.classification) + '">' + esc(account.classification || '—') + '</span>' +
        priorityPill(account.priority) +
        (account.interacted ? '<span class="pill" style="background:#E0F2FE;color:#0369A1">✓ Engaged</span>' : '') +
      '</div>' +
      '<div class="panel-cols">' +
        '<div>' +
          '<div class="panel-section-title">Details</div>' +
          field('Next Meeting', EC.formatDate(account.nextMeeting)) +
          pocField('Cluster POC', account.clusterPOC) +
          pocField('Key Account POC', account.keyPOC) +
          pocField('Client POC', account.clientPOC) +
          field('Existing Vendors', account.existingVendors) +
          field('Existing Work', account.existingWork) +
        '</div>' +
        '<div>' +
          (account.status ? '<div class="panel-section-title">Status / Next Steps</div>' +
            '<div class="panel-field-value" style="font-size:13px;margin-bottom:16px;white-space:pre-wrap">' + esc(account.status) + '</div>' : '') +
          (bullets.length ? '<div class="panel-section-title">Pipeline</div>' +
            '<ul class="panel-bullets">' +
              bullets.map(function(b){ return '<li>' + esc(b) + '</li>'; }).join('') +
            '</ul>' : '') +
          (actions.length ? '<div class="panel-section-title" style="margin-top:16px">Action Items</div>' +
            actions.map(function(a) {
              return '<div class="panel-action-item"><span class="panel-action-date">' +
                esc(a.date || '') + '</span><span>' + esc(a.text) + '</span></div>';
            }).join('') : '') +
        '</div>' +
      '</div>';

    if (allianceTags || account.lastUpdatedDate || account.lastUpdatedBy) {
      html += '<div class="panel-footer">' +
        (allianceTags ? '<div style="margin-bottom:6px">' + allianceTags + '</div>' : '') +
        'Last updated ' + EC.formatDate(account.lastUpdatedDate) +
        (account.lastUpdatedBy ? ' by ' + esc(account.lastUpdatedBy) : '') +
      '</div>';
    }

    document.getElementById('panelBody').innerHTML = html;
  }

  // ── Shared table builder ─────────────────────────────────────────────────
  function buildTable(accounts, compact) {
    if (!accounts.length) {
      return '<table class="data-table"><tbody><tr class="empty-rows"><td colspan="6">No accounts found</td></tr></tbody></table>';
    }
    var cols = compact
      ? ['company','classification','priority','nextMeeting','lastUpdatedDate']
      : ['company','classification','priority','interacted','nextMeeting','lastUpdatedDate','alliance'];

    var headers = {
      company: 'Company', classification: 'Classification', priority: 'Priority',
      interacted: 'Engaged', nextMeeting: 'Next Meeting',
      lastUpdatedDate: 'Last Updated', alliance: 'Alliance'
    };

    var thead = '<thead><tr>' + cols.map(function(c) {
      var cls = '';
      if (!compact) {
        if (state.sort.field === c) cls = ' class="sort-' + state.sort.dir + '"';
      }
      return '<th data-col="' + c + '"' + cls + '>' + headers[c] + '</th>';
    }).join('') + '</tr></thead>';

    var tbody = '<tbody>' + accounts.map(function(a, idx) {
      var row = cols.map(function(c) {
        if (c === 'priority')       return '<td>' + priorityPill(a.priority) + '</td>';
        if (c === 'classification') return '<td><span class="' + classTag(a.classification) + '">' + esc(a.classification || '—') + '</span></td>';
        if (c === 'interacted')     return '<td class="interacted-badge">' + (a.interacted ? '✓' : '—') + '</td>';
        if (c === 'nextMeeting' || c === 'lastUpdatedDate') return '<td>' + esc(EC.formatDate(a[c])) + '</td>';
        if (c === 'alliance') {
          var tags = (a.alliance || '').split(/[,;]/).map(function(s){ return s.trim(); }).filter(Boolean);
          return '<td>' + tags.map(function(t){ return '<span class="tag tag--default">' + esc(t) + '</span>'; }).join(' ') + '</td>';
        }
        return '<td>' + esc(a[c] || '—') + '</td>';
      }).join('');
      return '<tr data-idx="' + idx + '">' + row + '</tr>';
    }).join('') + '</tbody>';

    return '<table class="data-table">' + thead + tbody + '</table>';
  }

  function wireTableClicks(containerId, accounts) {
    var wrap = document.getElementById(containerId);
    if (!wrap) return;
    wrap.querySelectorAll('tbody tr').forEach(function(tr) {
      tr.addEventListener('click', function() {
        var idx = parseInt(tr.dataset.idx, 10);
        var account = accounts[idx];
        if (!account) return;
        wrap.querySelectorAll('.row--selected').forEach(function(r){ r.classList.remove('row--selected'); });
        tr.classList.add('row--selected');
        openPanel(account);
      });
    });
    // Sort on th click (accounts view only)
    if (!accounts._compact) {
      wrap.querySelectorAll('th[data-col]').forEach(function(th) {
        th.addEventListener('click', function() {
          var col = th.dataset.col;
          if (state.sort.field === col) {
            state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
          } else {
            state.sort.field = col; state.sort.dir = 'asc';
          }
          applyFilters();
        });
      });
    }
  }

  // ── Home view ────────────────────────────────────────────────────────────
  function renderHome() {
    var kpis = EC.computeKPIs(state.accounts);

    // KPI cards
    var kpiDefs = [
      { label: 'Total Accounts', value: kpis.total, sub: kpis.highPriority + ' high priority',
        icon: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>' },
      { label: 'High Priority', value: kpis.highPriority, sub: kpis.highPct + '% of total',
        icon: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>' },
      { label: 'Engaged', value: kpis.engagedPct + '%', sub: (100 - kpis.engagedPct) + '% not yet engaged',
        icon: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>' },
      { label: 'Meetings ≤30d', value: kpis.meetingsIn30,
        sub: kpis.nextMeeting ? 'next: ' + EC.formatDate(kpis.nextMeeting) : 'none scheduled',
        icon: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>' },
    ];
    document.getElementById('kpiGrid').innerHTML = kpiDefs.map(function(k) {
      return '<div class="kpi-card"><div class="kpi-icon">' + k.icon + '</div>' +
        '<div class="kpi-value">' + esc(String(k.value)) + '</div>' +
        '<div class="kpi-label">' + esc(k.label) + '</div>' +
        '<div class="kpi-sub">' + esc(k.sub) + '</div></div>';
    }).join('');

    // Mini charts
    renderCharts();

    // High priority table
    var high = EC.sortAccounts(
      state.accounts.filter(function(a) { return a.priority === 'HIGH'; }),
      'nextMeeting', 'asc'
    );
    var tableHtml = buildTable(high, true);
    document.getElementById('homeTable').innerHTML = tableHtml;
    var homeAccounts = high;
    homeAccounts._compact = true;
    wireTableClicks('homeTable', homeAccounts);
  }

  function renderCharts() {
    var byClass   = EC.groupBy(state.accounts, 'classification');
    var byPriority= EC.groupBy(state.accounts, 'priority');
    var byAlliance= EC.groupByAlliance(state.accounts);
    var total     = state.accounts.length || 1;

    function barRows(data, colorFn) {
      var entries = Object.entries(data).sort(function(a,b){ return b[1]-a[1]; });
      return entries.map(function(pair) {
        var label = pair[0], count = pair[1];
        var pct = Math.round(count / total * 100);
        return '<div class="chart-bar-row" data-filter-class="' + esc(label) + '">' +
          '<span class="chart-bar-label" title="' + esc(label) + '">' + esc(label) + '</span>' +
          '<div class="chart-bar-track"><div class="chart-bar-fill" style="width:' + pct + '%;background:' + (colorFn ? colorFn(label) : 'var(--green)') + '"></div></div>' +
          '<span class="chart-bar-count">' + count + '</span>' +
        '</div>';
      }).join('');
    }

    // Priority stacked bar
    var priColors = { HIGH: 'var(--green)', MEDIUM: 'var(--amber)', LOW: '#9CA3AF', UNSET: '#E5E7EB' };
    var priOrder  = ['HIGH','MEDIUM','LOW','UNSET'];
    var stackedSegs = priOrder.map(function(p) {
      var count = byPriority[p] || 0;
      var w = Math.round(count / total * 100);
      return w > 0 ? '<div class="stacked-seg" style="width:' + w + '%;background:' + priColors[p] + '" title="' + p + ': ' + count + '"></div>' : '';
    }).join('');
    var priLegend = priOrder.map(function(p) {
      var count = byPriority[p] || 0;
      return count > 0 ? '<span style="font-size:11px;color:var(--muted);margin-right:10px"><span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:' + priColors[p] + ';margin-right:3px"></span>' + p + ' ' + count + '</span>' : '';
    }).join('');

    var allianceEntries = Object.entries(byAlliance);
    var allianceHtml = allianceEntries.length
      ? barRows(byAlliance)
      : '<div style="color:var(--muted);font-size:12px">No alliance tags — add an Alliance column in Excel</div>';

    document.getElementById('chartsRow').innerHTML =
      '<div class="chart-card">' +
        '<div class="chart-title">By Classification</div>' +
        barRows(byClass) +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Priority Mix</div>' +
        '<div class="stacked-bar">' + stackedSegs + '</div>' +
        '<div style="margin-top:10px">' + priLegend + '</div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Alliance Split</div>' +
        allianceHtml +
      '</div>';

    // Click classification bar → filter in accounts view
    document.getElementById('chartsRow').querySelectorAll('.chart-bar-row[data-filter-class]').forEach(function(el) {
      el.addEventListener('click', function() {
        var cls = el.dataset.filterClass;
        state.filters.classifications = [cls];
        navigate('accounts');
      });
    });
  }

  // ── Accounts view ────────────────────────────────────────────────────────
  function renderAccounts() {
    renderFilterChips();
    applyFilters();
    wireSearch();
  }

  function renderFilterChips() {
    var chips = document.getElementById('filterChips');
    if (chips.dataset.wired) return;  // already built
    chips.dataset.wired = '1';

    var priorities = ['HIGH','MEDIUM','LOW','UNSET'];
    var classifications = Array.from(new Set(state.accounts.map(function(a){ return a.classification; }))).filter(Boolean).sort();
    var allianceNames = Object.keys(EC.groupByAlliance(state.accounts)).sort();

    function makeChips(values, filterKey) {
      return values.map(function(v) {
        var el = document.createElement('span');
        el.className = 'chip';
        el.textContent = v;
        el.dataset.value = v;
        el.dataset.filter = filterKey;
        el.addEventListener('click', function() {
          var arr = state.filters[filterKey];
          var idx = arr.indexOf(v);
          if (idx === -1) arr.push(v); else arr.splice(idx, 1);
          el.classList.toggle('chip--active', arr.indexOf(v) !== -1);
          applyFilters();
        });
        return el;
      });
    }

    makeChips(priorities, 'priorities').forEach(function(el){ chips.appendChild(el); });
    makeChips(classifications, 'classifications').forEach(function(el){ chips.appendChild(el); });
    if (allianceNames.length) makeChips(allianceNames, 'alliances').forEach(function(el){ chips.appendChild(el); });

    document.getElementById('clearFilters').addEventListener('click', function() {
      state.filters = { search: '', classifications: [], priorities: [], alliances: [], interacted: null };
      document.getElementById('searchInput').value = '';
      chips.querySelectorAll('.chip--active').forEach(function(el){ el.classList.remove('chip--active'); });
      applyFilters();
    });
  }

  function wireSearch() {
    var input = document.getElementById('searchInput');
    if (input.dataset.wired) return;
    input.dataset.wired = '1';
    var timer;
    input.addEventListener('input', function() {
      clearTimeout(timer);
      timer = setTimeout(function() {
        state.filters.search = input.value;
        applyFilters();
      }, 200);
    });
  }

  function applyFilters() {
    var filtered = EC.filterAccounts(state.accounts, state.filters);
    var sorted   = EC.sortAccounts(filtered, state.sort.field, state.sort.dir);

    var countEl = document.getElementById('recordCount');
    countEl.textContent = sorted.length === state.accounts.length
      ? sorted.length + ' accounts'
      : sorted.length + ' of ' + state.accounts.length + ' accounts';

    var hasFilters = state.filters.search || state.filters.classifications.length ||
      state.filters.priorities.length || state.filters.alliances.length || state.filters.interacted !== null;
    document.getElementById('clearFilters').style.display = hasFilters ? '' : 'none';

    document.getElementById('accountsTable').innerHTML = buildTable(sorted, false);
    wireTableClicks('accountsTable', sorted);

    // Re-apply sort indicators
    document.querySelectorAll('#accountsTable th[data-col]').forEach(function(th) {
      th.className = state.sort.field === th.dataset.col ? 'sort-' + state.sort.dir : '';
    });
  }

  // ── Alliances view ───────────────────────────────────────────────────────
  function renderAlliances() {
    document.getElementById('allianceCards').style.display = '';
    document.getElementById('allianceDetail').style.display = 'none';

    if (!state.alliances.length) {
      document.getElementById('allianceCards').innerHTML =
        '<div style="color:var(--muted);padding:32px">No Alliances sheet found in the workbook.</div>';
      return;
    }

    var allianceCounts = EC.groupByAlliance(state.accounts);

    var html = state.alliances.map(function(al, idx) {
      var count = allianceCounts[al.vendor] || 0;
      var highCount = state.accounts.filter(function(a) {
        return a.priority === 'HIGH' && (a.alliance || '').split(/[,;]/).map(function(s){ return s.trim(); }).indexOf(al.vendor) !== -1;
      }).length;
      return '<div class="alliance-card" data-idx="' + idx + '">' +
        '<div class="alliance-card-name">' + esc(al.vendor) + '</div>' +
        '<div class="alliance-card-stat">' + count + ' tagged accounts · ' + highCount + ' high priority</div>' +
        '<div class="alliance-card-stat">' + esc(EC.extractEmails(al.internalPOC)[0] || al.internalPOC || '—') + '</div>' +
        (al.nextAction ? '<div class="alliance-card-action">' + esc(al.nextAction) + '</div>' : '') +
      '</div>';
    }).join('');

    document.getElementById('allianceCards').innerHTML = html;

    document.getElementById('allianceCards').querySelectorAll('.alliance-card').forEach(function(card) {
      card.addEventListener('click', function() {
        showAllianceDetail(state.alliances[parseInt(card.dataset.idx, 10)]);
      });
    });
  }

  function showAllianceDetail(al) {
    document.getElementById('allianceCards').style.display = 'none';
    document.getElementById('allianceDetail').style.display = '';

    var tagged = state.accounts.filter(function(a) {
      return (a.alliance || '').split(/[,;]/).map(function(s){ return s.trim(); }).indexOf(al.vendor) !== -1;
    });
    tagged = EC.sortAccounts(tagged, 'priority', 'asc');

    var pocHtml = function(label, val) {
      if (!val) return '';
      var emails = EC.extractEmails(val).map(function(e) {
        return '<a href="mailto:' + esc(e) + '">' + esc(e) + '</a>';
      }).join(', ');
      return '<div class="panel-field"><div class="panel-field-label">' + esc(label) + '</div>' +
             '<div class="panel-field-value">' + (emails || esc(val)) + '</div></div>';
    };

    var rows = tagged.map(function(a, i) {
      return '<tr data-idx="' + i + '" style="cursor:pointer">' +
        '<td>' + esc(a.company) + '</td>' +
        '<td>' + priorityPill(a.priority) + '</td>' +
        '<td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + esc(a.status || '—') + '</td>' +
        '<td>' + esc(EC.formatDate(a.lastUpdatedDate)) + '</td>' +
      '</tr>';
    }).join('');

    document.getElementById('allianceDetailContent').innerHTML =
      '<h2 style="font-size:20px;font-weight:700;margin-bottom:16px">' + esc(al.vendor) + '</h2>' +
      pocHtml('Internal POC', al.internalPOC) +
      pocHtml('Vendor POC', al.vendorPOC) +
      (al.comment ? '<div class="panel-field"><div class="panel-field-label">Comment</div><div class="panel-field-value">' + esc(al.comment) + '</div></div>' : '') +
      (al.nextAction ? '<div class="panel-field"><div class="panel-field-label">Next Action</div><div class="panel-field-value">' + esc(al.nextAction) + '</div></div>' : '') +
      (tagged.length ? '<div style="margin-top:24px"><h3 style="font-size:14px;font-weight:600;margin-bottom:12px">Tagged Accounts (' + tagged.length + ')</h3>' +
        '<table class="data-table"><thead><tr><th>Company</th><th>Priority</th><th>Status</th><th>Last Updated</th></tr></thead>' +
        '<tbody>' + rows + '</tbody></table></div>'
        : '<div style="color:var(--muted);margin-top:16px">No accounts tagged with this alliance.</div>');

    // Wire row clicks to panel
    document.getElementById('allianceDetailContent').querySelectorAll('tbody tr').forEach(function(tr) {
      tr.addEventListener('click', function() {
        openPanel(tagged[parseInt(tr.dataset.idx, 10)]);
      });
    });
  }

  // ── Boot ─────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('importBtn').addEventListener('click', function() {
      document.getElementById('fileInput').click();
    });
    document.getElementById('fileInput').addEventListener('change', function(e) {
      handleFile(e.target.files[0]); e.target.value = '';
    });
    document.getElementById('panelClose').addEventListener('click', closePanel);
    document.getElementById('viewAllLink').addEventListener('click', function(e) {
      e.preventDefault();
      state.filters = { search: '', classifications: [], priorities: [], alliances: [], interacted: null };
      navigate('accounts');
    });
    document.getElementById('allianceBack').addEventListener('click', renderAlliances);
    document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closePanel(); });

    // Drag-and-drop
    var main = document.getElementById('main');
    main.addEventListener('dragover', function(e) { e.preventDefault(); main.classList.add('drag-over'); });
    main.addEventListener('dragleave', function() { main.classList.remove('drag-over'); });
    main.addEventListener('drop', function(e) { e.preventDefault(); main.classList.remove('drag-over'); handleFile(e.dataTransfer.files[0]); });

    // Sidebar nav
    document.querySelectorAll('.nav-item').forEach(function(el) {
      el.addEventListener('click', function(e) {
        e.preventDefault();
        if (!el.classList.contains('nav-item--disabled')) navigate(el.dataset.view);
      });
    });

    // Restore last filename
    var last = localStorage.getItem('ec_last_file');
    if (last) document.getElementById('importFilename').textContent = last;
  }

  document.addEventListener('DOMContentLoaded', init);
})();
```

- [ ] **Commit**

```
git add src/ui.js
git commit -m "feat: add ui.js — complete view rendering, sidebar nav, panel, filters"
```

---

## Task 6: Build & Smoke-Test

**Files:**
- None new — runs build, opens in browser

- [ ] **Run build**

```
node build.js
```

Expected:
```
Built → ~2500KB
```

- [ ] **Open in browser and verify:**
  1. Sidebar renders with Deloitte. logo and three disabled nav items
  2. Main area shows drop zone with dashed border
  3. Click "Import Excel" → file picker opens
  4. Import the real workbook → toast shows account count → navigates to Home
  5. Home shows 4 KPI cards, 3 mini-charts, high-priority table
  6. Click Accounts in sidebar → table with search/filter chips
  7. Type in search → table filters
  8. Click a row → slide-in panel opens from right
  9. Click × → panel closes
  10. Click Alliances → cards show; click a card → detail view

- [ ] **Fix any rendering issues** (CSS z-index, overflow, panel sizing) then rebuild

- [ ] **Run final build and push**

```
node build.js
git add EC_FieldService_Tracker.html docs/index.html src/core.js src/ui.js src/styles.css src/app.template.html build.js
git commit -m "feat: complete UI refactor — Salesforce-style CRM shell with sidebar, KPI cards, slide-in panel"
git push origin HEAD
```

- [ ] **Verify GitHub Pages deployment succeeded**

```
gh run list --repo d-msch1954/ec-fsm-tracker --limit 2
```

Expected: latest run shows `completed success`
