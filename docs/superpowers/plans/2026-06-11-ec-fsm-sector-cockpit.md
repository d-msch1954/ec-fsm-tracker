# E&C Field Service Sector Cockpit — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rebuild `EC_FieldService_Tracker.html` into an intuitive, Deloitte-branded "sector cockpit" that imports the team's Excel tracker (view-only) and gives a field-service sector lead a cockpit overview, a powerful filterable accounts view, and an alliance view with drill-down — plus leadership PDF and Excel-of-current-view exports.

**Architecture:** Logic is developed in small ES-module-free UMD files (`src/core.js` = pure data logic, `src/ui.js` = DOM) with a Deloitte CSS file and an HTML template. SheetJS (extracted from the existing v1) is the only vendored library. `node --test` (Node built-in, zero install) unit-tests `src/core.js` against the real workbook. A small `build.js` inlines CSS + SheetJS + core + ui into the single self-contained `EC_FieldService_Tracker.html` that gets deployed to SharePoint. The UI is verified in a real browser via the preview MCP tools. The imported Excel remains the source of truth; the app never writes back except via explicit "export current view to Excel."

**Tech Stack:** Vanilla JS (no framework, no build toolchain on the deploy target), SheetJS (xlsx, inlined), hand-built SVG/CSS charts, Node built-in test runner for logic, Deloitte brand (Open Sans, green `#86BC25`).

**Testing convention:**
- `src/core.js` is pure (no DOM, no file I/O). It receives a SheetJS workbook object or plain rows and returns plain data. This is unit-tested with `node --test`.
- `src/ui.js` touches the DOM and is verified in the browser via preview tools (`preview_start`, `preview_snapshot`, `preview_click`, `preview_eval`). To make UI testable without a file picker, `ui.js` exposes `window.__ECTEST = { loadData }` so a fixture dataset can be injected.
- SheetJS is the I/O boundary. Both Node tests and the browser call `XLSX` to read the file, then hand the workbook to `core.parseWorkbook()`.

**Run from:** `C:\Custom Projects\ec-fsm-oppty-tracker` (the primary working directory; the existing v1 + `Code.gs` live here). Git is NOT initialized here, so "commit" steps are written as `git`-style messages but should be treated as save points; if the user initializes git later they map 1:1. Each "Commit" step = a logical save point / stopping point for review.

---

## File Structure

```
EC_FieldService_Tracker.html          BUILD OUTPUT — single self-contained file, deploy this
src/
  app.template.html                    HTML skeleton with <!--INLINE:*--> markers
  styles.css                           Deloitte-branded styles + print CSS
  core.js                              pure logic (UMD: Node + browser), no DOM
  ui.js                                DOM rendering, events, view routing
vendor/
  xlsx.full.min.js                     SheetJS, extracted from the existing v1 HTML
test/
  core.test.js                         node --test suite for core.js
  fixtures/
    sample.xlsx                        copy of the real "E&C Accounts Tracking Planner.xlsx"
build.js                               Node concat → EC_FieldService_Tracker.html
README.md                              how to use + deploy to SharePoint
EC_Accounts_Tracking_Planner_TEMPLATE.xlsx   updated template w/ Alliance column + dropdown
archive/Code.gs                        retired Google Apps Script backend (moved here)
```

Responsibilities:
- `core.js` — the only place data is interpreted. Header detection, column mapping, normalization, date parsing, KPI math, filtering, alliance grouping, text extraction (emails, pipeline bullets, action dates). Pure functions, fully tested.
- `ui.js` — renders the three views from `core.js` output, wires events, handles import + exports. No business rules beyond presentation.
- `styles.css` — all visual/brand styling, including `@media print`.
- `build.js` — deterministic inliner; the deploy artifact has zero external requests.

---

## Phase 0 — Foundation, vendoring, build, and import logic

### Task 1: Vendor SheetJS and set up the build skeleton

**Files:**
- Create: `vendor/xlsx.full.min.js`
- Create: `src/app.template.html`
- Create: `src/styles.css` (stub)
- Create: `src/core.js` (stub)
- Create: `src/ui.js` (stub)
- Create: `build.js`
- Create: `scripts/extract-sheetjs.js`

- [ ] **Step 1: Extract the inlined SheetJS from the existing v1 into a vendor file**

Create `scripts/extract-sheetjs.js`:

```js
const fs = require('fs');
const path = require('path');

const srcHtml = fs.readFileSync(path.join(__dirname, '..', 'EC_FieldService_Tracker.html'), 'utf8');
// The SheetJS block is the first <script> that contains the xlsx.js banner.
const start = srcHtml.indexOf('<script>/*! xlsx.js');
if (start === -1) throw new Error('SheetJS banner not found in v1 HTML');
const open = srcHtml.indexOf('>', start) + 1;
const end = srcHtml.indexOf('</script>', open);
if (end === -1) throw new Error('SheetJS closing tag not found');
const js = srcHtml.slice(open, end);
const out = path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js');
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, js, 'utf8');
console.log('Extracted SheetJS:', js.length, 'bytes ->', out);
```

- [ ] **Step 2: Run the extractor**

Run: `node scripts/extract-sheetjs.js`
Expected: prints "Extracted SheetJS: <large number > 500000> bytes -> vendor/xlsx.full.min.js"

- [ ] **Step 3: Verify SheetJS loads in Node**

Run: `node -e "const X=require('./vendor/xlsx.full.min.js'); console.log(typeof X.read, typeof X.utils.sheet_to_json)"`
Expected: `function function`

- [ ] **Step 4: Create the HTML template with inline markers**

Create `src/app.template.html`:

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>E&amp;C Field Service — Sector Cockpit</title>
<style>/*INLINE:styles.css*/</style>
</head>
<body>
<div id="app"></div>
<script>/*INLINE:vendor/xlsx.full.min.js*/</script>
<script>/*INLINE:src/core.js*/</script>
<script>/*INLINE:src/ui.js*/</script>
</body>
</html>
```

- [ ] **Step 5: Create stubs so the build has something to inline**

Create `src/styles.css`:

```css
:root { --green:#86BC25; }
body { font-family:'Open Sans','Segoe UI',system-ui,sans-serif; margin:0; }
```

Create `src/core.js`:

```js
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ECCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  return {};
});
```

Create `src/ui.js`:

```js
(function () {
  'use strict';
  document.getElementById('app').textContent = 'EC cockpit — build OK';
})();
```

- [ ] **Step 6: Write the build script**

Create `build.js`:

```js
const fs = require('fs');
const path = require('path');

const root = __dirname;
const tpl = fs.readFileSync(path.join(root, 'src', 'app.template.html'), 'utf8');

function inline(rel) {
  return fs.readFileSync(path.join(root, rel), 'utf8');
}

const out = tpl
  .replace('/*INLINE:styles.css*/', () => inline('src/styles.css'))
  .replace('/*INLINE:vendor/xlsx.full.min.js*/', () => inline('vendor/xlsx.full.min.js'))
  .replace('/*INLINE:src/core.js*/', () => inline('src/core.js'))
  .replace('/*INLINE:src/ui.js*/', () => inline('src/ui.js'));

const dest = path.join(root, 'EC_FieldService_Tracker.html');
fs.writeFileSync(dest, out, 'utf8');
console.log('Built', dest, '-', out.length, 'bytes');
```

Note: `String.prototype.replace` with a function replacer avoids `$&`/`$1` interpretation inside the large SheetJS blob — important, do not use a plain string replacement here.

- [ ] **Step 7: Build and verify the output exists and is self-contained**

Run: `node build.js`
Expected: prints "Built …EC_FieldService_Tracker.html - <bytes > 800000>"

Run: `node -e "const h=require('fs').readFileSync('EC_FieldService_Tracker.html','utf8'); console.log('hasSrcAttr:', /<script[^>]+src=/.test(h)); console.log('hasMarker:', h.includes('INLINE:'));"`
Expected: `hasSrcAttr: false` and `hasMarker: false` (everything inlined, no leftover markers, no external scripts)

- [ ] **Step 8: Commit**

```bash
git add vendor/xlsx.full.min.js src/ build.js scripts/extract-sheetjs.js
git commit -m "build: vendor SheetJS + HTML inlining build skeleton"
```

---

### Task 2: Copy the real workbook as a test fixture

**Files:**
- Create: `test/fixtures/sample.xlsx`

- [ ] **Step 1: Copy the real workbook (it lives in the sibling working dir and may be open/locked in Excel)**

Run (PowerShell, handles the file lock with FileShare.ReadWrite):

```powershell
$src = "C:\Custom Projects\ec-fsm-tracker-new\E&C Accounts Tracking Planner.xlsx"
$dst = "C:\Custom Projects\ec-fsm-oppty-tracker\test\fixtures\sample.xlsx"
New-Item -ItemType Directory -Force (Split-Path $dst) | Out-Null
$fs = [System.IO.File]::Open($src,'Open','Read','ReadWrite')
$o = [System.IO.File]::Create($dst); $fs.CopyTo($o); $o.Close(); $fs.Close()
"Copied: " + (Get-Item $dst).Length + " bytes"
```

Expected: prints a byte count > 500000.

- [ ] **Step 2: Verify SheetJS can read both sheets from the fixture**

Run: `node -e "const X=require('./vendor/xlsx.full.min.js'); const wb=X.readFile('./test/fixtures/sample.xlsx',{cellDates:true}); console.log(wb.SheetNames)"`
Expected: `[ 'Lean View', 'Alliances' ]`

- [ ] **Step 3: Commit**

```bash
git add test/fixtures/sample.xlsx
git commit -m "test: add real workbook as fixture"
```

---

### Task 3: Header detection + column mapping (core.js)

**Files:**
- Modify: `src/core.js`
- Test: `test/core.test.js`

- [ ] **Step 1: Write the failing test**

Create `test/core.test.js`:

```js
const test = require('node:test');
const assert = require('node:assert');
const XLSX = require('../vendor/xlsx.full.min.js');
const core = require('../src/core.js');

function rows(sheetName) {
  const wb = XLSX.readFile(__dirname + '/fixtures/sample.xlsx', { cellDates: true });
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
}

test('detectHeaderRow finds the Lean View header on row index 2', () => {
  const r = rows('Lean View');
  const idx = core.detectHeaderRow(r, ['Company Name', 'Account Classification']);
  assert.strictEqual(idx, 2); // 0-based; row 3 in Excel
});

test('buildColumnMap maps known headers to column indices by name', () => {
  const r = rows('Lean View');
  const idx = core.detectHeaderRow(r, ['Company Name', 'Account Classification']);
  const map = core.buildColumnMap(r[idx]);
  assert.ok(map.company >= 0, 'company mapped');
  assert.ok(map.classification >= 0, 'classification mapped');
  assert.ok(map.priority >= 0, 'priority mapped');
  assert.strictEqual(map.alliance, -1, 'Alliance column absent in current file → -1');
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test`
Expected: FAIL — `core.detectHeaderRow is not a function`.

- [ ] **Step 3: Implement detection + mapping**

In `src/core.js`, inside the factory (before `return`), add and export these. The header→field synonyms are deliberately tolerant (lowercased, punctuation-insensitive):

```js
var HEADER_SYNONYMS = {
  classification: ['account classification', 'classification'],
  company: ['company name', 'company', 'account'],
  priority: ['priority focus', 'priority'],
  status: ['status / next steps', 'status/next steps', 'status', 'next steps'],
  interacted: ['interacted with?', 'interacted with', 'interacted'],
  nextMeeting: ['next meeting date', 'next meeting', 'meeting date'],
  clusterPOC: ['cluster poc'],
  keyPOC: ['key account poc', 'key accountpoc', 'key poc'],
  clientPOC: ['client poc'],
  existingVendors: ['existing vendors'],
  existingWork: ['existing work done', 'existing work'],
  pipeline: ['vendor pipeline status', 'pipeline status', 'pipeline'],
  action: ['action items', 'action', 'next action'],
  lastUpdatedDate: ['last updated date', 'last updated'],
  lastUpdatedBy: ['last updated by', 'updated by'],
  notes: ['notes', 'note'],
  alliance: ['alliance', 'alliances', 'alliance tag']
};

function norm(s) {
  return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
}

function detectHeaderRow(rows, mustHave) {
  var wanted = (mustHave || []).map(norm);
  for (var i = 0; i < Math.min(rows.length, 8); i++) {
    var cells = (rows[i] || []).map(norm);
    var all = wanted.every(function (w) {
      return cells.some(function (c) { return c === w; });
    });
    if (all) return i;
  }
  return -1;
}

function buildColumnMap(headerRow) {
  var cells = (headerRow || []).map(norm);
  var map = {};
  Object.keys(HEADER_SYNONYMS).forEach(function (field) {
    map[field] = -1;
    var syns = HEADER_SYNONYMS[field];
    for (var c = 0; c < cells.length; c++) {
      if (syns.indexOf(cells[c]) !== -1) { map[field] = c; break; }
    }
  });
  return map;
}
```

Add to the returned object: `detectHeaderRow: detectHeaderRow, buildColumnMap: buildColumnMap, _norm: norm`.

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test`
Expected: PASS (3 tests in this file so far).

- [ ] **Step 5: Commit**

```bash
git add src/core.js test/core.test.js
git commit -m "feat(core): tolerant header detection and column mapping"
```

---

### Task 4: Value parsers — dates, emails, priority, pipeline, action dates (core.js)

**Files:**
- Modify: `src/core.js`
- Test: `test/core.test.js`

- [ ] **Step 1: Write the failing tests**

Append to `test/core.test.js`:

```js
test('parseDate handles Date, Excel serial, ISO/US strings, and blanks', () => {
  const d = core.parseDate(new Date(2026, 5, 1));
  assert.ok(d instanceof Date && d.getFullYear() === 2026);
  assert.strictEqual(core.parseDate('').valueOf?.(), undefined); // blank -> null
  assert.strictEqual(core.parseDate(''), null);
  assert.strictEqual(core.parseDate('not a date'), null);
  const us = core.parseDate('06/15/2026');
  assert.strictEqual(us.getMonth(), 5);
  const serial = core.parseDate(46000); // a 2025-ish Excel serial
  assert.ok(serial instanceof Date && serial.getFullYear() >= 2025);
});

test('extractEmails pulls all addresses from POC text', () => {
  const t = 'Tom Kirby - tokirby@deloitte.com\nThomas Hale - thhale@deloitte.com';
  assert.deepStrictEqual(core.extractEmails(t), ['tokirby@deloitte.com', 'thhale@deloitte.com']);
  assert.deepStrictEqual(core.extractEmails(''), []);
});

test('normalizePriority maps to HIGH/MEDIUM/LOW/null', () => {
  assert.strictEqual(core.normalizePriority('high'), 'HIGH');
  assert.strictEqual(core.normalizePriority(' MEDIUM '), 'MEDIUM');
  assert.strictEqual(core.normalizePriority(''), null);
  assert.strictEqual(core.normalizePriority('whatever'), null);
});

test('parsePipeline splits a lead line from bullet items', () => {
  const t = 'Target account for SF, greenfield account\nPipeline:\n- Mulesoft pipe\n- Informatica takeout';
  const p = core.parsePipeline(t);
  assert.match(p.lead, /greenfield/);
  assert.deepStrictEqual(p.bullets, ['Mulesoft pipe', 'Informatica takeout']);
});

test('latestActionDate extracts the most recent [MM/DD] tag', () => {
  assert.strictEqual(core.latestActionDate('[06/04] follow up; [10/29] earlier note'), '10/29');
  assert.strictEqual(core.latestActionDate('no dates here'), null);
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test`
Expected: FAIL — `core.parseDate is not a function`.

- [ ] **Step 3: Implement the parsers**

Add to `src/core.js` (inside factory):

```js
function parseDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number' && isFinite(v)) {
    // Excel serial date (1900 system): day 25569 = 1970-01-01
    var ms = Math.round((v - 25569) * 86400 * 1000);
    var d = new Date(ms);
    return isNaN(d) ? null : d;
  }
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
  if (m) {
    var mo = +m[1] - 1, day = +m[2];
    var yr = m[3] ? (+m[3] < 100 ? 2000 + +m[3] : +m[3]) : new Date().getFullYear();
    var d2 = new Date(yr, mo, day);
    return isNaN(d2) ? null : d2;
  }
  var t = Date.parse(s);
  return isNaN(t) ? null : new Date(t);
}

function extractEmails(text) {
  if (!text) return [];
  var re = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g;
  return String(text).match(re) || [];
}

function normalizePriority(v) {
  var s = norm(v).toUpperCase().trim();
  return (s === 'HIGH' || s === 'MEDIUM' || s === 'LOW') ? s : null;
}

function parsePipeline(text) {
  var out = { lead: '', bullets: [] };
  if (!text) return out;
  var lines = String(text).split(/\r?\n/);
  var leadLines = [];
  lines.forEach(function (ln) {
    var t = ln.trim();
    if (!t) return;
    if (/^pipeline\s*:?$/i.test(t)) return;
    if (/^potential\s*:?$/i.test(t)) return;
    if (/^[-•*]\s*/.test(t)) out.bullets.push(t.replace(/^[-•*]\s*/, '').trim());
    else leadLines.push(t);
  });
  out.lead = leadLines.join(' ');
  return out;
}

function latestActionDate(text) {
  if (!text) return null;
  var matches = String(text).match(/\[(\d{1,2}\/\d{1,2})\]/g);
  if (!matches) return null;
  // return the textually last tag (notes are appended chronologically)
  var last = matches[matches.length - 1];
  return last.replace(/[\[\]]/g, '');
}
```

Add to returned object: `parseDate, extractEmails, normalizePriority, parsePipeline, latestActionDate`.

- [ ] **Step 4: Run to verify pass**

Run: `node --test`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core.js test/core.test.js
git commit -m "feat(core): date/email/priority/pipeline/action parsers"
```

---

### Task 5: Account + alliance normalization and parseWorkbook (core.js)

**Files:**
- Modify: `src/core.js`
- Test: `test/core.test.js`

- [ ] **Step 1: Write the failing tests**

Append to `test/core.test.js`:

```js
test('parseWorkbook returns normalized accounts + alliances from the fixture', () => {
  const wb = XLSX.readFile(__dirname + '/fixtures/sample.xlsx', { cellDates: true });
  const data = core.parseWorkbook(wb);
  assert.strictEqual(data.accounts.length, 27);
  const cmp = data.accounts.find(a => a.company === 'Suncor');
  assert.strictEqual(cmp.priority, 'HIGH');
  assert.strictEqual(cmp.classification, 'Upstream');
  // alliances sheet
  const vendors = data.alliances.map(a => a.vendor).sort();
  assert.deepStrictEqual(vendors, ['Microsoft', 'Salesforce', 'Service Now']);
  // alliance tag absent in current file → empty array, not crash
  assert.deepStrictEqual(cmp.alliances, []);
  // meta describes what happened
  assert.strictEqual(data.meta.accountHeaderRow, 2);
  assert.strictEqual(data.meta.hasAllianceColumn, false);
});

test('parseWorkbook splits a comma/semicolon Alliance tag into an array', () => {
  // simulate a row with an Alliance column present
  const headers = ['Account Classification','Company Name','Priority Focus','Alliance'];
  const body = ['Upstream','TestCo','HIGH','Microsoft; Salesforce'];
  const ws = XLSX.utils.aoa_to_sheet([[], [], headers, body]); // header on row index 2
  const wb = { SheetNames:['Lean View'], Sheets:{ 'Lean View': ws } };
  const data = core.parseWorkbook(wb);
  assert.deepStrictEqual(data.accounts[0].alliances, ['Microsoft', 'Salesforce']);
  assert.strictEqual(data.meta.hasAllianceColumn, true);
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test`
Expected: FAIL — `core.parseWorkbook is not a function`.

- [ ] **Step 3: Implement normalization + parseWorkbook**

Add to `src/core.js`:

```js
var ALLIANCE_CANON = {
  'microsoft': 'Microsoft', 'msft': 'Microsoft',
  'service now': 'Service Now', 'servicenow': 'Service Now', 'snow': 'Service Now',
  'salesforce': 'Salesforce', 'sfdc': 'Salesforce'
};

function canonAlliance(s) {
  var k = norm(s);
  return ALLIANCE_CANON[k] || (s ? String(s).trim() : null);
}

function splitAlliances(v) {
  if (!v) return [];
  return String(v).split(/[;,/]/).map(function (x) { return canonAlliance(x); })
    .filter(Boolean);
}

function cell(row, i) { return (i >= 0 && row[i] != null) ? row[i] : ''; }

function findSheet(wb, name) {
  var target = norm(name);
  var hit = wb.SheetNames.find(function (n) { return norm(n) === target; });
  return hit ? wb.Sheets[hit] : null;
}

function normalizeAccounts(aoa) {
  var hr = detectHeaderRow(aoa, ['Company Name', 'Account Classification']);
  if (hr === -1) hr = detectHeaderRow(aoa, ['Company Name']);
  if (hr === -1) return { accounts: [], headerRow: -1, hasAlliance: false };
  var map = buildColumnMap(aoa[hr]);
  var accounts = [];
  for (var r = hr + 1; r < aoa.length; r++) {
    var row = aoa[r] || [];
    var company = String(cell(row, map.company)).trim();
    var classification = String(cell(row, map.classification)).trim();
    if (!company && !classification) continue; // skip blank/spacer rows
    accounts.push({
      id: accounts.length + 1,
      classification: classification,
      company: company,
      priority: normalizePriority(cell(row, map.priority)),
      priorityRaw: String(cell(row, map.priority)).trim(),
      status: String(cell(row, map.status)).trim(),
      interacted: /^y(es)?$/i.test(String(cell(row, map.interacted)).trim()),
      nextMeeting: parseDate(cell(row, map.nextMeeting)),
      nextMeetingRaw: String(cell(row, map.nextMeeting)).trim(),
      clusterPOC: String(cell(row, map.clusterPOC)).trim(),
      keyPOC: String(cell(row, map.keyPOC)).trim(),
      clientPOC: String(cell(row, map.clientPOC)).trim(),
      existingVendors: String(cell(row, map.existingVendors)).trim(),
      existingWork: String(cell(row, map.existingWork)).trim(),
      pipeline: String(cell(row, map.pipeline)).trim(),
      action: String(cell(row, map.action)).trim(),
      lastUpdatedDate: parseDate(cell(row, map.lastUpdatedDate)),
      lastUpdatedBy: String(cell(row, map.lastUpdatedBy)).trim(),
      notes: String(cell(row, map.notes)).trim(),
      alliances: splitAlliances(cell(row, map.alliance))
    });
  }
  return { accounts: accounts, headerRow: hr, hasAlliance: map.alliance !== -1 };
}

function normalizeAlliances(aoa) {
  if (!aoa || !aoa.length) return [];
  // headers on row 0; columns: Vendor, Internal POC, Vendor POC, Comment, Next Action
  var hr = 0;
  var map = {};
  var heads = (aoa[hr] || []).map(norm);
  map.vendor = heads.indexOf('vendor');
  map.internalPOC = heads.findIndex(function (h) { return h.indexOf('internal poc') !== -1; });
  map.vendorPOC = heads.findIndex(function (h) { return h.indexOf('vendor poc') !== -1; });
  map.comment = heads.indexOf('comment');
  map.nextAction = heads.findIndex(function (h) { return h.indexOf('next action') !== -1; });
  var out = [];
  for (var r = hr + 1; r < aoa.length; r++) {
    var row = aoa[r] || [];
    var vendor = String(cell(row, map.vendor)).trim();
    if (!vendor) continue;
    out.push({
      vendor: canonAlliance(vendor) || vendor,
      internalPOC: String(cell(row, map.internalPOC)).trim(),
      vendorPOC: String(cell(row, map.vendorPOC)).trim(),
      comment: String(cell(row, map.comment)).trim(),
      nextAction: String(cell(row, map.nextAction)).trim()
    });
  }
  return out;
}

function sheetToAoa(ws) {
  return ws ? require_xlsx().utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }) : [];
}
// XLSX is a global in the browser and required in Node; resolve either way:
function require_xlsx() {
  if (typeof XLSX !== 'undefined') return XLSX;
  if (typeof module === 'object') return require('../vendor/xlsx.full.min.js');
  throw new Error('XLSX not available');
}

function parseWorkbook(wb) {
  var leanWs = findSheet(wb, 'Lean View') || wb.Sheets[wb.SheetNames[0]];
  var allianceWs = findSheet(wb, 'Alliances');
  var leanAoa = sheetToAoa(leanWs);
  var allianceAoa = sheetToAoa(allianceWs);
  var acc = normalizeAccounts(leanAoa);
  var alliances = normalizeAlliances(allianceAoa);
  return {
    accounts: acc.accounts,
    alliances: alliances,
    meta: {
      accountHeaderRow: acc.headerRow,
      hasAllianceColumn: acc.hasAlliance,
      allianceSheetFound: !!allianceWs,
      sheetNames: wb.SheetNames.slice()
    }
  };
}
```

Add to returned object: `canonAlliance, splitAlliances, normalizeAccounts, normalizeAlliances, parseWorkbook`.

Note: `require_xlsx` resolves `XLSX` whether running in the browser (global from the inlined vendor script) or Node (require). In Node tests, `XLSX` is not a global, so it requires the vendor file — keep the relative path correct (`../vendor/...` from `src/`).

- [ ] **Step 4: Run to verify pass**

Run: `node --test`
Expected: PASS. (If the alliance vendor canonicalization makes `Service Now` vs `ServiceNow` mismatch the assertion, align the test and `ALLIANCE_CANON` so the sheet value "Service Now" canonicalizes consistently — the fixture's Alliances sheet stores it as "Service Now".)

- [ ] **Step 5: Commit**

```bash
git add src/core.js test/core.test.js
git commit -m "feat(core): account + alliance normalization and parseWorkbook"
```

---

## Phase 1 — Derived metrics + filtering (still pure logic)

### Task 6: KPI computation + per-account flags (core.js)

**Files:**
- Modify: `src/core.js`
- Test: `test/core.test.js`

- [ ] **Step 1: Write the failing tests**

Append to `test/core.test.js`:

```js
const NOW = new Date(2026, 5, 11); // 2026-06-11 to match the staleness window deterministically

test('computeKpis derives counts from existing fields only', () => {
  const wb = XLSX.readFile(__dirname + '/fixtures/sample.xlsx', { cellDates: true });
  const { accounts } = core.parseWorkbook(wb);
  const k = core.computeKpis(accounts, { now: NOW, staleDays: 30 });
  assert.strictEqual(k.total, 27);
  assert.strictEqual(k.high, 9);          // HIGH-priority accounts in the fixture
  assert.ok(k.openActions >= 1);          // several have [06/04] action items
  assert.ok(typeof k.engagedPct === 'number');
  assert.ok(k.whitespace >= 1);           // multiple untouched accounts
});

test('per-account flag helpers', () => {
  const a = { priority: null, status: '', action: '', lastUpdatedDate: null, nextMeeting: null };
  assert.strictEqual(core.isWhitespace(a), true);
  assert.strictEqual(core.hasOpenAction({ action: '[06/04] x' }), true);
  assert.strictEqual(core.hasOpenAction({ action: '' }), false);
  assert.strictEqual(core.isUpcomingMeeting({ nextMeeting: new Date(2026,5,20) }, NOW, 30), true);
  assert.strictEqual(core.isUpcomingMeeting({ nextMeeting: new Date(2026,8,1) }, NOW, 30), false);
  assert.strictEqual(core.isStale({ lastUpdatedDate: new Date(2026,3,1) }, NOW, 30), true);
  assert.strictEqual(core.isStale({ lastUpdatedDate: new Date(2026,5,10) }, NOW, 30), false);
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test`
Expected: FAIL — `core.computeKpis is not a function`.

- [ ] **Step 3: Implement KPI + flag helpers**

Add to `src/core.js`:

```js
function daysBetween(a, b) { return Math.floor((b - a) / 86400000); }

function isWhitespace(a) {
  return !a.priority && !a.status && !a.action && !a.nextMeeting && (!a.alliances || !a.alliances.length);
}
function hasOpenAction(a) { return !!(a.action && a.action.trim()); }
function isUpcomingMeeting(a, now, days) {
  if (!a.nextMeeting) return false;
  var d = daysBetween(now, a.nextMeeting);
  return d >= 0 && d <= days;
}
function isStale(a, now, days) {
  if (!a.lastUpdatedDate) return false; // unknown handled via whitespace/own panel, not "stale"
  return daysBetween(a.lastUpdatedDate, now) > days;
}

function computeKpis(accounts, opts) {
  opts = opts || {};
  var now = opts.now || new Date();
  var staleDays = opts.staleDays || 30;
  var total = accounts.length;
  var high = accounts.filter(function (a) { return a.priority === 'HIGH'; }).length;
  var engaged = accounts.filter(function (a) { return a.interacted; }).length;
  return {
    total: total,
    high: high,
    medium: accounts.filter(function (a) { return a.priority === 'MEDIUM'; }).length,
    low: accounts.filter(function (a) { return a.priority === 'LOW'; }).length,
    unset: accounts.filter(function (a) { return !a.priority; }).length,
    engaged: engaged,
    engagedPct: total ? Math.round((engaged / total) * 100) : 0,
    upcomingMeetings: accounts.filter(function (a) { return isUpcomingMeeting(a, now, staleDays); }).length,
    openActions: accounts.filter(hasOpenAction).length,
    stale: accounts.filter(function (a) { return isStale(a, now, staleDays); }).length,
    whitespace: accounts.filter(isWhitespace).length,
    byClassification: countBy(accounts, 'classification'),
    byAlliance: countAlliances(accounts)
  };
}

function countBy(accounts, key) {
  var m = {};
  accounts.forEach(function (a) { var v = a[key] || 'Unspecified'; m[v] = (m[v] || 0) + 1; });
  return m;
}
function countAlliances(accounts) {
  var m = { Microsoft: 0, 'Service Now': 0, Salesforce: 0 };
  accounts.forEach(function (a) {
    (a.alliances || []).forEach(function (al) { m[al] = (m[al] || 0) + 1; });
  });
  return m;
}
```

Add to returned object: `computeKpis, isWhitespace, hasOpenAction, isUpcomingMeeting, isStale, countBy, countAlliances, daysBetween`.

- [ ] **Step 4: Run to verify pass**

Run: `node --test`
Expected: PASS. If `k.high !== 9`, print the HIGH companies (`accounts.filter(a=>a.priority==='HIGH').map(a=>a.company)`) and reconcile against the fixture before changing the assertion — the fixture has 9 HIGH accounts (Suncor, Equinor, HF Sinclair, Enbridge, CITGO, Phillips 66, TechnipFMC, Nabors, Baker Hughes).

- [ ] **Step 5: Commit**

```bash
git add src/core.js test/core.test.js
git commit -m "feat(core): KPI computation and per-account flags"
```

---

### Task 7: Filtering, searching, sorting, alliance grouping (core.js)

**Files:**
- Modify: `src/core.js`
- Test: `test/core.test.js`

- [ ] **Step 1: Write the failing tests**

Append to `test/core.test.js`:

```js
test('filterAccounts applies search + multi-select + quick filters', () => {
  const wb = XLSX.readFile(__dirname + '/fixtures/sample.xlsx', { cellDates: true });
  const { accounts } = core.parseWorkbook(wb);

  const upstreamHigh = core.filterAccounts(accounts, {
    classifications: ['Upstream'], priorities: ['HIGH']
  });
  assert.ok(upstreamHigh.every(a => a.classification === 'Upstream' && a.priority === 'HIGH'));
  assert.ok(upstreamHigh.some(a => a.company === 'Suncor'));

  const search = core.filterAccounts(accounts, { search: 'agentforce' });
  assert.ok(search.some(a => a.company === 'HF Sinclair')); // appears in pipeline text

  const needsAction = core.filterAccounts(accounts, { quick: 'needsAction' });
  assert.ok(needsAction.every(a => core.hasOpenAction(a)));
});

test('sortAccounts sorts by a field with direction and stable priority order', () => {
  const a = [{company:'B',priority:'LOW'},{company:'A',priority:'HIGH'},{company:'C',priority:null}];
  const byCompany = core.sortAccounts(a, 'company', 'asc');
  assert.deepStrictEqual(byCompany.map(x=>x.company), ['A','B','C']);
  const byPriority = core.sortAccounts(a, 'priority', 'desc');
  assert.strictEqual(byPriority[0].priority, 'HIGH'); // HIGH > MEDIUM > LOW > unset
});

test('groupByAlliance buckets accounts by tag', () => {
  const accounts = [
    { company:'X', alliances:['Salesforce'] },
    { company:'Y', alliances:['Microsoft','Salesforce'] },
    { company:'Z', alliances:[] }
  ];
  const g = core.groupByAlliance(accounts);
  assert.deepStrictEqual(g['Salesforce'].map(a=>a.company), ['X','Y']);
  assert.deepStrictEqual(g['Microsoft'].map(a=>a.company), ['Y']);
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test`
Expected: FAIL — `core.filterAccounts is not a function`.

- [ ] **Step 3: Implement filter/sort/group**

Add to `src/core.js`:

```js
var PRIORITY_RANK = { HIGH: 3, MEDIUM: 2, LOW: 1 };

function matchesSearch(a, q) {
  if (!q) return true;
  var hay = [a.company, a.classification, a.status, a.notes, a.pipeline,
             a.clusterPOC, a.keyPOC, a.clientPOC, a.existingVendors, a.existingWork, a.action]
            .join(' ').toLowerCase();
  return hay.indexOf(q.toLowerCase()) !== -1;
}

function filterAccounts(accounts, c, ctx) {
  c = c || {}; ctx = ctx || {};
  var now = ctx.now || new Date();
  var staleDays = ctx.staleDays || 30;
  return accounts.filter(function (a) {
    if (!matchesSearch(a, c.search)) return false;
    if (c.classifications && c.classifications.length &&
        c.classifications.indexOf(a.classification) === -1) return false;
    if (c.priorities && c.priorities.length) {
      var p = a.priority || 'UNSET';
      if (c.priorities.indexOf(p) === -1) return false;
    }
    if (c.alliances && c.alliances.length) {
      var has = (a.alliances || []).some(function (al) { return c.alliances.indexOf(al) !== -1; });
      if (!has) return false;
    }
    if (c.interacted === true && !a.interacted) return false;
    if (c.interacted === false && a.interacted) return false;
    switch (c.quick) {
      case 'needsAction': if (!hasOpenAction(a)) return false; break;
      case 'meetingSoon': if (!isUpcomingMeeting(a, now, 7)) return false; break;
      case 'stale': if (!isStale(a, now, staleDays)) return false; break;
      case 'whitespace': if (!isWhitespace(a)) return false; break;
      default: break;
    }
    return true;
  });
}

function sortAccounts(accounts, field, dir) {
  var mul = dir === 'desc' ? -1 : 1;
  var copy = accounts.slice();
  copy.sort(function (x, y) {
    var xv, yv;
    if (field === 'priority') { xv = PRIORITY_RANK[x.priority] || 0; yv = PRIORITY_RANK[y.priority] || 0; }
    else if (field === 'nextMeeting' || field === 'lastUpdatedDate') {
      xv = x[field] ? x[field].getTime() : 0; yv = y[field] ? y[field].getTime() : 0;
    } else { xv = String(x[field] || '').toLowerCase(); yv = String(y[field] || '').toLowerCase(); }
    if (xv < yv) return -1 * mul;
    if (xv > yv) return 1 * mul;
    return 0;
  });
  return copy;
}

function groupByAlliance(accounts) {
  var g = { Microsoft: [], 'Service Now': [], Salesforce: [] };
  accounts.forEach(function (a) {
    (a.alliances || []).forEach(function (al) { (g[al] = g[al] || []).push(a); });
  });
  return g;
}
```

Add to returned object: `filterAccounts, sortAccounts, groupByAlliance, matchesSearch, PRIORITY_RANK`.

- [ ] **Step 4: Run to verify pass**

Run: `node --test`
Expected: PASS (all suites).

- [ ] **Step 5: Commit**

```bash
git add src/core.js test/core.test.js
git commit -m "feat(core): filtering, search, sorting, alliance grouping"
```

---

## Phase 2 — UI shell, styles, import (browser-verified)

### Task 8: Deloitte-branded styles + app shell + nav

**Files:**
- Modify: `src/styles.css`
- Modify: `src/ui.js`
- Modify: `src/app.template.html` (add the toolbar/import markup is done in JS; template just hosts `#app`)

- [ ] **Step 1: Write the styles**

Replace `src/styles.css` with the full Deloitte-light theme. Key tokens and rules (write these in full):

```css
:root{
  --green:#86BC25; --green-dark:#3B6D11; --green-tint:#EAF3DE;
  --ink:#282728; --muted:#5F5E5A; --line:#E3E2DC; --surface:#F6F6F2; --white:#fff;
  --amber:#EF9F27; --gray:#B4B2A9; --blue:#00A3E0;
  --radius:10px; --shadow:0 1px 2px rgba(0,0,0,.06);
}
*{box-sizing:border-box}
body{margin:0;background:var(--white);color:var(--ink);
  font-family:'Open Sans','Segoe UI',system-ui,sans-serif;font-size:14px;line-height:1.5}
.app-header{display:flex;align-items:center;gap:12px;padding:14px 22px;border-bottom:1px solid var(--line)}
.brand-dot{width:14px;height:14px;border-radius:50%;background:var(--green)}
.app-title{font-size:18px;font-weight:700}
.import-info{margin-left:auto;font-size:12px;color:var(--muted)}
.nav{display:flex;gap:22px;padding:0 22px;border-bottom:1px solid var(--line)}
.nav button{background:none;border:none;padding:12px 2px;font-size:14px;color:var(--muted);
  cursor:pointer;border-bottom:2px solid transparent}
.nav button.active{color:var(--ink);font-weight:600;border-bottom-color:var(--green)}
.view{padding:22px;max-width:1200px}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:22px}
.kpi{background:var(--surface);border-radius:var(--radius);padding:14px 16px}
.kpi.accent{border-left:3px solid var(--green)}
.kpi .label{font-size:12px;color:var(--muted)}
.kpi .value{font-size:26px;font-weight:700;margin-top:2px}
.card{background:var(--white);border:1px solid var(--line);border-radius:var(--radius);padding:16px 18px}
.cards-2{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;margin-bottom:22px}
.bar-row{display:flex;align-items:center;gap:10px;margin:8px 0;font-size:13px}
.bar-row .lab{width:120px;color:var(--muted)}
.bar-track{flex:1;height:14px;background:var(--surface);border-radius:4px;overflow:hidden}
.bar-fill{height:100%;background:var(--green);border-radius:4px}
.pill{font-size:11px;padding:2px 8px;border-radius:5px;font-weight:600;display:inline-block}
.pill.high{background:var(--green-tint);color:var(--green-dark)}
.pill.medium{background:#FAEEDA;color:#854F0B}
.pill.low{background:#F1EFE8;color:#5F5E5A}
.pill.unset{background:#F1EFE8;color:#888780}
table.accounts{width:100%;border-collapse:collapse;font-size:13px}
table.accounts th{text-align:left;padding:10px 12px;border-bottom:1px solid var(--line);
  color:var(--muted);font-weight:600;cursor:pointer;white-space:nowrap}
table.accounts td{padding:10px 12px;border-bottom:1px solid var(--line);vertical-align:top}
table.accounts tr:hover{background:var(--surface);cursor:pointer}
.toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:16px}
.toolbar input[type=search]{flex:1;min-width:220px;padding:9px 12px;border:1px solid var(--line);border-radius:8px;font-size:14px}
.chip{font-size:12px;padding:6px 12px;border:1px solid var(--line);border-radius:20px;background:var(--white);cursor:pointer}
.chip.on{background:var(--green);color:var(--white);border-color:var(--green)}
.btn{padding:9px 14px;border:1px solid var(--line);border-radius:8px;background:var(--white);cursor:pointer;font-size:13px}
.btn.primary{background:var(--green);color:#fff;border-color:var(--green)}
.alliance-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px}
.drawer-overlay{position:fixed;inset:0;background:rgba(0,0,0,.4);display:flex;justify-content:flex-end;z-index:50}
.drawer{width:min(560px,92vw);background:#fff;height:100%;overflow:auto;padding:24px}
.empty{padding:60px 20px;text-align:center;color:var(--muted)}
.dropzone{border:2px dashed var(--line);border-radius:var(--radius);padding:50px;text-align:center;color:var(--muted)}
.dropzone.drag{border-color:var(--green);background:var(--green-tint)}
@media print{
  .nav,.toolbar,.app-header .import-info,.btn,.no-print{display:none!important}
  .view{padding:0}
  body{font-size:11px}
  .card,.kpi{break-inside:avoid}
}
```

- [ ] **Step 2: Build the app shell + nav in ui.js**

Replace `src/ui.js` with a structure that: holds app state, renders the header + nav + the active view, and exposes a test hook. Write the shell now (views added in later tasks call into `renderCockpit`/`renderAccounts`/`renderAlliances`, which start as stubs):

```js
(function () {
  'use strict';
  var C = window.ECCore;
  var STALE_DAYS = 30;

  var state = {
    data: null,            // { accounts, alliances, meta }
    view: 'cockpit',
    filters: { search:'', classifications:[], priorities:[], alliances:[], interacted:null, quick:null },
    sort: { field:'priority', dir:'desc' },
    now: new Date()
  };

  function el(html){ var t=document.createElement('template'); t.innerHTML=html.trim(); return t.content.firstChild; }
  function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }

  function render(){
    var app = document.getElementById('app');
    app.innerHTML = '';
    app.appendChild(renderHeader());
    if (!state.data){ app.appendChild(renderImportEmpty()); return; }
    app.appendChild(renderNav());
    var view = el('<div class="view"></div>');
    if (state.view==='cockpit') renderCockpit(view);
    else if (state.view==='accounts') renderAccounts(view);
    else if (state.view==='alliances') renderAlliances(view);
    app.appendChild(view);
  }

  function renderHeader(){
    var meta = state.data && state.data.meta;
    var info = state.data
      ? esc(state.fileName||'workbook') + ' · ' + state.data.accounts.length + ' accounts'
      : 'No file loaded';
    var h = el('<div class="app-header">'
      + '<span class="brand-dot"></span>'
      + '<span class="app-title">E&amp;C field service — sector cockpit</span>'
      + '<span class="import-info">'+info+'</span></div>');
    return h;
  }

  function renderNav(){
    var nav = el('<div class="nav"></div>');
    [['cockpit','Cockpit'],['accounts','Accounts'],['alliances','Alliances']].forEach(function(t){
      var b = el('<button'+(state.view===t[0]?' class="active"':'')+'>'+t[1]+'</button>');
      b.onclick = function(){ state.view=t[0]; render(); };
      nav.appendChild(b);
    });
    return nav;
  }

  function renderImportEmpty(){
    var wrap = el('<div class="view"><div class="dropzone" id="dz">'
      + '<p><strong>Import your tracker</strong></p>'
      + '<p>Drag the Excel here, or</p>'
      + '<p><label class="btn primary">Choose file<input id="file" type="file" accept=".xlsx,.xls,.csv" hidden></label></p>'
      + '</div></div>');
    return wrap;
  }

  // View stubs — replaced in later tasks
  function renderCockpit(v){ v.appendChild(el('<div class="empty">Cockpit (todo)</div>')); }
  function renderAccounts(v){ v.appendChild(el('<div class="empty">Accounts (todo)</div>')); }
  function renderAlliances(v){ v.appendChild(el('<div class="empty">Alliances (todo)</div>')); }

  function setData(data, fileName){ state.data=data; state.fileName=fileName; state.view='cockpit'; render(); }

  // Test hook: inject parsed data without a file picker
  window.__ECTEST = {
    loadData: function(data, name){ setData(data, name||'fixture.xlsx'); },
    state: function(){ return state; }
  };

  // expose for later tasks
  window.__ECUI = { render:render, el:el, esc:esc, state:state, STALE_DAYS:STALE_DAYS, setData:setData,
    renderCockpit:renderCockpit, renderAccounts:renderAccounts, renderAlliances:renderAlliances };

  render();
})();
```

- [ ] **Step 3: Build and start the preview server**

Run: `node build.js`
Then start the preview (serves the working dir): use `preview_start` on `EC_FieldService_Tracker.html`.

- [ ] **Step 4: Verify the shell renders the import empty-state**

Use `preview_snapshot`.
Expected: header "E&C field service — sector cockpit", a dropzone with "Import your tracker" and a "Choose file" button. No nav yet (no data).

- [ ] **Step 5: Verify the test hook + nav by injecting fixture data**

Use `preview_eval`:
```js
const XLSXok = typeof XLSX; const Cok = typeof ECCore.parseWorkbook;
window.__ECTEST.loadData({accounts:[{company:'Test',classification:'Upstream',priority:'HIGH',alliances:[]}],alliances:[],meta:{}}, 'demo.xlsx');
document.querySelector('.nav') ? 'nav-rendered' : 'no-nav';
```
Expected: returns `'nav-rendered'`; snapshot shows the three tabs and the import-info "demo.xlsx · 1 accounts".

- [ ] **Step 6: Commit**

```bash
git add src/styles.css src/ui.js
git commit -m "feat(ui): Deloitte-branded shell, nav, import empty-state, test hook"
```

---

### Task 9: Real Excel import (file picker + drag/drop) wired to core.parseWorkbook

**Files:**
- Modify: `src/ui.js`

- [ ] **Step 1: Implement import handling**

Add to `src/ui.js` (call `wireImport()` at the end of `render()` when the dropzone is present; and also support a persistent "Re-import" button in the header once data is loaded):

```js
function readFileToWorkbook(file, cb){
  var reader = new FileReader();
  reader.onload = function(e){
    try {
      var wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
      cb(null, wb);
    } catch(err){ cb(err); }
  };
  reader.onerror = function(){ cb(new Error('Could not read file')); };
  reader.readAsArrayBuffer(file);
}

function handleFile(file){
  if (!file) return;
  readFileToWorkbook(file, function(err, wb){
    if (err){ alert('Import failed: ' + err.message); return; }
    try {
      var data = C.parseWorkbook(wb);
      if (!data.accounts.length){ alert('No accounts found. Is this the right workbook (a "Lean View" sheet)?'); return; }
      try { localStorage.setItem('ec_last_file', file.name); } catch(_){}
      setData(data, file.name);
      showImportSummary(data);
    } catch(err2){ alert('Import error: ' + err2.message); }
  });
}

function showImportSummary(data){
  var m = data.meta;
  var msg = 'Imported ' + data.accounts.length + ' accounts, ' + data.alliances.length + ' alliances.';
  if (!m.hasAllianceColumn) msg += ' (No "Alliance" column found — add it in Excel to enable alliance roll-ups.)';
  var toast = el('<div class="no-print" style="position:fixed;bottom:20px;left:50%;transform:translateX(-50%);'
    + 'background:var(--ink);color:#fff;padding:12px 18px;border-radius:8px;z-index:80;font-size:13px">'+esc(msg)+'</div>');
  document.body.appendChild(toast);
  setTimeout(function(){ toast.remove(); }, 6000);
}

function wireImport(){
  var dz = document.getElementById('dz');
  var fi = document.getElementById('file');
  if (fi) fi.onchange = function(e){ handleFile(e.target.files[0]); };
  if (dz){
    ['dragenter','dragover'].forEach(function(ev){ dz.addEventListener(ev,function(e){e.preventDefault();dz.classList.add('drag');});});
    ['dragleave','drop'].forEach(function(ev){ dz.addEventListener(ev,function(e){e.preventDefault();dz.classList.remove('drag');});});
    dz.addEventListener('drop', function(e){ handleFile(e.dataTransfer.files[0]); });
  }
}
```

Wire it: at the end of `render()`, add `wireImport();`. Also add a "Re-import" `<label class="btn">` with a hidden file input into the header when `state.data` exists, calling `handleFile`.

- [ ] **Step 2: Build**

Run: `node build.js`

- [ ] **Step 3: Verify real import in the browser**

Because a file picker can't be scripted, drive import via `preview_eval` using `fetch` of the served fixture, then feed bytes through the same path:
```js
const buf = await (await fetch('test/fixtures/sample.xlsx')).arrayBuffer();
const wb = XLSX.read(new Uint8Array(buf), {type:'array', cellDates:true});
const data = ECCore.parseWorkbook(wb);
window.__ECTEST.loadData(data, 'sample.xlsx');
data.accounts.length + ' / ' + data.alliances.length;
```
Expected: returns `'27 / 3'`; snapshot shows nav + import-info "sample.xlsx · 27 accounts".

(Ensure the preview server serves `test/fixtures/`. If not reachable, copy `sample.xlsx` beside the HTML for the test, or use `preview_eval` to build a workbook from a small inline AOA via `XLSX.utils.aoa_to_sheet`.)

- [ ] **Step 4: Commit**

```bash
git add src/ui.js
git commit -m "feat(ui): real Excel import via picker + drag/drop with summary"
```

---

## Phase 3 — Accounts view

### Task 10: Accounts table + toolbar filters + sort

**Files:**
- Modify: `src/ui.js`

- [ ] **Step 1: Implement renderAccounts (replace the stub)**

Replace the `renderAccounts` stub. It renders: a toolbar (search input, classification chips, priority chips, alliance chips, interacted toggle, quick-filter chips, clear-all, and the export buttons placeholder), then the filtered+sorted table. Use `C.filterAccounts`, `C.sortAccounts`. Each row click opens the detail drawer (Task 12). Full implementation:

```js
function uniqueClassifications(){ return Object.keys(C.countBy(state.data.accounts,'classification')); }
function priorityPill(p){ var k=(p||'unset').toLowerCase(); return '<span class="pill '+k+'">'+(p||'Unset')+'</span>'; }
function fmtDate(d){ return d ? (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear() : '—'; }

function chip(label, on, onclick){
  var c = el('<button class="chip'+(on?' on':'')+'">'+esc(label)+'</button>');
  c.onclick = onclick; return c;
}

function renderAccounts(v){
  var f = state.filters;
  var tb = el('<div class="toolbar"></div>');
  var search = el('<input type="search" placeholder="Search company, POC, status, pipeline, notes…">');
  search.value = f.search;
  search.oninput = function(){ f.search=search.value; refreshTable(); };
  tb.appendChild(search);
  v.appendChild(tb);

  var chipsRow = el('<div class="toolbar no-print"></div>');
  uniqueClassifications().forEach(function(cl){
    chipsRow.appendChild(chip(cl, f.classifications.indexOf(cl)!==-1, function(){ toggle(f.classifications,cl); render(); }));
  });
  ['HIGH','MEDIUM','LOW','UNSET'].forEach(function(p){
    chipsRow.appendChild(chip(p, f.priorities.indexOf(p)!==-1, function(){ toggle(f.priorities,p); render(); }));
  });
  ['Microsoft','Service Now','Salesforce'].forEach(function(al){
    chipsRow.appendChild(chip(al, f.alliances.indexOf(al)!==-1, function(){ toggle(f.alliances,al); render(); }));
  });
  v.appendChild(chipsRow);

  var quickRow = el('<div class="toolbar no-print"></div>');
  [['needsAction','Needs action'],['meetingSoon','Meeting this week'],['stale','Stale'],['whitespace','Untouched']].forEach(function(q){
    quickRow.appendChild(chip(q[1], f.quick===q[0], function(){ f.quick = (f.quick===q[0]?null:q[0]); render(); }));
  });
  var clear = el('<button class="btn">Clear all</button>');
  clear.onclick = function(){ state.filters={search:'',classifications:[],priorities:[],alliances:[],interacted:null,quick:null}; render(); };
  quickRow.appendChild(clear);
  v.appendChild(quickRow);

  var tableWrap = el('<div id="tablewrap"></div>');
  v.appendChild(tableWrap);
  refreshTable();

  function refreshTable(){
    var rows = C.filterAccounts(state.data.accounts, f, {now:state.now, staleDays:STALE_DAYS});
    rows = C.sortAccounts(rows, state.sort.field, state.sort.dir);
    var cols = [['company','Company'],['classification','Classification'],['priority','Priority'],
      ['interacted','Interacted?'],['nextMeeting','Next meeting'],['alliances','Alliance'],['lastUpdatedDate','Updated']];
    var thead = cols.map(function(c){
      var arrow = state.sort.field===c[0] ? (state.sort.dir==='asc'?' ▲':' ▼') : '';
      return '<th data-f="'+c[0]+'">'+c[1]+arrow+'</th>';
    }).join('');
    var body = rows.map(function(a){
      return '<tr data-id="'+a.id+'">'
        + '<td><strong>'+esc(a.company)+'</strong></td>'
        + '<td>'+esc(a.classification)+'</td>'
        + '<td>'+priorityPill(a.priority)+'</td>'
        + '<td>'+(a.interacted?'Yes':'—')+'</td>'
        + '<td>'+fmtDate(a.nextMeeting)+'</td>'
        + '<td>'+esc((a.alliances||[]).join(', ')||'—')+'</td>'
        + '<td>'+fmtDate(a.lastUpdatedDate)+'</td></tr>';
    }).join('');
    tableWrap.innerHTML = '<div style="margin:8px 0;color:var(--muted);font-size:12px">'+rows.length+' of '+state.data.accounts.length+' accounts</div>'
      + '<table class="accounts"><thead><tr>'+thead+'</tr></thead><tbody>'+(body||'<tr><td colspan="7" class="empty">No matches</td></tr>')+'</tbody></table>';
    tableWrap.querySelectorAll('th[data-f]').forEach(function(th){
      th.onclick = function(){ var fld=th.getAttribute('data-f');
        if (state.sort.field===fld) state.sort.dir = state.sort.dir==='asc'?'desc':'asc';
        else { state.sort.field=fld; state.sort.dir='asc'; }
        refreshTable(); };
    });
    tableWrap.querySelectorAll('tr[data-id]').forEach(function(tr){
      tr.onclick = function(){ openDetail(+tr.getAttribute('data-id')); };
    });
  }
}

function toggle(arr, v){ var i=arr.indexOf(v); if(i===-1)arr.push(v); else arr.splice(i,1); }
```

Add `openDetail` as a temporary stub (`function openDetail(id){ console.log('detail', id); }`) — implemented in Task 12. Expose new helpers on `window.__ECUI` if needed by later tasks (`fmtDate`, `priorityPill`).

- [ ] **Step 2: Build**

Run: `node build.js`

- [ ] **Step 3: Verify in browser**

Inject the fixture (as in Task 9 Step 3), then:
```js
window.__ECUI.state.view='accounts'; window.__ECUI.render();
document.querySelectorAll('table.accounts tbody tr').length;
```
Expected: `27`. Then exercise a filter:
```js
const f = window.__ECUI.state.filters; f.priorities=['HIGH']; window.__ECUI.render();
document.querySelectorAll('table.accounts tbody tr').length;
```
Expected: `9`. `preview_snapshot` shows pills, chips, and the count "9 of 27 accounts".

- [ ] **Step 4: Verify sort via click**

Use `preview_click` on the "Company" header, then `preview_snapshot`; confirm the first row is alphabetical and the header shows the ▲ arrow.

- [ ] **Step 5: Commit**

```bash
git add src/ui.js
git commit -m "feat(ui): accounts table with search, multi-select + quick filters, sort"
```

---

### Task 11: Account detail drawer

**Files:**
- Modify: `src/ui.js`

- [ ] **Step 1: Implement the detail drawer (replace the openDetail stub)**

```js
function emailLinks(text){
  var emails = C.extractEmails(text);
  if (!emails.length) return esc(text) || '—';
  var html = esc(text);
  emails.forEach(function(e){ html = html.replace(e, '<a href="mailto:'+e+'">'+e+'</a>'); });
  return html.replace(/\n/g,'<br>');
}

function openDetail(id){
  var a = state.data.accounts.find(function(x){ return x.id===id; });
  if (!a) return;
  var pipe = C.parsePipeline(a.pipeline);
  var stale = C.isStale(a, state.now, STALE_DAYS);
  var rows = [
    ['Status / next steps', esc(a.status).replace(/\n/g,'<br>')],
    ['Cluster POC', emailLinks(a.clusterPOC)],
    ['Key account POC', emailLinks(a.keyPOC)],
    ['Client POC', emailLinks(a.clientPOC)],
    ['Existing vendors', esc(a.existingVendors)||'—'],
    ['Existing work done', esc(a.existingWork)||'—'],
    ['Action items', esc(a.action).replace(/\n/g,'<br>')||'—'],
    ['Notes', esc(a.notes).replace(/\n/g,'<br>')||'—'],
    ['Last updated', (a.lastUpdatedBy?esc(a.lastUpdatedBy):'—') + ' · ' + fmtDate(a.lastUpdatedDate)]
  ];
  var pipeHtml = pipe.lead || pipe.bullets.length
    ? '<div class="card" style="margin:12px 0"><div style="font-weight:600;margin-bottom:6px">Pipeline</div>'
      + (pipe.lead?'<p style="margin:0 0 8px">'+esc(pipe.lead)+'</p>':'')
      + (pipe.bullets.length?'<ul style="margin:0;padding-left:18px">'+pipe.bullets.map(function(b){return '<li>'+esc(b)+'</li>';}).join('')+'</ul>':'')
      + '</div>' : '';
  var overlay = el('<div class="drawer-overlay"><div class="drawer">'
    + '<button class="btn no-print" id="closeDetail" style="float:right">Close</button>'
    + '<h2 style="margin:0 0 4px">'+esc(a.company)+'</h2>'
    + '<div style="margin-bottom:12px">'+priorityPill(a.priority)+' <span style="color:var(--muted)">'+esc(a.classification)+'</span>'
    + (a.interacted?' · Engaged':'') + (stale?' · <span style="color:#A32D2D">Stale</span>':'')
    + ((a.alliances||[]).length?' · '+esc(a.alliances.join(', ')):'') + '</div>'
    + pipeHtml
    + '<table style="width:100%;border-collapse:collapse">'
    + rows.map(function(r){return '<tr><td style="color:var(--muted);padding:8px 10px 8px 0;vertical-align:top;width:150px">'+r[0]+'</td><td style="padding:8px 0">'+r[1]+'</td></tr>';}).join('')
    + '</table></div></div>');
  overlay.addEventListener('click', function(e){ if (e.target===overlay) overlay.remove(); });
  overlay.querySelector('#closeDetail').onclick = function(){ overlay.remove(); };
  document.body.appendChild(overlay);
}
```

- [ ] **Step 2: Build**

Run: `node build.js`

- [ ] **Step 3: Verify in browser**

Inject fixture, go to accounts, then:
```js
window.__ECUI.state.view='accounts'; window.__ECUI.render();
const id = ECCore.parseWorkbook; // ensure loaded
document.querySelector('table.accounts tbody tr').click();
document.querySelector('.drawer h2').textContent;
```
Expected: returns a company name; `preview_snapshot` shows the drawer with priority pill, classification, pipeline bullets (for an account like ConocoPhillips/HF Sinclair), and POC emails as `mailto:` links.

- [ ] **Step 4: Commit**

```bash
git add src/ui.js
git commit -m "feat(ui): account detail drawer with parsed pipeline + mailto POCs"
```

---

## Phase 4 — Cockpit overview + Alliances view

### Task 12: Cockpit overview (KPIs, charts, needs-attention, whitespace)

**Files:**
- Modify: `src/ui.js`

- [ ] **Step 1: Implement renderCockpit (replace the stub)**

```js
function kpiCard(label, value, accent){
  return '<div class="kpi'+(accent?' accent':'')+'"><div class="label">'+esc(label)+'</div><div class="value">'+value+'</div></div>';
}
function barRow(label, count, max){
  var pct = max ? Math.round((count/max)*100) : 0;
  return '<div class="bar-row"><span class="lab">'+esc(label)+'</span>'
    + '<span class="bar-track"><span class="bar-fill" style="width:'+pct+'%"></span></span>'
    + '<span style="width:20px;text-align:right">'+count+'</span></div>';
}
function attentionList(title, icon, items, render){
  if (!items.length) return '';
  return '<div class="card" style="margin-bottom:16px"><div style="font-weight:600;margin-bottom:8px">'+esc(title)+' ('+items.length+')</div>'
    + items.slice(0,8).map(render).join('') + '</div>';
}

function renderCockpit(v){
  var acc = state.data.accounts;
  var k = C.computeKpis(acc, { now:state.now, staleDays:STALE_DAYS });
  v.innerHTML = '<div class="kpi-grid">'
    + kpiCard('Focus accounts', k.total, true)
    + kpiCard('High priority', k.high)
    + kpiCard('Engaged', k.engagedPct+'%')
    + kpiCard('Meetings ≤'+STALE_DAYS+'d', k.upcomingMeetings)
    + kpiCard('Open actions', k.openActions)
    + kpiCard('Stale >'+STALE_DAYS+'d', k.stale)
    + '</div>';

  // charts row
  var byCls = k.byClassification; var maxCls = Math.max.apply(null, Object.values(byCls).concat([1]));
  var clsHtml = Object.keys(byCls).map(function(c){ return barRow(c, byCls[c], maxCls); }).join('');
  var pr = [['High',k.high,'var(--green)'],['Medium',k.medium,'var(--amber)'],['Low',k.low,'var(--gray)'],['Unset',k.unset,'var(--line)']];
  var prTotal = k.total||1;
  var prBar = '<div style="display:flex;height:16px;border-radius:4px;overflow:hidden;margin-bottom:10px">'
    + pr.map(function(p){return '<span style="width:'+Math.round((p[1]/prTotal)*100)+'%;background:'+p[2]+'"></span>';}).join('')+'</div>'
    + pr.map(function(p){return '<span style="font-size:12px;color:var(--muted);margin-right:14px"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:'+p[2]+'"></span> '+p[0]+' · '+p[1]+'</span>';}).join('');
  var charts = el('<div class="cards-2">'
    + '<div class="card"><div style="font-weight:600;margin-bottom:10px">Portfolio by classification</div>'+clsHtml+'</div>'
    + '<div class="card"><div style="font-weight:600;margin-bottom:10px">Priority mix</div>'+prBar+'</div></div>');
  v.appendChild(charts);

  // needs attention + whitespace
  var now=state.now;
  var meetings = acc.filter(function(a){return C.isUpcomingMeeting(a,now,STALE_DAYS);}).sort(function(a,b){return a.nextMeeting-b.nextMeeting;});
  var staleList = acc.filter(function(a){return C.isStale(a,now,STALE_DAYS);});
  var actions = acc.filter(C.hasOpenAction);
  var whitespace = acc.filter(C.isWhitespace);
  function line(a, right){ return '<div style="display:flex;gap:10px;padding:6px 0;border-bottom:1px solid var(--line);cursor:pointer" data-id="'+a.id+'">'
    + '<span style="font-weight:600;width:150px">'+esc(a.company)+'</span>'+priorityPill(a.priority)
    + '<span style="margin-left:auto;color:var(--muted)">'+right+'</span></div>'; }
  var attn = el('<div></div>');
  attn.innerHTML =
      attentionList('Upcoming meetings','',meetings,function(a){return line(a, fmtDate(a.nextMeeting));})
    + attentionList('Open action items','',actions,function(a){return line(a, esc(C.latestActionDate(a.action)||'open'));})
    + attentionList('Stale accounts','',staleList,function(a){return line(a, 'updated '+fmtDate(a.lastUpdatedDate));})
    + attentionList('Whitespace to target','',whitespace,function(a){return line(a, esc(a.classification));});
  attn.querySelectorAll('[data-id]').forEach(function(n){ n.onclick=function(){ openDetail(+n.getAttribute('data-id')); }; });
  v.appendChild(attn);
}
```

- [ ] **Step 2: Build**

Run: `node build.js`

- [ ] **Step 3: Verify in browser**

Inject fixture; default view is cockpit. `preview_eval`:
```js
document.querySelectorAll('.kpi').length + '|' + document.querySelector('.kpi .value').textContent;
```
Expected: `'6|27'`. `preview_snapshot` shows 6 KPI cards, the two charts, and "Whitespace to target" + "Open action items" panels. `preview_resize` to a narrow width and snapshot to confirm the grids reflow.

- [ ] **Step 4: Verify click-through**

`preview_eval`: `document.querySelector('[data-id]').click(); document.querySelector('.drawer h2')?.textContent;`
Expected: a company name (drawer opened from a cockpit panel).

- [ ] **Step 5: Commit**

```bash
git add src/ui.js
git commit -m "feat(ui): cockpit overview — KPIs, charts, needs-attention, whitespace"
```

---

### Task 13: Alliances view + drill-down

**Files:**
- Modify: `src/ui.js`

- [ ] **Step 1: Implement renderAlliances (replace the stub)**

```js
function renderAlliances(v){
  var groups = C.groupByAlliance(state.data.accounts);
  var byVendor = {};
  state.data.alliances.forEach(function(al){ byVendor[al.vendor]=al; });
  var vendors = ['Microsoft','Service Now','Salesforce'];
  if (!state.data.meta.hasAllianceColumn){
    v.appendChild(el('<div class="card no-print" style="margin-bottom:16px;color:var(--muted)">No "Alliance" column found in the Excel. Add an <strong>Alliance</strong> column to the Lean View sheet (values: Microsoft / Service Now / Salesforce) to see tagged accounts here. Alliance contacts below still work.</div>'));
  }
  var grid = el('<div class="alliance-grid"></div>');
  vendors.forEach(function(name){
    var al = byVendor[name] || {};
    var count = (groups[name]||[]).length;
    var card = el('<div class="card" style="cursor:pointer">'
      + '<div style="font-weight:700;margin-bottom:6px">'+esc(name)+'</div>'
      + '<div style="font-size:24px;font-weight:700">'+count+' <span style="font-size:12px;color:var(--muted);font-weight:400">accounts</span></div>'
      + (al.nextAction?'<div style="font-size:12px;color:var(--muted);margin-top:6px">Next: '+esc(al.nextAction.slice(0,80))+'</div>':'')
      + '</div>');
    card.ondblclick = function(){ openAlliance(name, al, groups[name]||[]); };
    card.onclick = function(){ openAlliance(name, al, groups[name]||[]); }; // single click also opens (dbl-click hint kept)
    grid.appendChild(card);
  });
  v.appendChild(grid);
}

function openAlliance(name, al, accounts){
  function block(label, text){ return '<tr><td style="color:var(--muted);padding:8px 10px 8px 0;vertical-align:top;width:130px">'+label+'</td><td style="padding:8px 0">'+(emailLinks(text)||'—')+'</td></tr>'; }
  var acctRows = accounts.length
    ? accounts.map(function(a){ return '<div style="display:flex;gap:10px;padding:6px 0;border-bottom:1px solid var(--line);cursor:pointer" data-id="'+a.id+'"><span style="font-weight:600;width:160px">'+esc(a.company)+'</span>'+priorityPill(a.priority)+'<span style="margin-left:auto;color:var(--muted)">'+esc((a.status||'').slice(0,60))+'</span></div>'; }).join('')
    : '<div class="empty">No accounts tagged to '+esc(name)+' yet.</div>';
  var overlay = el('<div class="drawer-overlay"><div class="drawer">'
    + '<button class="btn no-print" id="closeAlliance" style="float:right">Close</button>'
    + '<h2 style="margin:0 0 12px">'+esc(name)+'</h2>'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:16px">'
    + block('Internal POC', al.internalPOC) + block('Vendor POC', al.vendorPOC)
    + block('Comment', al.comment) + block('Next action', al.nextAction) + '</table>'
    + '<div style="font-weight:600;margin-bottom:6px">Tagged accounts ('+accounts.length+')</div>'+acctRows
    + '</div></div>');
  overlay.addEventListener('click', function(e){ if (e.target===overlay) overlay.remove(); });
  overlay.querySelector('#closeAlliance').onclick=function(){ overlay.remove(); };
  overlay.querySelectorAll('[data-id]').forEach(function(n){ n.onclick=function(){ overlay.remove(); openDetail(+n.getAttribute('data-id')); }; });
  document.body.appendChild(overlay);
}
```

- [ ] **Step 2: Build**

Run: `node build.js`

- [ ] **Step 3: Verify in browser**

Inject fixture (no Alliance column → the prompt should show). `preview_eval`:
```js
window.__ECUI.state.view='alliances'; window.__ECUI.render();
document.querySelectorAll('.alliance-grid .card').length;
```
Expected: `3`. Then open one:
```js
document.querySelectorAll('.alliance-grid .card')[2].click();
document.querySelector('.drawer h2').textContent;
```
Expected: `'Salesforce'`; snapshot shows Internal/Vendor POC, Comment, Next action, and (since no tags) the "No accounts tagged" empty state.

- [ ] **Step 4: Verify the tagged-account path with a synthetic Alliance column**

`preview_eval` builds a workbook with an Alliance column so the roll-up is exercised:
```js
const aoa=[[],[],['Account Classification','Company Name','Priority Focus','Alliance'],
  ['Upstream','Suncor','HIGH','Salesforce'],['Mining','Alcoa','LOW','Service Now']];
const ws=XLSX.utils.aoa_to_sheet(aoa);
const wb={SheetNames:['Lean View','Alliances'],Sheets:{'Lean View':ws,'Alliances':XLSX.utils.aoa_to_sheet([['Vendor','Internal POC','Vendor POC','Comment','Next Action'],['Salesforce','a@x.com','b@y.com','c','do x']])}};
window.__ECTEST.loadData(ECCore.parseWorkbook(wb),'synthetic.xlsx');
window.__ECUI.state.view='alliances'; window.__ECUI.render();
document.querySelectorAll('.alliance-grid .card')[2].click();
document.querySelectorAll('.drawer [data-id]').length;
```
Expected: `1` (Suncor tagged to Salesforce).

- [ ] **Step 5: Commit**

```bash
git add src/ui.js
git commit -m "feat(ui): alliances view with drill-down + tagged-account roll-up"
```

---

## Phase 5 — Exports, template, docs, retire Code.gs, full verification

### Task 14: Leadership PDF (print) + export current view to Excel

**Files:**
- Modify: `src/ui.js`

- [ ] **Step 1: Add export buttons to the header and implement handlers**

In `renderHeader()`, when `state.data` exists, append a right-aligned button group:

```js
function exportButtons(){
  var wrap = el('<span class="no-print" style="margin-left:12px;display:inline-flex;gap:8px"></span>');
  var pdf = el('<button class="btn">Print / PDF</button>');
  pdf.onclick = function(){ window.print(); };
  var xls = el('<button class="btn primary">Export view to Excel</button>');
  xls.onclick = exportCurrentViewToExcel;
  wrap.appendChild(pdf); wrap.appendChild(xls);
  return wrap;
}

function exportCurrentViewToExcel(){
  var f = state.filters;
  var rows = C.sortAccounts(C.filterAccounts(state.data.accounts, f, {now:state.now, staleDays:STALE_DAYS}), state.sort.field, state.sort.dir);
  var header = ['Account Classification','Company Name','Priority Focus','Status / Next Steps','Interacted With?',
    'Next Meeting Date','Cluster POC','Key Account POC','Client POC','Existing Vendors','Existing Work Done',
    'Vendor Pipeline Status','Action Items','Last Updated Date','Last Updated By','Notes','Alliance'];
  var aoa = [header].concat(rows.map(function(a){
    return [a.classification,a.company,a.priority||'',a.status,a.interacted?'Yes':'',
      a.nextMeetingRaw||(a.nextMeeting?fmtDate(a.nextMeeting):''),a.clusterPOC,a.keyPOC,a.clientPOC,
      a.existingVendors,a.existingWork,a.pipeline,a.action,
      a.lastUpdatedDate?fmtDate(a.lastUpdatedDate):'',a.lastUpdatedBy,a.notes,(a.alliances||[]).join(', ')];
  }));
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Lean View');
  var stamp = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, 'EC_FieldService_View_'+stamp+'.xlsx');
}
```

Wire `renderHeader` to append `exportButtons()` after `.import-info` when `state.data` exists. For a clean leadership print, ensure the cockpit prints well: the `@media print` CSS already hides nav/toolbar/buttons. Optionally add a print-only title line `<h1 class="print-only">…</h1>` (add `.print-only{display:none} @media print{.print-only{display:block}}`).

- [ ] **Step 2: Build**

Run: `node build.js`

- [ ] **Step 3: Verify export logic in the browser (without triggering a download dialog)**

`preview_eval` validates the AOA the exporter would write (refactor `exportCurrentViewToExcel` to build via a helper `buildExportAoa()` that returns the AOA, so it can be asserted; `writeFile` stays in the click handler):
```js
const aoa = window.__ECUI.buildExportAoa ? window.__ECUI.buildExportAoa() : null;
aoa[0].length + '|' + aoa.length;
```
Expected: `'17|<filteredCount+1>'` — 17 columns incl. Alliance, and one header row + one row per filtered account.

(Refactor note: extract `buildExportAoa()` from `exportCurrentViewToExcel` and expose on `window.__ECUI` so this assertion works; the click handler calls `XLSX.writeFile` on its result.)

- [ ] **Step 4: Verify print CSS**

`preview_eval`: `getComputedStyle(document.querySelector('.nav')).display;` returns the screen value; then rely on the `@media print` rule. Optionally use `preview_screenshot` after `preview_eval("matchMedia('print')")` is not reliable — instead visually confirm via snapshot that nav/toolbar carry the `no-print` class.

- [ ] **Step 5: Commit**

```bash
git add src/ui.js
git commit -m "feat(ui): leadership print/PDF + export current view to Excel"
```

---

### Task 15: Last-view persistence + empty/error states polish

**Files:**
- Modify: `src/ui.js`

- [ ] **Step 1: Persist last view + filters and restore on next import**

Add a `persist()` that writes `{view, filters, sort}` to `localStorage.ec_view_state` and call it whenever they change; on `setData`, read it back and apply (guard against unknown classifications). Full code:

```js
function persist(){ try{ localStorage.setItem('ec_view_state', JSON.stringify({view:state.view,filters:state.filters,sort:state.sort})); }catch(_){}}
function restore(){ try{ var s=JSON.parse(localStorage.getItem('ec_view_state')||'null'); if(s){ state.view=s.view||'cockpit'; state.filters=Object.assign(state.filters,s.filters||{}); state.sort=s.sort||state.sort; } }catch(_){}}
```

Call `restore()` inside `setData` before `render()`. Call `persist()` at the end of `render()`. Also show the remembered file name from `localStorage.ec_last_file` as a hint in the import empty-state ("Last imported: …").

- [ ] **Step 2: Build, then verify persistence**

Run: `node build.js`. In the browser: inject fixture, switch to accounts, set a filter, then `preview_eval('location.reload()')` is not enough (data is gone after reload since it's session-only). Instead verify `localStorage.getItem('ec_view_state')` is populated:
```js
window.__ECUI.state.view='accounts'; window.__ECUI.state.filters.priorities=['HIGH']; window.__ECUI.render();
JSON.parse(localStorage.getItem('ec_view_state')).filters.priorities[0];
```
Expected: `'HIGH'`.

- [ ] **Step 3: Verify error states**

`preview_eval` feed a junk workbook:
```js
try { ECCore.parseWorkbook({SheetNames:['X'],Sheets:{X:XLSX.utils.aoa_to_sheet([['nope']])}}); 'no-throw'; } catch(e){ 'threw:'+e.message; }
```
Expected: `'no-throw'` and the returned accounts length is 0 (the UI shows the "No accounts found" alert path on real import).

- [ ] **Step 4: Commit**

```bash
git add src/ui.js
git commit -m "feat(ui): last-view persistence + empty/error polish"
```

---

### Task 16: Updated Excel template with Alliance column + dropdown

**Files:**
- Create: `scripts/make-template.js`
- Create: `EC_Accounts_Tracking_Planner_TEMPLATE.xlsx` (generated)

- [ ] **Step 1: Write a generator that adds the Alliance column to a copy of the real workbook**

Because SheetJS community build does not write data-validation dropdowns, generate the template by (a) copying the fixture, (b) appending an "Alliance" header at column R on the Lean View header row, and (c) writing a sibling `README` note instructing how to add the Excel dropdown manually (Data ▸ Data Validation ▸ List: `Microsoft,Service Now,Salesforce`). Create `scripts/make-template.js`:

```js
const fs = require('fs');
const XLSX = require('../vendor/xlsx.full.min.js');
const wb = XLSX.readFile(__dirname + '/../test/fixtures/sample.xlsx', { cellDates:true });
const ws = wb.Sheets['Lean View'];
const range = XLSX.utils.decode_range(ws['!ref']);
const headerRow = 2; // 0-based row index 2 = Excel row 3
const col = range.e.c + 1; // append after last column
const addr = XLSX.utils.encode_cell({ r: headerRow, c: col });
ws[addr] = { t:'s', v:'Alliance' };
range.e.c = col;
ws['!ref'] = XLSX.utils.encode_range(range);
XLSX.writeFile(wb, __dirname + '/../EC_Accounts_Tracking_Planner_TEMPLATE.xlsx');
console.log('Template written with Alliance column at', addr);
```

- [ ] **Step 2: Generate the template**

Run: `node scripts/make-template.js`
Expected: prints "Template written with Alliance column at R3" (or the correct cell address).

- [ ] **Step 3: Verify the app reads the template and reports hasAllianceColumn:true**

Run: `node -e "const X=require('./vendor/xlsx.full.min.js'); const c=require('./src/core.js'); const wb=X.readFile('./EC_Accounts_Tracking_Planner_TEMPLATE.xlsx',{cellDates:true}); const d=c.parseWorkbook(wb); console.log(d.meta.hasAllianceColumn, d.accounts.length);"`
Expected: `true 27`

- [ ] **Step 4: Commit**

```bash
git add scripts/make-template.js EC_Accounts_Tracking_Planner_TEMPLATE.xlsx
git commit -m "feat: updated Excel template with Alliance column"
```

---

### Task 17: README + retire Code.gs

**Files:**
- Create: `README.md`
- Move: `Code.gs` → `archive/Code.gs`

- [ ] **Step 1: Write the README**

Create `README.md` covering: what the app is; the workflow (maintain Excel in SharePoint → open the HTML → import → use cockpit → export PDF/Excel); how the Alliance column works + how to add the Excel dropdown (Data Validation ▸ List `Microsoft,Service Now,Salesforce`); how to deploy (upload `EC_FieldService_Tracker.html` to the SharePoint document library, open in the browser); the dev workflow (`node --test`, `node build.js`); and an explicit note that the app is view-only and the Excel co-authored in SharePoint is the source of truth.

- [ ] **Step 2: Retire Code.gs**

Run (PowerShell):
```powershell
New-Item -ItemType Directory -Force "C:\Custom Projects\ec-fsm-oppty-tracker\archive" | Out-Null
Move-Item "C:\Custom Projects\ec-fsm-oppty-tracker\Code.gs" "C:\Custom Projects\ec-fsm-oppty-tracker\archive\Code.gs" -Force
```
Add a one-line header note in `archive/Code.gs` (or a sibling `archive/README.md`) that this Google Apps Script backend is retired (Google services are out of scope; the app is now SharePoint + Excel only).

- [ ] **Step 3: Commit**

```bash
git add README.md archive/
git commit -m "docs: README + retire Code.gs to archive"
```

---

### Task 18: Full end-to-end verification against the real workbook

**Files:** none (verification only)

- [ ] **Step 1: Run the full test suite**

Run: `node --test`
Expected: all suites PASS, 0 failures.

- [ ] **Step 2: Rebuild and confirm a clean single-file artifact**

Run: `node build.js`
Run: `node -e "const h=require('fs').readFileSync('EC_FieldService_Tracker.html','utf8'); console.log('size',h.length,'markers',h.includes('INLINE:'),'extScripts',/<script[^>]+src=/.test(h));"`
Expected: size > 800000, markers `false`, extScripts `false`.

- [ ] **Step 3: Full browser walkthrough**

With the preview server running, inject the real fixture and verify each view in sequence using `preview_eval` + `preview_snapshot` + `preview_screenshot`:
- Cockpit: 6 KPIs (Focus accounts = 27, High = 9), both charts, whitespace + open-actions panels.
- Accounts: 27 rows; apply search "agentforce" → HF Sinclair present; priority chip HIGH → 9 rows; open a row → drawer with pipeline + mailto links.
- Alliances: 3 cards; open Salesforce → POC/comment/next-action; no-tag prompt shown for the real file.
- Exports: `buildExportAoa()` returns 17 columns; `window.print()` triggers print view (nav/toolbar hidden).
Capture a `preview_screenshot` of the cockpit for the user.

- [ ] **Step 4: Stop the preview server**

Use `preview_stop`.

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "chore: final build + verified end-to-end against real workbook"
```

---

## Self-Review (completed by plan author)

**Spec coverage:**
- Upload-driven / Excel-is-truth, single self-contained file → Tasks 1, 9, 18. ✓
- View-only (no editing) → enforced; only export, never write-back to source. ✓
- Three views (Cockpit/Accounts/Alliances + drill-down) → Tasks 10–13. ✓
- Data model incl. optional Alliance column → Tasks 3, 5, 16. ✓
- KPIs from existing fields only → Task 6 (no value/stage fields introduced). ✓
- Powerful filtration (search, multi-select, quick filters, sort) → Tasks 7, 10. ✓
- Whitespace-to-target → Tasks 6, 12. ✓
- Alliance tag maintained in Excel + template delivered → Tasks 5, 16. ✓
- Exports: leadership PDF + export-current-view to Excel → Task 14. ✓
- Light Deloitte branding → Task 8. ✓
- Robust import + summary + errors → Tasks 5, 9, 15. ✓
- Deliverables (HTML, template, README, retire Code.gs) → Tasks 16, 17, 18. ✓

**Placeholder scan:** View stubs in Task 8 are intentional and explicitly replaced in Tasks 10/12/13; no "TBD"/"add error handling" hand-waving — error paths are concrete (Task 9 alert paths, Task 15 junk-workbook test). ✓

**Type/name consistency:** `parseWorkbook` returns `{accounts, alliances, meta}` used consistently; `computeKpis(accounts,{now,staleDays})` signature consistent across Tasks 6/12; `filterAccounts(accounts, criteria, ctx)` consistent across Tasks 7/10/14; alliance canonical form is `'Service Now'` (two words) everywhere (core canon, groupByAlliance buckets, UI chips, template note) — matches the fixture's Alliances sheet value. `buildExportAoa()` extraction noted in Task 14 so the Task 14/18 assertions resolve. ✓
