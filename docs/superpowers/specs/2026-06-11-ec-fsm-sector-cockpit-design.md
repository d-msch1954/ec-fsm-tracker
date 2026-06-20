# E&C Field Service — Sector Cockpit (Design Spec)

**Date:** 2026-06-11
**Author:** micschuler@deloitte.com (with Claude)
**Status:** Approved for planning

---

## 1. Purpose

A single, self-contained HTML dashboard that turns the team's E&C field-service
account tracker (an Excel workbook) into an intuitive "cockpit" for the field
service sector lead. The user imports the current Excel and instantly gets a
filterable portfolio view, an alliance lens with drill-down, and leadership-ready
exports. It is a **viewer and reporting lens** over the Excel — not an editing or
database system.

## 2. Users & context

- **Team:** 2–3 people who maintain the tracker.
- **Hosting:** A single `.html` file placed in an internal **SharePoint** document
  library, opened in the browser (Edge) by each user.
- **Source of truth & collaboration:** The **Excel workbook in SharePoint**, which
  the team **co-authors** (SharePoint real-time co-authoring + version history).
  The dashboard does NOT replace or sync this — it reads whatever Excel the user
  imports.
- **Primary persona:** the sector lead who wants a "cockpit" glance — where to push
  next, what needs attention, how alliances map to accounts — plus a printable
  readout for leadership.

## 3. Architecture & constraints

- **Single self-contained HTML file**, vanilla JS, no build step, no runtime
  network dependencies.
  - **SheetJS inlined** (as in the existing v1) so XLSX parsing works on a
    locked-down SharePoint with no CDN/CSP reliance.
  - **Charts hand-built in lightweight SVG/CSS** — no chart library (avoids bloat
    and CSP issues).
  - **Fonts:** `'Open Sans', 'Segoe UI', system-ui, sans-serif`. Attempt to load
    Open Sans via Google Fonts but degrade gracefully to Segoe UI if the network
    blocks it (do not depend on it).
- **No backend, no database, no authentication, no live multi-user sync, no
  in-app editing/CRUD.**
- **Identity:** none in-app. `Last Updated By` / `Last Updated Date` are existing
  Excel columns maintained by the team in Excel; the app only reads/displays them.
- Evolves the existing `EC_FieldService_Tracker.html`; **retires `Code.gs`** (the
  Google Apps Script backend is out of scope — Google services are disallowed).

### Rejected alternatives (recorded)
- **React/Vite build** — needs a toolchain; bundle/asset loading risks SharePoint
  CSP issues; overkill for a 2–3 person internal tool.
- **SharePoint-native (SPFx / Power Apps)** — needs SP dev tooling + IT involvement
  and reintroduces the backend complexity we are deliberately avoiding.

## 4. Data model

The workbook has two sheets. Column mapping is by **header name** (not position)
for resilience.

### 4.1 "Lean View" sheet (accounts)
Headers are on **row 3** in the source file (the table starts at B3). The parser
detects the header row by scanning the first ~6 rows for known header labels, then
maps columns by name.

| Field (internal) | Excel header | Notes |
|---|---|---|
| `classification` | Account Classification | 5 values: `ACP Accounts`, `Upstream`, `Mid / Downstream`, `ETS / OFS`, `Mining` |
| `company` | Company Name | display key; may repeat — handle gracefully |
| `priority` | Priority Focus | HIGH / MEDIUM / LOW / blank (case-insensitive) |
| `status` | Status / Next Steps | rich multi-line text |
| `interacted` | Interacted With? | "Yes" / blank → boolean |
| `nextMeeting` | Next Meeting Date | date or text; parse tolerantly |
| `clusterPOC` | Cluster POC | may contain name + email |
| `keyPOC` | Key Account POC | may contain name + email |
| `clientPOC` | Client POC | may contain name + email |
| `existingVendors` | Existing Vendors | text |
| `existingWork` | Existing Work Done | text |
| `pipeline` | Vendor Pipeline Status | often contains a `Pipeline:` bullet block |
| `action` | Action Items | often dated like `[06/04] ...` |
| `lastUpdatedDate` | Last Updated Date | date or text |
| `lastUpdatedBy` | Last Updated By | text |
| `notes` | Notes | text |
| `alliance` | **Alliance** (NEW) | optional; multi-value (Microsoft/ServiceNow/Salesforce), comma/semicolon separated |

**Alliance column** is new. The team maintains it in Excel (a delivered template
adds the column with a data-validation dropdown). The app reads it; if the column
is absent (older file), alliance roll-ups show a "no tags yet — add an Alliance
column in Excel" prompt and the rest of the app works unchanged.

### 4.2 "Alliances" sheet
Headers on row 1. Three rows today: Microsoft, ServiceNow, Salesforce.

| Field | Excel header |
|---|---|
| `vendor` | Vendor |
| `internalPOC` | Internal POC |
| `vendorPOC` | Vendor POC |
| `comment` | Comment |
| `nextAction` | Next Action |

POC fields contain one or more `Name — email` lines; the app extracts emails and
renders them as `mailto:` links.

## 5. Import behavior

- File picker + drag-and-drop. Remembers last imported file name (localStorage).
- Locate "Lean View" by case-insensitive trimmed name; fall back to first sheet.
- Locate "Alliances" by name; if missing, alliance view shows an empty state.
- Detect header rows by label scan; map by header name; tolerate reordered/missing
  columns and the optional Alliance column.
- Parse dates tolerantly (Excel serials, JS Dates, common string formats); blanks
  allowed.
- **Import summary toast/panel:** "Imported N accounts, M alliances. Header row
  detected at row R. Unmapped columns: …" plus clear errors for wrong/empty/
  non-xlsx files.

## 6. Views

### 6.1 Cockpit (overview) — landing screen
- **KPI strip** (all derived from existing fields):
  - Total focus accounts = row count.
  - High priority = count `priority == HIGH`.
  - Engaged % = `interacted == Yes` / total.
  - Meetings ≤30 days = count where `nextMeeting` parses within `[today, today+30]`.
  - Open actions = count where `action` non-empty.
  - Stale >N days = count where `lastUpdatedDate` present and older than N
    (default **N = 30**, defined as a constant).
- **Mini-charts (SVG/CSS):** portfolio by classification (bar), priority mix
  (stacked bar/segments: High/Medium/Low/Unset), alliance split (counts per
  Microsoft/ServiceNow/Salesforce).
- **Needs-attention panels** with click-through to the account detail:
  - Upcoming meetings (next 30 days, sorted by date).
  - Stale accounts (older than N days).
  - Open action items (showing the latest `[MM/DD]` extracted from the text).
  - **Whitespace to target** — accounts with **blank priority AND blank status**
    (untouched targets), surfaced as a dedicated panel for "where to push next."

### 6.2 Accounts — the workhorse
- **Filtration (explicit user priority):**
  - Free-text search across company, all POCs, status, notes, pipeline.
  - Multi-select filter chips: Classification, Priority, Alliance, Interacted?.
  - Attention quick-filters: needs action, meeting this week, stale, untouched.
  - Sort on any column; "clear all filters."
- **Table** columns (sensible default subset, all fields available in detail):
  Company, Classification, Priority, Interacted?, Next Meeting, Alliance,
  Last Updated. Priority is color-coded.
- **Account detail** (panel/modal) renders every field cleanly:
  - Parsed `Pipeline:` bullets, dated action items, POC emails as `mailto:` links.
  - Priority + stale badges, last-updated stamp, alliance tags.

### 6.3 Alliances — cards + drill-down
- Cards for Microsoft / ServiceNow / Salesforce showing account counts and a
  short next-action preview.
- **Double-click → alliance detail:** Internal POC, Vendor POC, Comment, Next
  Action (emails linked), plus a roll-up of every tagged account with priority +
  status, click-through to account detail.
- If no Alliance tags exist, show the standalone alliance info + the "add column"
  prompt.

## 7. Exports
- **Leadership PDF / print snapshot:** print-optimized layout (`@media print` +
  `window.print()`) — KPI summary + the currently-filtered account list (and an
  alliance summary). Designed for a sector-lead readout.
- **Export current view to Excel:** SheetJS writes the currently filtered/sorted
  accounts in the canonical Lean View column order; filename stamped with date.

## 8. Visual / brand
- **Light, on-brand Deloitte.** White/light surfaces; **Deloitte Green `#86BC25`**
  as the primary accent (KPI accent, active nav, primary chart fills, header
  circular motif). Black / `#282728` text. Open Sans. ADA-compliant contrast.
- Priority color scale: High = `#86BC25`, Medium = amber `#EF9F27`, Low = gray,
  Unset = light gray.
- One small circular motif in the header (no overlapping circles / Mastercard
  pattern). No logo lockup misuse.

## 9. Non-goals (YAGNI)
No backend, DB, auth, or live multi-user sync. No in-app editing/CRUD. No new
tracking fields (value, stage, owner, close date). No Google Sheets / external
services. No SPFx / Power Apps. No real-time in-app collaboration.

## 10. Edge cases & error handling
Wrong/empty/non-xlsx file; missing "Lean View" or "Alliances" sheet; missing or
reordered columns; missing Alliance column; malformed/blank dates; duplicate
company names; very long text fields; blank workbook. Each fails gracefully with a
clear message; the app never hard-crashes on a bad import.

## 11. Deliverables
1. `EC_FieldService_Tracker.html` — rebuilt, self-contained dashboard.
2. Updated **Excel template** with the `Alliance` column + data-validation dropdown.
3. Short README: how to use + how to deploy to SharePoint.
4. Retire `Code.gs` (remove/mark obsolete). Keep `.claude/serve.js` + `launch.json`
   for local dev/testing only.

## 12. Build order (vertical slices, each testable against the real Excel)
- **Phase 0 — Foundation & import:** HTML scaffold + Deloitte CSS shell + nav;
  robust SheetJS import for both sheets; data-model normalization; import summary +
  errors; deliver the updated Excel template; retire `Code.gs`.
- **Phase 1 — Accounts view:** table, search, multi-select filters, attention
  quick-filters, account detail drill-in.
- **Phase 2 — Cockpit overview:** KPI strip, three mini-charts, needs-attention
  panels (incl. whitespace) with click-through.
- **Phase 3 — Alliances view:** cards → double-click drill-down with tagged-account
  roll-up.
- **Phase 4 — Exports & polish:** Leadership PDF (print-CSS), export-to-Excel,
  last-view/filter persistence, empty/responsive states, accessibility.
- **Phase 5 — Verify & hand off:** test against the real workbook in-browser;
  write the use + SharePoint-deploy note.

## 13. Open assumptions
- KPI numbers in the mockup were illustrative; real values come from the imported
  file.
- `N = 30` days is the default staleness/upcoming-meeting window (adjustable
  constant).
- Git is not initialized in this workspace, so the spec is saved to disk but not
  committed.
