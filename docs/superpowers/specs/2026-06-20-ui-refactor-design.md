# E&C Field Service Tracker — UI Refactor Design Spec

**Date:** 2026-06-20
**Author:** micschuler@deloitte.com (with Claude)
**Status:** Approved for planning
**Supersedes:** Visual/brand section of `2026-06-11-ec-fsm-sector-cockpit-design.md`

---

## 1. Purpose

Refactor `EC_FieldService_Tracker.html` from its current basic tab-based layout into a
polished Salesforce-style CRM dashboard, using the provided UX mockup as a visual
and structural guide — adapted to the app's actual data model (accounts tracker, not
a sales pipeline). Data logic is unchanged; only the visual shell and layout are redesigned.

---

## 2. Approach

**Option 1 — Extract → Refactor → Rebuild.** Decompose the existing monolithic HTML into
clean source files, write new CSS and HTML structure, keep core data logic, reassemble
via `build.js`. This implements the planned `src/` architecture from the original spec
as part of the visual refactor.

---

## 3. File Structure

```
src/
  app.template.html     HTML skeleton with <!--INLINE:styles-->, <!--INLINE:core-->,
                        <!--INLINE:ui-->, <!--INLINE:xlsx--> markers
  styles.css            Complete visual design — layout, sidebar, cards, tables,
                        pills, slide-in panel, charts, print CSS
  core.js               Pure data logic (UMD, no DOM) — parse workbook, normalize
                        rows, compute KPIs, filter/sort. Extracted from existing HTML,
                        minimally modified.
  ui.js                 All DOM rendering & event wiring — sidebar nav, view routing,
                        home/accounts/alliances views, slide-in panel, import flow.
                        Full rewrite for new layout.
vendor/
  xlsx.full.min.js      SheetJS blob extracted one-time from existing HTML
build.js                Reads template, inlines vendor + src → EC_FieldService_Tracker.html
                        and docs/index.html (GitHub Pages artifact)
```

The extraction script reads the existing HTML once to pull out:
- SheetJS blob → `vendor/xlsx.full.min.js`
- Data parsing / normalization functions → `src/core.js`
- Render / event functions → `src/ui.js` (starting point; gets rewritten)

---

## 4. Visual Design System

### 4.1 Colors
| Token | Value | Usage |
|---|---|---|
| `--color-bg` | `#F3F3F3` | Page background |
| `--color-surface` | `#FFFFFF` | Cards, sidebar, panels |
| `--color-border` | `#E5E7EB` | Dividers, card borders |
| `--color-text-primary` | `#1A1A1A` | Body text, headings |
| `--color-text-secondary` | `#6B7280` | Metadata, labels |
| `--color-green` | `#86BC25` | Active nav, primary buttons, HIGH pill, chart accent |
| `--color-amber` | `#EF9F27` | MEDIUM priority pill |
| `--color-gray-pill` | `#6B7280` | LOW priority pill |
| `--color-row-hover` | `#F9FAFB` | Table row hover |

**Priority pill colors** (background is 15% opacity tint of the base color):
- HIGH: green text on green tint
- MEDIUM: amber text on amber tint
- LOW: gray text on gray tint
- Unset: `#9CA3AF` text on `#F3F3F3`

**Classification chip colors** — 5 distinct muted chips (no green/amber collision):
- ACP Accounts: indigo tint
- Upstream: sky tint
- Mid / Downstream: teal tint
- ETS / OFS: purple tint
- Mining: orange tint

### 4.2 Typography
- Font stack: `'Open Sans', 'Segoe UI', system-ui, sans-serif`
- Open Sans loaded via Google Fonts `<link>` in template; degrades gracefully
- Sizes: 11px (metadata), 13px (table rows), 14px (body), 16px (card values), 22px (KPI numbers), 18px (section headers)
- Weights: 400 regular, 600 semibold, 700 bold

### 4.3 Spacing & Shape
- Sidebar: `220px` fixed width, full viewport height, `box-shadow: 2px 0 8px rgba(0,0,0,0.06)`
- Cards: `border-radius: 8px`, `box-shadow: 0 1px 3px rgba(0,0,0,0.08)`
- Table rows: `44px` min height, `1px` bottom border
- Slide-in panel: `420px` fixed right, full viewport height, `box-shadow: -4px 0 16px rgba(0,0,0,0.10)`
- KPI cards: 4-column grid, equal width

---

## 5. Layout & Navigation

```
┌──────────────────────────────────────────────────────────────┐
│ Sidebar (220px fixed)  │  Main content (flex-1, scrollable)  │
│                        │                                      │
│  Deloitte.             │  [Active view renders here]          │
│                        │                                      │
│  ● Home                │                                      │
│    Accounts            │                                      │
│    Alliances           │                                      │
│                        │                                      │
│  ────────────────      │                                      │
│  ⬆ Import Excel        │                                      │
│  filename.xlsx         │                                      │
└──────────────────────────────────────────────────────────────┘
```

- Sidebar fixed, always visible. Active item: left green `3px` bar + green text + `#F3F3F3` pill background.
- Import button at sidebar bottom. Last-imported filename shown below in secondary text.
- No top header bar. View title is the view's own section heading.
- Drag-and-drop on the main content area still supported.

### 5.1 Empty / Import State
On first load (no file imported): sidebar visible, nav items disabled (grayed, non-clickable),
main area shows centered drop zone — "Drop your Excel here or click Import Excel" with
a file icon and secondary instruction text. After import, toast confirms ("Imported 42
accounts, 3 alliances"), nav enables, lands on Home.

---

## 6. Home View

```
Home
──────────────────────────────────────────────────────────────
[Total Accounts]  [High Priority]  [Engaged]   [Meetings ≤30d]
42                12               68%          5
                  28% of total     32% unengaged  next: Jun 24

──────────────────────────────────────────────────────────────
[By Classification]   [Priority Mix]      [Alliance Split]
SVG bar chart         SVG stacked bar     SVG bar chart

──────────────────────────────────────────────────────────────
High Priority Accounts
Company | Classification | Priority | Next Meeting | Last Updated
...
[View all accounts →]
```

### 6.1 KPI Cards
Four cards in a grid row:
| Card | Value | Subtext |
|---|---|---|
| Total Accounts | Row count | "(N high priority)" |
| High Priority | `priority == HIGH` count | "N% of total" |
| Engaged | `interacted == Yes` / total | "N% not yet engaged" |
| Meetings ≤30d | `nextMeeting` within today + 30 | "next: [date]" or "none scheduled" |

Each card: icon (inline SVG, green), large number, label, subtext.

### 6.2 Mini-Charts Row
Three equal-width SVG bar charts rendered inline:
- **By Classification:** horizontal bars, one per classification value, count labels
- **Priority Mix:** stacked horizontal bar (HIGH green / MEDIUM amber / LOW gray / Unset light)
- **Alliance Split:** horizontal bars for Microsoft / ServiceNow / Salesforce counts; "No alliance tags" empty state if column absent

Clicking a bar sets the corresponding filter and navigates to Accounts view.

### 6.3 High Priority Accounts Table
- Filtered to `priority == HIGH`, sorted by `nextMeeting` ascending (soonest first), then alpha
- Columns: Company, Classification chip, Priority pill, Next Meeting, Last Updated
- Clicking a row opens the slide-in detail panel (same panel as Accounts view)
- "View all accounts →" link at bottom navigates to Accounts view with no filters applied

---

## 7. Accounts View

### 7.1 Toolbar
- **Search input** — searches company, all POC fields, status, pipeline, notes (debounced 200ms)
- **Filter chips** — Classification (multi-select), Priority (multi-select), Alliance (multi-select), Interacted (Yes / No toggle)
- **Sort** — clickable column headers, toggle asc/desc, default: Company A–Z
- **Record count** — "42 accounts" / "12 of 42 accounts" when filtered
- **Clear filters** — appears when any filter active

### 7.2 Table
Columns: Company, Classification chip, Priority pill, Interacted badge (✓ / –), Next Meeting, Last Updated, Alliance tags.
- Row height 44px, hover `#F9FAFB`
- Clicking a row opens slide-in panel; selected row gets a green left border
- Table scrolls independently; sidebar and panel stay fixed

### 7.3 Slide-In Detail Panel
420px fixed to right edge, full viewport height. Opens by sliding in from right (CSS transition 200ms). Clicking × or pressing Escape closes it. Clicking another row switches content without closing.

**Panel layout — two columns:**

Left column:
- Company name (heading) + Classification chip + Priority pill
- Interacted badge + Next Meeting date
- Cluster POC, Key Account POC, Client POC — each on a line; emails extracted and rendered as `mailto:` links
- Existing Vendors, Existing Work (collapsed to 3 lines with "show more" if long)

Right column:
- Status / Next Steps (full text, scrollable if long)
- Pipeline bullets — parsed from `Pipeline:` block in the pipeline field; rendered as a bullet list
- Action items — `[MM/DD]` dated items extracted from action field, listed newest-first

Panel footer:
- Alliance tags
- "Last updated [date] by [name]"

---

## 8. Alliances View

### 8.1 Cards
One card per alliance (Microsoft, ServiceNow, Salesforce). Each card:
- Alliance name + logo placeholder icon
- Internal POC name (email linked)
- Vendor POC name (email linked)
- Tagged account count + priority breakdown (N high / M medium)
- Next Action preview (first line only)

"No Alliances sheet found" empty state if sheet missing.
"No alliance tags yet — add an Alliance column in Excel" prompt if sheet present but no accounts tagged.

### 8.2 Drill-Down
Clicking a card expands or navigates to a detail view:
- Full Internal POC, Vendor POC, Comment, Next Action (all emails linked)
- Table of tagged accounts: Company, Priority pill, Status (truncated), last updated — each row clickable to slide-in panel
- "← Back to Alliances" link

---

## 9. Build Process

`build.js` (Node.js, ESM):
1. Read `src/app.template.html`
2. Replace `<!--INLINE:styles-->` with `<style>` + contents of `src/styles.css` + `</style>`
3. Replace `<!--INLINE:xlsx-->` with `<script>` + contents of `vendor/xlsx.full.min.js` + `</script>`
4. Replace `<!--INLINE:core-->` with `<script>` + contents of `src/core.js` + `</script>`
5. Replace `<!--INLINE:ui-->` with `<script>` + contents of `src/ui.js` + `</script>`
6. Write result to `EC_FieldService_Tracker.html` (root) and `docs/index.html` (GitHub Pages)

Run: `node build.js` — no npm install, no toolchain.

---

## 10. Extraction Plan (one-time setup step)

A one-time `scripts/extract.js` script reads the existing `EC_FieldService_Tracker.html` and:
- Finds the SheetJS script block (identifiable by `XLSX` global) → writes `vendor/xlsx.full.min.js`
- Finds all other `<script>` content → writes as starting point for `src/core.js` + `src/ui.js`
- Finds all `<style>` content → writes `src/styles.css.bak` (reference only; new CSS is written fresh)

Manual cleanup of `core.js` / `ui.js` split happens after extraction.

---

## 11. Non-Goals (unchanged from original spec)
- No backend, DB, auth, or live multi-user sync
- No in-app editing / CRUD
- No new data fields (value, stage, owner, close date)
- No Google Sheets / external services
- No dollar-value pipeline (data doesn't have it)

---

## 12. Implementation Phases

| Phase | Work |
|---|---|
| 0 — Extraction | Run extract.js; split into core.js / ui.js; verify build.js assembles correctly |
| 1 — CSS shell | Write styles.css: sidebar layout, empty state, KPI cards, table, pills, slide-in panel |
| 2 — Home view | KPI computation in core.js; home rendering + mini-charts in ui.js |
| 3 — Accounts view | Table, search/filter, sort, slide-in panel wiring |
| 4 — Alliances view | Cards + drill-down restyled |
| 5 — Polish | Empty states, transitions, print CSS, GitHub Pages deploy |
