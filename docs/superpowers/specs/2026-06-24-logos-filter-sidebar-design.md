# Design Spec: Company Logos + Filter Sidebar
**Date:** 2026-06-24  
**Project:** EC FSM Opportunity Tracker  
**Scope:** Two UI enhancements to the Accounts view

---

## 1. Company Logos

### Goal
Display each company's official logo next to its name in the accounts table and in the detail panel header, giving the tracker a richer, more recognizable visual identity.

### Data Source
Logos are fetched from the Clearbit Logo API:
```
https://logo.clearbit.com/{domain}
```
No API key required. Free tier is sufficient for this use case (~50 unique companies).

### Domain Inference
Domain resolution happens in two layers:

**Layer 1 — Manual map** (for companies where name ≠ domain):
| Company name | Domain |
|---|---|
| SLB | slb.com |
| TotalEnergies | totalenergies.com |
| Phillips 66 | phillips66.com |
| HF Sinclair | hfsinclair.com |
| TC Energy | tcenergy.com |
| TechnipFMC | technipfmc.com |
| Baker Hughes | bakerhughes.com |
| ConocoPhillips | conocophillips.com |
| Thyssenkrupp | thyssenkrupp.com |
| Rio Tinto | riotinto.com |

**Layer 2 — Algorithmic fallback:**
Lowercase the company name, strip non-alphanumeric characters (except hyphens), collapse spaces to nothing, append `.com`.  
Example: `"Marathon Petroleum"` → `marathonpetroleum.com`

### Fallback: Initials Avatar
If the Clearbit image fails to load (404, network error, CORS block), an initials avatar is shown instead:
- A circle the same size as the logo thumbnail
- Background color deterministically derived from the company name (hash → HSL with fixed saturation/lightness)
- Up to 2 initials extracted from the company name (first letter of first two words)
- White text, same font as the app

This ensures the table always looks consistent regardless of logo availability.

### Rendering

**In the accounts table:**
- Logo or avatar: 32×32px, `border-radius: 6px`, vertically centered with the company name
- Company name sits to the right of the logo with ~10px gap
- The "Company" column widens slightly to accommodate (~180px min)

**In the detail panel:**
- Logo or avatar: 40×40px, shown to the left of the company name in the panel header
- Same fallback logic applies

### Implementation Notes
- Logo `<img>` elements set `onerror` to swap in the initials avatar (pure DOM, no JS framework needed)
- Initials avatar rendered as a `<span>` with inline background-color and text
- Domain inference and avatar generation extracted into helpers in `core.js` (pure functions, no DOM)
- The logo `<img>` uses `loading="lazy"` to avoid blocking initial table render

---

## 2. Filter Sidebar

### Goal
Replace the flat, unlabeled chip row with a structured sidebar that groups filters by category, making the filter panel scannable and self-explanatory.

### Layout
The accounts view changes from:
```
[search bar                    ] [chips...]
[table                                   ]
```
To:
```
[search bar                              ]
[sidebar  | table                        ]
```

The sidebar is **200px wide**, sticky within the scrollable main area (stays visible as the table scrolls). The table area uses the remaining width.

### Sidebar Structure
Three sections, always visible (no collapse toggle):

```
Filters          [Clear all]

PRIORITY
☐ HIGH
☐ MEDIUM
☐ LOW
☐ TBD

CLUSTER
☐ ACP
☐ Upstream
☐ Mid-Downstream
☐ ETS-OFS
☐ Mining

ENGAGEMENT STATUS
☐ Active
☐ Not Started
☐ Monitoring
☐ On Hold
☐ Restricted
```

- Section headers: small caps, muted color, same style as existing `chart-title` class
- Checkboxes: native `<input type="checkbox">` styled to match the app's green accent
- Active filters: section header gets a count badge (e.g. `PRIORITY · 2`) in green
- "Clear all" link: top-right of sidebar, only visible when ≥1 filter is active
- Values for Cluster and Engagement Status are derived dynamically from the loaded data (same as the old chip logic), so new clusters/statuses appear automatically

### Filter State
Filter state model is unchanged (`state.filters.clusters`, `state.filters.priorities`, `state.filters.engagementStatuses`). Only the rendering of controls changes — checkboxes replace chips. `applyFilters()` is called on every checkbox change, same as before.

### Search Bar
Moves to a full-width position above the sidebar+table row. Same input styling, same debounce logic.

### Removal
The `filter-chips` div, `.chip` and `.chip--active` CSS, and `buildFilterChips` / `syncChipState` functions are removed. The sidebar takes over their responsibility.

---

## Files Changed
| File | Change |
|---|---|
| `src/core.js` | Add `inferDomain(name)` and `companyInitials(name)` helpers; export both |
| `src/ui.js` | Add `logoHtml(company)` helper; update `buildTable` (add logo cell), `renderPanel` (add logo to header), remove chip functions, add `buildFilterSidebar` / `renderFilterSidebar`, update `applyFilters`, update `renderAccounts` |
| `src/styles.css` | Add `.logo-thumb`, `.logo-avatar`, `.filter-sidebar`, `.filter-section`, `.filter-checkbox`, `.filter-badge` styles; update `.accounts-layout` to use flexbox; remove `.chip`, `.chip--active`, `.filter-chips` |
| `src/app.template.html` | Update accounts view markup to wrap sidebar + table in `.accounts-layout` div |
| `build.js` | No change |

---

## Out of Scope
- Logo caching / localStorage persistence (Clearbit responses are already cached by the browser)
- Mobile/responsive layout (existing app is desktop-only)
- User-editable domain overrides
- Sidebar collapse/expand toggle
