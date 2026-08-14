# SkyLimit Outwards Dashboard

Enterprise **outward logistics / dispatch intelligence** — a zero-backend web app that tracks outward order volume, open bottlenecks, dispatch velocity and fulfilment across customers, branches and locations, visualised live from Google Sheets.

🔗 **Live:** [inakm.github.io/outward-dashboard](https://inakm.github.io/outward-dashboard/)

## Features

### Data ingestion
- **Drag & drop** Excel/CSV files (`.xlsx`, `.xls`, `.csv`) — parsed entirely in the browser with SheetJS, never uploaded
- **Live Google Sheet sync** — pulls the public dispatch sheet via CSV export with a gviz fallback, and loads the customer master from the same workbook's **Location pull** sheet (Primary Zone → R1–R7 routing)
- **Auto-routing from `routes-planning.xlsx`** — the R1–R7 plan (R1–R4 GHMC circles, R5 Outside Hyderabad, R6 Madhapur, R7 Outside Telangana) is the routing authority; orders are auto-routed by matching their **Branch** column, supplemented by the GHMC ward/locality map for fine-grained Hyderabad areas. A small **branch override** (`Madhapur→R6`, Andhra cities `Srikakulam/Kakinada/Lalapet Guntur→R7`, Telangana-outside-city `Choutuppal/Tukkuguda/Ramoji Film City→R5`) runs first so the old GHMC scheme (Madhapur→R4) is superseded. Orders whose branch isn't in the plan fall back to the customer master's **Primary Zone** (tolerant match: exact → case-insensitive → whole-phrase → first-token).
- **Data-cleaning layer** — header aliasing, place-name typo correction (Hyderabad area), priority/ack normalisation
- **IndexedDB persistence** — reload-safe local cache of records, customer master and branch list

### Analytics
- **KPI grid** — total outward volume, open bottlenecks (P1 not yet acknowledged), dispatch velocity, fulfilment rate
- **Volume over time** — line chart of orders vs dispatched, grouped by **day / week / month**
- **Volume by dimension** — top-N bar chart switchable between **Customer / Location / Branch**
- **Acknowledgement state** — Done · In Transit · Pending doughnut

### Operations
- **Filters** — priority (P1–P4), ack state, **route (R1–R7)**, **order status**, **order-date range**, free-text search
- **Sortable matrix table** — customer · branch · location with all KPIs
- **Backlog aging** — age buckets (0–3d → 31+d) and the oldest open P1 bottlenecks
- **Routes overview** — per-route volume, open P1, ack breakdown and fulfilment bar; click a route to filter
- **Customer directory** — master list (code, city, **primary zone**, contact, GST) with branches and outward activity, expandable branch drawer

### Reliability & data quality
- **Auto-refresh** — optional 5-minute poll of the Google Sheet (persisted in `localStorage`)
- **Last-updated stamp** — relative freshness indicator in the nav
- **Data quality report** — flags missing dates, missing fields, unrecognised priorities, dispatch-before-order and future-dated orders

## Tech stack

| Concern | Choice |
|---|---|
| Core | Vanilla HTML / CSS / JS (ES5-compatible, no build step) |
| Charts | [Chart.js](https://www.chartjs.org/) 4.4 |
| Spreadsheets | [SheetJS](https://sheetjs.com/) (`xlsx`) |
| Icons | [Lucide](https://lucide.dev/) |
| Storage | IndexedDB |
| Fonts | Inter + JetBrains Mono (Google Fonts) |
| Data source | Google Sheets public CSV export |

## Getting started

The app is fully static — no build or server required.

```bash
# serve locally (any static file server works)
python -m http.server 8000
# then open http://localhost:8000
```

Open the page, then either:
1. **Drag & drop** a dispatch/customer/branch workbook, or
2. Click **Pull live from Google Sheet** to sync the configured sheet.

Your data is restored automatically from the local cache on reload. Everything runs in the browser — no data leaves your machine except the Google Sheet sync request.

> **Note:** Google Sheet sync uses `fetch` and won't work from `file://`. Use a local server or GitHub Pages.

## Project structure

```
index.html              — markup, SEO/OG meta, JSON-LD structured data
assets/css/style.css    — Linear-inspired light design system (CSS variables)
assets/js/script.js     — ingestion, cleaning, persistence, analytics, charts, events
assets/js/route-plan.js — R1–R7 route plan (window.ROUTE_PLAN, includes map coordinates)
assets/js/fleet.js      — fleet config (window.GHMC_FLEET)
assets/xcl/             — Excel workbooks (routing authority, customer mappings)
assets/json/            — JSON + manifest (vehicles.json fleet config, route_plan.json, localities.json, manifest.webmanifest)
assets/images/          — brand mark, iOS/Android icons, share card
llms.txt / llms-full.txt — LLM-friendly site description and full docs
robots.txt              — crawl rules (hides raw data files)
sitemap.xml             — XML sitemap (hosted SPA root)
404.html                — branded not-found page (noindex)
_config.yml             — GitHub Pages / Jekyll build config
DESIGN.md               — original design-system reference
```

### Data model

Dispatch records are normalised to a canonical shape:

| Field | Source column | Notes |
|---|---|---|
| `orderDate` | Order Date | `dd/mm/yyyy` etc. |
| `invoiceNo` | Invoice No. | invoice reference, searchable via the matrix search box |
| `customer` | Customer Name | trimmed / line-break cleaned |
| `customerCode` | _derived_ | resolved customer Code — tolerant name match against the customer master, falling back to the branch workbooks (branch name → owning customer Code); used to align orders with a customer's branches and to route via the master's Primary Zone |
| `branch` | Branch | place-name typo-corrected |
| `location` | Location | place-name typo-corrected |
| `route` | ROUTE | delivery route (R1–R7, filterable) — auto-assigned by matching Branch against the `routes-planning.xlsx` R1–R7 plan; falls back to the customer master's Primary Zone (North-East→R1, North-West→R2, South-East→R3, South-West→R4, Outside Hyderabad→R5, Madhapur→R6, Outside Telangana→R7) when the branch isn't in the plan |
| `priority` | Priority | normalised to P1–P4 / — |
| `dispatchDate` | Dispatch date | |
| `ack` | Ack | normalised to Done / In Transit / Pending |
| `status` | Status | normalised operational status (Packed / Pulled / …), optional |

Pure analysis helpers (`computeKpis`, `buildBarData`, `buildTrendData`, `buildAgingData`, `auditData`, `filteredRows`, …) are exposed on `window.DashboardCore` for debugging and testing.

## Deployment

This repo is deployed to GitHub Pages from the `main` branch. Push and it publishes automatically.

`_config.yml` keeps the Jekyll build minimal and explicitly publishes the `.xlsx` / `.json` / `.webmanifest` assets the app fetches at runtime. It is safe to extend with core Jekyll keys, but avoid plugins — the Pages builder only allows Jekyll's built-in allowlist.

## SEO & discoverability

The dashboard is a JS-rendered single-page app, so the SEO weight lives in the document shell:

- **Meta & OG/Twitter** — title, description, Open Graph and Twitter Card tags with a 1200×630 share image (`og-image.png`).
- **JSON-LD** — `WebApplication` structured data (features, versioning, free offer, publisher) in `index.html`.
- **Manifest** — `assets/json/manifest.webmanifest` for installability / home-screen metadata with square icons (`icon-192.png`, `icon-512.png`).
- **Crawl control** — `robots.txt` blocks indexing of raw workbooks (`*.xlsx`) and the customer master; `sitemap.xml` lists the single hosted URL with its image.
- **Fallbacks** — branded `404.html` (noindex) and a `<noscript>` block with crawlable feature content.
- **Performance** — preconnect to fonts/CDNs and `preload` of `style.css`.

> Note: live data is fetched client-side, so crawlers see the empty dashboard shell — the meta tags and structured data carry the indexable description.

## Roadmap ideas

- Export filtered view (CSV / Excel) + print stylesheet
- Light / dark theme toggle
- PWA / offline install (service worker)

## License

© SkyLimit. Internal tool.
