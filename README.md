# SkyLimit Outwards Dashboard

Enterprise **outward logistics / dispatch intelligence** — a zero-backend web app that tracks outward order volume, open bottlenecks, dispatch velocity and fulfilment across customers, branches and locations, visualised live from Google Sheets.

🔗 **Live:** [inakm.github.io/outward-dashboard](https://inakm.github.io/outward-dashboard/)

## Features

### Data ingestion
- **Drag & drop** Excel/CSV files (`.xlsx`, `.xls`, `.csv`) — parsed entirely in the browser with SheetJS, never uploaded
- **Live Google Sheet sync** — pulls the public dispatch sheet via CSV export with a gviz fallback
- **Data-cleaning layer** — header aliasing, place-name typo correction (Hyderabad area), priority/ack normalisation
- **IndexedDB persistence** — reload-safe local cache of records, customer master and branch list

### Analytics
- **KPI grid** — total outward volume, open bottlenecks (P1 not yet acknowledged), dispatch velocity, fulfilment rate
- **Volume over time** — line chart of orders vs dispatched, grouped by **day / week / month**
- **Volume by dimension** — top-N bar chart switchable between **Customer / Location / Branch**
- **Acknowledgement state** — Done · In Transit · Pending doughnut

### Operations
- **Filters** — priority (P1–P4), ack state, **route (R1–R5)**, **order-date range**, free-text search
- **Sortable matrix table** — customer · branch · location with all KPIs
- **Backlog aging** — age buckets (0–3d → 31+d) and the oldest open P1 bottlenecks
- **Routes overview** — per-route volume, open P1, ack breakdown and fulfilment bar; click a route to filter
- **Customer directory** — master list with branches and outward activity, expandable branch drawer

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
index.html     — markup, SEO/OG meta, structured data
style.css      — Linear-inspired light design system (CSS variables)
script.js      — ingestion, cleaning, persistence, analytics, charts, events
DESIGN.md      — original design-system reference
logo.png       — brand mark
```

### Data model

Dispatch records are normalised to a canonical shape:

| Field | Source column | Notes |
|---|---|---|
| `orderDate` | Order Date | `dd/mm/yyyy` etc. |
| `customer` | Customer Name | trimmed / line-break cleaned |
| `branch` | Branch | place-name typo-corrected |
| `location` | Location | place-name typo-corrected |
| `route` | ROUTE | delivery route (R1–R5, filterable) |
| `priority` | Priority | normalised to P1–P4 / — |
| `dispatchDate` | Dispatch date | |
| `ack` | Ack | normalised to Done / In Transit / Pending |

Pure analysis helpers (`computeKpis`, `buildBarData`, `buildTrendData`, `buildAgingData`, `auditData`, `filteredRows`, …) are exposed on `window.DashboardCore` for debugging and testing.

## Deployment

This repo is deployed to GitHub Pages from the `main` branch. Push and it publishes automatically.

## Roadmap ideas

- Export filtered view (CSV / Excel) + print stylesheet
- Light / dark theme toggle
- PWA / offline install (service worker)

## License

© SkyLimit. Internal tool.
