#!/usr/bin/env python3
"""Build the dashboard route-plan data from routes-planning.xlsx.

routes-planning.xlsx is the authoritative R1-R7 routing plan:
  - Routes sheet, planning section (rows 1..~126):
      R1 = North-East circles (col B), R2 = North-West (col C),
      R3 = South-East (col D), R4 = South-West (col E),
      R5 = outside-GHMC fringe areas (col F), R7 = non-Hyderabad cities (col G).
  - The generated route_plan.json + route-plan.js are what the dashboard loads
    to auto-route dispatch orders by their Branch column into R1..R7.

Workflow:
  1. Edit routes-planning.xlsx (the planning source of truth).
  2. Run:  python ghmc-routing/build_route_plan.py
  3. Commit the regenerated JSON + JS bundle.
"""
import json
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = REPO_ROOT / "routes-planning.xlsx"
OUT_DIR = Path(__file__).resolve().parent / "output"
JSON_PATH = OUT_DIR / "route_plan.json"
JS_PATH = REPO_ROOT / "route-plan.js"

# (route_code, column_index, kind) — column index is the 0-based cell index
# into the Routes sheet row tuple: A=0, B=1, C=2, D=3, E=4, F=5, G=6
ROUTE_COLUMNS = [
    ("R1", 1, "circle"),
    ("R2", 2, "circle"),
    ("R3", 3, "circle"),
    ("R4", 4, "circle"),
    ("R5", 5, "area"),
    ("R7", 6, "city"),
]

R7_TITLE_PREFIX = "R7"
R7_ZONE = "Other (non-Hyderabad)"
R5_ZONE_FALLBACK = "Outside GHMC (non-municipal)"


def clean(v):
    return str(v).strip() if v is not None else ""


def planning_rows(rows):
    """Yield (index, row) for the planning section (stops at the customer ledger)."""
    for i, row in enumerate(rows):
        if i == 0:
            continue
        if clean(row[3]).strip().lower() == "route":
            break
        yield i, row


def build_records():
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    ws = wb["Routes"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    records = []
    seen = set()
    current_zone = {code: "" for code, _, _ in ROUTE_COLUMNS}

    for _i, row in planning_rows(rows):
        s_no = clean(row[0])
        for code, col, kind in ROUTE_COLUMNS:
            val = clean(row[col]) if col < len(row) else ""
            if not val:
                continue
            if code == "R7":
                if val.startswith(R7_TITLE_PREFIX) or "cities served" in val:
                    continue
                zone = R7_ZONE
            else:
                if not s_no:
                    current_zone[code] = val
                    continue
                zone = current_zone[code] or (
                    R5_ZONE_FALLBACK if code == "R5" else ""
                )
            key = (code, val.lower())
            if key in seen:
                continue
            seen.add(key)
            records.append(
                {
                    "name": val,
                    "route": code,
                    "zone": zone,
                    "kind": kind,
                }
            )

    order = {"R1": 0, "R2": 1, "R3": 2, "R4": 3, "R5": 4, "R7": 5}
    records.sort(key=lambda r: (order[r["route"]], r["name"].lower()))
    return records


def main():
    records = build_records()
    counts = {}
    for r in records:
        counts[r["route"]] = counts.get(r["route"], 0) + 1

    payload = {
        "meta": {
            "title": "SkyLimit Outward Delivery Routes - R1-R7 Plan",
            "source": str(XLSX_PATH.name),
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "row_count": len(records),
            "counts": {k: counts.get(k, 0) for k in ("R1", "R2", "R3", "R4", "R5", "R7")},
        },
        "records": records,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    JSON_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    js = (
        "/* Auto-generated from routes-planning.xlsx by\n"
        "   ghmc-routing/build_route_plan.py. Bundled as a plain script so the\n"
        "   R1-R7 route plan works even when the app is opened from the file\n"
        "   system (where fetch() to local JSON is blocked by CORS).\n"
        "   Regenerate with:  python ghmc-routing/build_route_plan.py\n"
        "*/\n"
        "window.ROUTE_PLAN = "
        + json.dumps(payload, ensure_ascii=False)
        + ";\n"
    )
    JS_PATH.write_text(js, encoding="utf-8")

    print(f"Wrote {JSON_PATH}  ({len(records)} rows)")
    print(f"Wrote {JS_PATH}")
    print("Per-route counts: " + ", ".join(f"{k}={counts.get(k, 0)}" for k in ("R1", "R2", "R3", "R4", "R5", "R7")))


if __name__ == "__main__":
    main()
