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
LOCALITIES_PATH = Path(__file__).resolve().parent / "localities.json"

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

# Circle/area centroids (lat, lng) for the interactive map. Geocoded from
# Nominatim (OpenStreetMap) against the circle/area name + Hyderabad, India.
CIRCLE_COORDS = {
    "Bowenpally": [17.475537, 78.479228],
    "Alwal": [17.502229, 78.508858],
    "Jeedimetla": [17.519687, 78.446888],
    "Gajularamaram": [17.527176, 78.420008],
    "Nizampet": [17.497127, 78.376883],
    "Chintal": [17.502176, 78.440609],
    "Kompally": [17.535487, 78.509698],
    "Dundigal": [17.50621, 78.505108],
    "Kapra": [17.484636, 78.56101],
    "Uppal": [17.402509, 78.561256],
    "Boduppal": [17.398841, 78.536622],
    "Nacharam": [17.428494, 78.55281],
    "Malkajgiri": [17.451176, 78.5369],
    "Moula Ali": [17.46171, 78.55714],
    "Keesara": [17.481583, 78.592682],
    "Tarnaka": [17.428548, 78.537943],
    "Mettuguda": [17.435504, 78.519557],
    "Serilingampally": [17.466717, 78.340421],
    "Madhapur": [17.440892, 78.39163],
    "Miyapur": [17.498161, 78.356763],
    "Narsingi": [17.387417, 78.356624],
    "Patancheru": [17.528609, 78.267425],
    "Ameenpur": [17.523691, 78.33173],
    "Kukatpally": [17.493084, 78.405441],
    "Moosapet": [17.468531, 78.42067],
    "Allwyn Colony": [17.492035, 78.349953],
    "Ameerpet": [17.437501, 78.448251],
    "Jubilee Hills": [17.430836, 78.410288],
    "Yousufguda": [17.43875, 78.427987],
    "Borabanda": [17.459069, 78.407866],
    "Falaknuma": [17.33266, 78.475198],
    "Bahadurpura": [17.357067, 78.454542],
    "Chandrayangutta": [17.324696, 78.481356],
    "Yakutpura": [17.358628, 78.485805],
    "Jangammet": [17.336251, 78.474404],
    "Santoshnagar": [17.346719, 78.508195],
    "Charminar": [17.361602, 78.474642],
    "Malakpet": [17.373671, 78.499648],
    "Moosarambagh": [17.374353, 78.516084],
    "Goshamahal": [17.380576, 78.468846],
    "Karwan": [17.376013, 78.433189],
    "Golconda": [17.387329, 78.405734],
    "Attapur": [17.367224, 78.430728],
    "Rajendranagar": [17.334621, 78.40868],
    "Khairatabad": [17.412974, 78.461058],
    "Mehdipatnam": [17.394263, 78.434251],
    "Masab Tank": [17.402962, 78.450754],
    "Kavadiguda": [17.422702, 78.49177],
    "Musheerabad": [17.419142, 78.498573],
    "Amberpet": [17.386178, 78.511471],
    "Nagole": [17.377531, 78.560123],
    "Saroornagar": [17.361166, 78.538744],
    "L.B. Nagar": [17.349807, 78.547888],
    "Hayathnagar": [17.328115, 78.60454],
    "Sangareddy District": [17.528026, 78.267025],
    "Medchal-Malkajgiri Rural": [17.633993, 78.484315],
    "Shamshabad (Rangareddy)": [17.257207, 78.345104],
    "Adibatla (2026-merged)": [17.230899, 78.5559],
    "Badangpet (2026-merged)": [17.338347, 78.522213],
    "Jalpally (2026-merged)": [17.306154, 78.473859],
    "Ghatkesar Belt": [17.451084, 78.684302],
    "Yadadri-Bhuvanagiri District": [17.517279, 78.886338],
}

# Live-sheet Branch spellings not covered by the recovered GHMC locality
# aliases. Each maps to (route, parent_circle, note); the parent circle
# supplies the zone, centroids and dispatch guidance.
#
# Names are stored in the form cleanPlace() produces (title-cased, suffix
# words kept) so their slugs match dispatch Branch values after the trailing
# noise tokens ("Branch", "Road No 12", "2nd Floor", …) are stripped.
EXTRA_LOCALITIES = {
    "Gachibowli": (
        "R4",
        "Serilingampally",
        "ORR/Gachibowli offices; 11 AM-4 PM window; corporate receptions.",
    ),
    "Gacchibowli": (
        "R4",
        "Serilingampally",
        "Alt. spelling of Gachibowli (live sheet); ORR tech belt; 11 AM-4 PM.",
    ),
    "Kondapur": (
        "R4",
        "Serilingampally",
        "Gachibowli-HITEC ORR tech belt; deliver 11 AM-4 PM.",
    ),
    "Begumpet": (
        "R4",
        "Ameerpet",
        "Ameerpet/Begumpet commercial core; metered curbside drops; avoid 9-11 AM peak.",
    ),
    "SR Nagar": (
        "R4",
        "Yousufguda",
        "Residential (Srinagar Colony/SR Nagar); combine Ameerpet; mid-morning.",
    ),
    "AJC": (
        "R4",
        "Madhapur",
        "AJC building (2nd-4th floor branches) in HITEC City core; after 11 AM.",
    ),
    "Bala Nagar": (
        "R4",
        "Moosapet",
        "Balanagar industrial/residential belt (MSME-DI side); mid-day.",
    ),
    "MSME": (
        "R4",
        "Moosapet",
        "MSME Development Institute, Narsapur X Roads, Balanagar; office hours.",
    ),
    "East Marredpally": (
        "R3",
        "Kavadiguda",
        "Marredpally east wing, Secunderabad cantonment; avoid 5-7 PM rush.",
    ),
    "JNTU": (
        "R2",
        "Kukatpally",
        "JNTU Kukatpally campus belt; combine KPHB; 6-10 AM window.",
    ),
    "Kodad": (
        "R7",
        None,
        "Kodad, Suryapet district (~150 km); long-haul, plan toll; R7.",
    ),
    "MG Road": (
        "R3",
        "Kavadiguda",
        "Mahatma Gandhi Road, Secunderabad core; combine Patny; morning.",
    ),
    "Narayanguda": (
        "R4",
        "Khairatabad",
        "Central corridor (Himayatnagar side); 10 AM-5 PM windows.",
    ),
    "Puppalaguda": (
        "R4",
        "Narsingi",
        "ORR corridor near Manikonda; wide roads; morning loop.",
    ),
    "R P Road": (
        "R3",
        "Kavadiguda",
        "RP Road, Secunderabad (Paradise/Patny side); LCV/2W.",
    ),
    "Sardar Patel Road": (
        "R3",
        "Kavadiguda",
        "SP Road, Secunderabad; dense commercial; early morning.",
    ),
    "Shivam Road": (
        "R3",
        "Saroornagar",
        "Shivam Road, Dilsukhnagar; combine Kothapet; after 11 AM.",
    ),
    "Suryodaya Chambers": (
        "R4",
        "Ameerpet",
        "Suryodaya building, Begumpet belt; metered curbside; mid-day.",
    ),
    "SVM Grand Medipally": (
        "R3",
        "Boduppal",
        "Branch master (C0178) - Medipally; combine Boduppal loop.",
    ),
    "Allu Cinemas - Dolby Cinema & Cineplex": (
        "R4",
        "Narsingi",
        "ALLU Cinemas, Kokapet/Narsingi (500075); morning loop.",
    ),
}


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


def locality_specs():
    """Yield (route, circle, name, note) locality aliases.

    Sources, in merge order (earlier wins on duplicate names):
      1. localities.json groups   — circle memberships (54 groups)
      2. localities.json localities — standalone alias records (23)
      3. EXTRA_LOCALITIES           — live-sheet spellings missing above
    """
    specs = []
    try:
        loc_data = json.loads(LOCALITIES_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        loc_data = {}
    for key, (note, names) in loc_data.get("groups", {}).items():
        route, circle = key.split("|", 1)
        for name in names:
            specs.append((route, circle, name, note))
    for name, (route, circle, note) in loc_data.get("localities", {}).items():
        specs.append((route, circle, name, note))
    for name, (route, circle, note) in EXTRA_LOCALITIES.items():
        specs.append((route, circle, name, note))
    return specs


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
                    "lat": CIRCLE_COORDS.get(val, [None, None])[0],
                    "lng": CIRCLE_COORDS.get(val, [None, None])[1],
                }
            )

    index = {(r["route"], r["name"].lower()): r for r in records}
    for route, circle, name, note in locality_specs():
        parent = index.get((route, circle.lower())) if circle else None
        key = (route, name.lower())
        if key in seen:
            continue
        if not parent:
            if route != "R7":
                continue
            seen.add(key)
            records.append(
                {
                    "name": name,
                    "route": route,
                    "zone": R7_ZONE,
                    "kind": "locality",
                    "note": note,
                    "lat": None,
                    "lng": None,
                }
            )
            continue
        seen.add(key)
        records.append(
            {
                "name": name,
                "route": route,
                "circle": parent["name"],
                "zone": parent["zone"],
                "kind": "locality",
                "note": note,
                "lat": parent["lat"],
                "lng": parent["lng"],
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
