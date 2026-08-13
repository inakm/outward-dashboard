#!/usr/bin/env python3
"""Bundles vehicles.json into fleet.js (window.GHMC_FLEET) so the dispatch
planner keeps working when the page is opened straight from the file system
(fetch() to local JSON is CORS-blocked). Run from the repo root or anywhere:

    python ghmc-routing/generate_fleet.py
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "assets" / "json" / "vehicles.json"
OUT = ROOT / "assets" / "js" / "fleet.js"


def main() -> None:
    data = json.loads(SRC.read_text(encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False, indent=2)
    out = (
        "/* Bundled fleet config for the dispatch planner.\n"
        "   Source of truth: vehicles.json (edit that, then regenerate). */\n"
        "window.GHMC_FLEET = " + payload + ";\n"
    )
    OUT.write_text(out, encoding="utf-8")
    print(f"Wrote {OUT.name} ({len(data.get('fleet', []))} vehicles)")


if __name__ == "__main__":
    main()
