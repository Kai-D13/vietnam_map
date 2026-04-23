"""
Download Vietnamese district-level GeoJSON boundaries from the dvhcvn repo
(https://github.com/daohoangson/dvhcvn) and merge them into data/districts.geojson.

Run once:
  python download_districts.py

The file data/districts.geojson is served by Flask at /api/boundaries/districts.
Data source: data/gis/{id}.json files — each has level2s (districts) with coordinates.
"""
import sys
import json
import os
import time

import requests

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUTPUT_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "districts.geojson")
API_URL     = "https://api.github.com/repos/daohoangson/dvhcvn/contents/data/gis"
RAW_URL     = "https://raw.githubusercontent.com/daohoangson/dvhcvn/master/data/gis/{name}"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Get list of province files
print("Fetching file list from GitHub...")
resp = requests.get(API_URL, timeout=15)
resp.raise_for_status()
files = [item["name"] for item in resp.json()
         if item["name"].endswith(".json") and item["name"] != "level1s_bbox.json"]
print(f"Found {len(files)} province files")

features = []
for fname in sorted(files):
    url = RAW_URL.format(name=fname)
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        province = r.json()
        province_name = province.get("name", "")
        level2s = province.get("level2s", [])
        for district in level2s:
            feat = {
                "type": "Feature",
                "geometry": {
                    "type": district.get("type", "MultiPolygon"),
                    "coordinates": district["coordinates"],
                },
                "properties": {
                    "name":         district.get("name", ""),
                    "level2_id":    district.get("level2_id", ""),
                    "province":     province_name,
                    "level1_id":    province.get("level1_id", ""),
                },
            }
            features.append(feat)
        print(f"  {fname}: {len(level2s)} districts (total {len(features)})")
    except Exception as e:
        print(f"  {fname}: ERROR - {e}")
    time.sleep(0.05)

merged = {"type": "FeatureCollection", "features": features}
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(merged, f, ensure_ascii=False)

print(f"\nSaved {len(features)} district features -> {OUTPUT_FILE}")
