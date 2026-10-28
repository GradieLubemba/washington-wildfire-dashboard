#!/usr/bin/env python3
"""
Washington State Wildfire Preparedness Dashboard - data pre-processor.

Reads the real Washington DNR Fire Statistics CSV and produces a compact
JSON file (data/dashboard_data.json) consumed by index.html.

Audience: Washington State emergency management & wildfire planning officials.

Rules applied (per requirements):
  * Excludes false alarms from ALL wildfire statistics.
    Event classes treated as false alarms: "False Alarm", "FA".
  * Missing / blank fire causes are reported as "Unknown".
  * Records without a usable discovery year (>= MIN_YEAR) or without valid
    Washington coordinates are reported in meta and excluded from the
    year/month/map analyses respectively (still counted where applicable).
  * Only real CSV data is used. No mock or synthetic data is generated.

Usage:
    python3 build_data.py
"""

import csv
import json
import os
from collections import defaultdict

CSV_FILE = "DNR_Fire_Statistics_2008_-_Present.csv"
OUT_DIR = "data"
OUT_FILE = os.path.join(OUT_DIR, "dashboard_data.json")

# Event classes that represent false alarms (excluded from all stats).
FALSE_ALARM_CLASSES = {"False Alarm", "FA"}

# Dataset scope: discovery year 2008 onward.
MIN_YEAR = 2008

# Rough Washington State bounding box used to validate coordinates.
WA_BBOX = {"lat_min": 45.0, "lat_max": 49.5, "lon_min": -125.0, "lon_max": -116.0}

MONTH_ORDER = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER",
]
MONTH_INDEX = {m: i + 1 for i, m in enumerate(MONTH_ORDER)}


def clean(value):
    return (value or "").strip()


def to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def valid_wa(lat, lon):
    return (
        lat is not None and lon is not None
        and WA_BBOX["lat_min"] <= lat <= WA_BBOX["lat_max"]
        and WA_BBOX["lon_min"] <= lon <= WA_BBOX["lon_max"]
    )


def main():
    if not os.path.exists(CSV_FILE):
        raise SystemExit(f"CSV not found: {CSV_FILE}")
    os.makedirs(OUT_DIR, exist_ok=True)

    total_rows = 0
    false_alarms_excluded = 0
    no_valid_year = 0
    no_valid_coords = 0

    records = []          # compact per-record rows for client-side filtering/map
    counties = set()
    regions = set()
    causes = set()
    years = set()

    with open(CSV_FILE, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            total_rows += 1

            event_class = clean(row.get("FIREEVNT_CLASS_LABEL_NM"))
            if event_class in FALSE_ALARM_CLASSES:
                false_alarms_excluded += 1
                continue

            county = clean(row.get("COUNTY_LABEL_NM")) or "Unknown"
            region = clean(row.get("REGION_NAME")) or "Unknown"
            cause = clean(row.get("FIREGCAUSE_LABEL_NM")) or "Unknown"
            acres = to_float(row.get("ACRES_BURNED")) or 0.0

            discovery = clean(row.get("DSCVR_DT"))
            year = None
            if len(discovery) >= 4 and discovery[:4].isdigit():
                y = int(discovery[:4])
                if y >= MIN_YEAR:
                    year = y
            if year is None:
                no_valid_year += 1

            month_name = clean(row.get("DSCVR_MONTH")).upper()
            month = MONTH_INDEX.get(month_name)  # 1..12 or None

            lat = to_float(row.get("LAT_COORD"))
            lon = to_float(row.get("LON_COORD"))
            has_coords = valid_wa(lat, lon)
            if not has_coords:
                no_valid_coords += 1

            counties.add(county)
            regions.add(region)
            causes.add(cause)
            if year is not None:
                years.add(year)

            # Compact record. Round coords to shrink payload.
            records.append({
                "c": county,
                "r": region,
                "u": cause,
                "y": year,                       # may be None
                "m": month,                      # may be None (1..12)
                "a": round(acres, 2),
                "lat": round(lat, 5) if has_coords else None,
                "lon": round(lon, 5) if has_coords else None,
            })

    records_used = len(records)
    total_acres = round(sum(r["a"] for r in records), 2)

    output = {
        "meta": {
            "source_file": CSV_FILE,
            "generated_scope": f"discovery year >= {MIN_YEAR}",
            "total_rows_in_csv": total_rows,
            "false_alarm_classes": sorted(FALSE_ALARM_CLASSES),
            "false_alarms_excluded": false_alarms_excluded,
            "records_used": records_used,
            "total_acres_burned": total_acres,
            "records_without_valid_year": no_valid_year,
            "records_without_valid_coords": no_valid_coords,
            "wa_bbox": WA_BBOX,
            "limitations": [
                "False alarms (classes 'False Alarm','FA') are excluded from all statistics.",
                "Blank fire causes are shown as 'Unknown'; 'Undetermine' is a separate DNR-recorded category and is kept distinct.",
                "Records missing a valid year are excluded from year/month trends only.",
                "Records missing valid Washington coordinates are omitted from the map only.",
                "ACRES_BURNED is highly right-skewed; a few large fires dominate totals.",
                "This dashboard shows HISTORICAL patterns only. It does not predict future fires.",
            ],
        },
        "filters": {
            "counties": sorted(counties),
            "regions": sorted(regions),
            "causes": sorted(causes),
            "years": sorted(years),
            "months": MONTH_ORDER,
        },
        "records": records,
    }

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, separators=(",", ":"))

    size_mb = os.path.getsize(OUT_FILE) / (1024 * 1024)
    print(f"Wrote {OUT_FILE} ({size_mb:.2f} MB)")
    print(f"  Total rows in CSV:            {total_rows}")
    print(f"  False alarms excluded:        {false_alarms_excluded}")
    print(f"  Records used:                 {records_used}")
    print(f"  Total acres burned:           {total_acres}")
    print(f"  Counties / Regions / Causes:  {len(counties)} / {len(regions)} / {len(causes)}")
    print(f"  Missing valid year:           {no_valid_year}")
    print(f"  Missing valid coordinates:    {no_valid_coords}")


if __name__ == "__main__":
    main()
