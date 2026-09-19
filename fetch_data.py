"""Fetch historical launches from Launch Library 2 and save a flat CSV.

Usage:
    python fetch_data.py                 # download every page not saved yet
    python fetch_data.py --max-pages 3   # only fetch 3 new pages (good first test)

The free API tier is rate limited (roughly 15 requests per hour), so each raw
page is cached in data/raw/. If the limit is hit the script stops cleanly;
just re-run it later and it picks up where it left off. data/launches.csv is
rebuilt from whatever pages exist every time, so you can start analysing
with partial data.
"""

import argparse
import json
from pathlib import Path

import pandas as pd
import requests

BASE_URL = "https://ll.thespacedevs.com/2.2.0/launch/previous/"
PAGE_SIZE = 100
RAW_DIR = Path("data/raw")
OUT_CSV = Path("data/launches.csv")
HEADERS = {"User-Agent": "launch-analysis-student-project"}


def fetch_page(offset):
    """Return one page of results as a dict, or None if rate limited."""
    resp = requests.get(
        BASE_URL,
        params={"limit": PAGE_SIZE, "offset": offset},
        headers=HEADERS,
        timeout=30,
    )
    if resp.status_code == 429:
        return None
    resp.raise_for_status()
    return resp.json()


def download(max_pages=None):
    """Download pages one by one, skipping any already saved on disk."""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    offset, total, fetched = 0, None, 0

    while total is None or offset < total:
        path = RAW_DIR / f"page_{offset:06d}.json"

        if path.exists():  # already have it, skip
            if total is None:
                total = json.loads(path.read_text(encoding="utf-8"))["count"]
            offset += PAGE_SIZE
            continue

        if max_pages is not None and fetched >= max_pages:
            print(f"Reached --max-pages ({max_pages}). Re-run to fetch more.")
            break

        data = fetch_page(offset)
        if data is None:
            print("Rate limited (HTTP 429). Saved pages are kept. Re-run later to resume.")
            break

        path.write_text(json.dumps(data), encoding="utf-8")
        total = data["count"]
        fetched += 1
        print(f"Saved offset {offset} (total launches in API: {total})")
        offset += PAGE_SIZE


def name_of(value):
    """Some fields are plain strings, some are {'name': ...} objects."""
    return value.get("name") if isinstance(value, dict) else value


def flatten(launch):
    """Turn one nested launch record into a flat row."""
    provider = launch.get("launch_service_provider") or {}
    config = (launch.get("rocket") or {}).get("configuration") or {}
    mission = launch.get("mission") or {}
    orbit = mission.get("orbit") or {}
    pad = launch.get("pad") or {}
    location = pad.get("location") or {}
    status = launch.get("status") or {}

    return {
        "id": launch.get("id"),
        "name": launch.get("name"),
        "net": launch.get("net"),
        "status": status.get("name"),
        "fail_reason": launch.get("failreason"),
        "provider": provider.get("name"),
        "provider_type": name_of(provider.get("type")),
        "rocket": config.get("name"),
        "rocket_family": config.get("family"),
        "mission_type": name_of(mission.get("type")),
        "orbit": orbit.get("name"),
        "orbit_abbrev": orbit.get("abbrev"),
        "pad": pad.get("name"),
        "location": location.get("name"),
        "country_code": location.get("country_code"),
    }


def build_csv():
    """Combine every saved page into one CSV."""
    rows = []
    for path in sorted(RAW_DIR.glob("page_*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        rows.extend(flatten(launch) for launch in data["results"])

    if not rows:
        print("No pages saved yet, nothing to build.")
        return

    df = pd.DataFrame(rows).drop_duplicates(subset="id")
    df["net"] = pd.to_datetime(df["net"], utc=True, errors="coerce")
    df = df.sort_values("net")
    df.to_csv(OUT_CSV, index=False)
    print(f"Wrote {len(df)} launches to {OUT_CSV}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--max-pages", type=int, default=None,
                        help="stop after fetching this many new pages")
    args = parser.parse_args()

    download(args.max_pages)
    build_csv()