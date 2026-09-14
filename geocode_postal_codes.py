"""
Geocode Worker's Postal Codes
------------------------------
Reads an input spreadsheet, geocodes every postal code found in the
`Worker's Postal` column using the Google Maps Geocoding API, and writes
out a copy of the spreadsheet with two new columns added: `Latitude` and
`Longitude`.

Requirements:
    pip install pandas openpyxl requests

Usage:
    python geocode_postal_codes.py path/to/input.xlsx [path/to/output.xlsx]

    If no output path is given, the result is written next to the input
    file with "_geocoded" appended to the filename.

API key:
Supply your Google Maps API key one of two ways (checked in this order):
    1. Set a GOOGLE_MAPS_API_KEY environment variable before running.
    2. Paste it into GOOGLE_MAPS_API_KEY below.
"""

import argparse
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

import pandas as pd

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

GOOGLE_MAPS_API_KEY = ""  # optionally paste your key here

POSTAL_COLUMN = "Worker's Postal"
LAT_COLUMN = "Latitude"
LONG_COLUMN = "Longitude"

GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
REQUEST_TIMEOUT_SECONDS = 20
RETRY_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = 2


def _get_api_key():
    if GOOGLE_MAPS_API_KEY:
        return GOOGLE_MAPS_API_KEY
    env_key = os.environ.get("GOOGLE_MAPS_API_KEY")
    if env_key:
        return env_key
    raise RuntimeError(
        "No Google Maps API key found. Set a GOOGLE_MAPS_API_KEY environment "
        "variable, or paste your key into GOOGLE_MAPS_API_KEY at the top of "
        "geocode_postal_codes.py."
    )


def _normalize_column_name(name):
    text = str(name).strip().lower()
    text = text.replace("’", "'").replace("‘", "'")
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text


def _find_postal_column(columns):
    wanted = _normalize_column_name(POSTAL_COLUMN)
    for col in columns:
        if _normalize_column_name(col) == wanted:
            return col
    raise ValueError(
        f"Could not find a '{POSTAL_COLUMN}' column. Found columns: {list(columns)}"
    )


def geocode_postal_code(postal_code, api_key):
    """Geocode a single postal code via the Google Maps Geocoding API,
    biased to Canada. Returns (lat, lng), or (None, None) if it can't be
    resolved. Retries transient network errors a few times."""
    params = {"address": postal_code, "region": "ca", "key": api_key}
    url = GOOGLE_GEOCODE_URL + "?" + urllib.parse.urlencode(params)

    last_error = None
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            with urllib.request.urlopen(url, timeout=REQUEST_TIMEOUT_SECONDS) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            break
        except (urllib.error.URLError, TimeoutError) as exc:
            last_error = exc
            if attempt < RETRY_ATTEMPTS:
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
    else:
        print(f"  Warning: network error geocoding '{postal_code}': {last_error}")
        return None, None

    status = payload.get("status")
    if status == "OK":
        location = payload["results"][0]["geometry"]["location"]
        return location["lat"], location["lng"]

    if status == "ZERO_RESULTS":
        print(f"  Warning: no geocoding result for postal code '{postal_code}'.")
        return None, None

    if status in ("OVER_QUERY_LIMIT", "OVER_DAILY_LIMIT"):
        raise RuntimeError(
            f"Google Geocoding API rate/quota limit hit (status={status!r}) while "
            f"geocoding '{postal_code}'. Check your API quota/billing and retry."
        )

    raise RuntimeError(
        f"Google Geocoding API returned status={status!r} for postal code "
        f"'{postal_code}' ({payload.get('error_message', 'no further detail')}). "
        f"Check that the Geocoding API is enabled and billing is set up for "
        f"this API key in Google Cloud Console."
    )


def geocode_spreadsheet(input_path, output_path, api_key):
    df = pd.read_excel(input_path)
    postal_column = _find_postal_column(df.columns)

    lats = []
    longs = []
    cache = {}

    total = len(df)
    for idx, raw_postal in enumerate(df[postal_column], start=1):
        postal_code = str(raw_postal).strip() if pd.notna(raw_postal) else ""
        if not postal_code:
            lats.append(None)
            longs.append(None)
            continue

        key = postal_code.upper()
        if key in cache:
            lat, lng = cache[key]
        else:
            print(f"[{idx}/{total}] Geocoding '{postal_code}'...")
            lat, lng = geocode_postal_code(postal_code, api_key)
            cache[key] = (lat, lng)

        lats.append(lat)
        longs.append(lng)

    df[LAT_COLUMN] = lats
    df[LONG_COLUMN] = longs

    df.to_excel(output_path, index=False)
    print(f"\nDone. Wrote {len(df)} row(s) to {output_path}")

    missing = df[LAT_COLUMN].isna().sum()
    if missing:
        print(f"Warning: {missing} row(s) could not be geocoded (left blank).")


def _default_output_path(input_path):
    root, ext = os.path.splitext(input_path)
    return f"{root}_geocoded{ext}"


def main():
    parser = argparse.ArgumentParser(
        description="Geocode postal codes in the Worker's Postal column of a spreadsheet."
    )
    parser.add_argument("input", help="Path to the input .xlsx spreadsheet")
    parser.add_argument(
        "output",
        nargs="?",
        default=None,
        help="Path to write the output .xlsx (default: <input>_geocoded.xlsx)",
    )
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: input file not found: {args.input}")
        sys.exit(1)

    output_path = args.output or _default_output_path(args.input)
    api_key = _get_api_key()

    geocode_spreadsheet(args.input, output_path, api_key)


if __name__ == "__main__":
    main()
