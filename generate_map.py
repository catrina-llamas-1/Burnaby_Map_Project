"""
Postal Code Map Generator
==========================

Plots worker postal codes (British Columbia & Alberta) on an interactive
Leaflet/Folium map, with two independent-but-combinable checkbox filters:

  1) Region     -> classified from a keyword column into ClaimsPro / SCM / Pario
  2) Work Address - Line 1 -> one of three known office addresses

ClaimsPro and SCM share one color. Pario gets its own color. Both filters
can be used together (AND logic) or on their own.

Each worker is also connected to their classified Work Address by a
straight line. Drive time/distance is computed via the Google Maps
Distance Matrix API for BOTH directions at once -- 7:30 AM (Worker's
Postal -> Work Address) and 4:30 PM (Work Address -> Worker's Postal,
reverse direction) -- and both are baked into the exported map.html, which
has a live toggle to switch between them (no build-time choice, no API
calls from the viewer's browser). A left-hand sidebar lists every worker
with an individual include/exclude checkbox, and a KPI panel recomputes
live, for whichever direction is toggled, from whichever workers are
currently included: base size, average/median drive time, % under
15/20/30 minutes, average distance, and the closest/farthest postal code
by drive time.

Designed to run top-to-bottom as a single Google Colab cell (or as a normal
Python script). Everything you're likely to want to tweak lives in the
CONFIG block below.

Input:  an .xlsx spreadsheet with (at least) these columns:
          - Worker's Postal              (postal code, shown in the popup)
          - Lat, Long                    (coordinates used to place the pin)
          - Region                       (free text scanned for keywords)
          - Work Address - Line 1
          - City
          - Postal or ZIP code           (display-only postal/zip for the address)

Requires a Google Maps API key (Geocoding API + Distance Matrix API, both
enabled with billing set up) -- see GOOGLE_MAPS_API_KEY in the CONFIG
block. You'll be prompted for it in a text box when running interactively
(Colab/Jupyter) if it isn't already configured.

Output: a .zip file containing map.html (and a small README), ready to
        upload to any static host (GitHub Pages, S3, Netlify, etc).
"""

import datetime
import hashlib
import json
import os
import re
import shutil
import time
import urllib.parse
import urllib.request
import zipfile
import warnings
from html import escape
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG - edit these to match your spreadsheet / branding
# ---------------------------------------------------------------------------

# Column names as they appear in the spreadsheet.
COLUMN_GEOCODE_POSTAL = "Worker's Postal"          # shown in the popup / sidebar
COLUMN_LAT = "Lat"                                 # pin latitude, taken as-is from the spreadsheet
COLUMN_LONG = "Long"                               # pin longitude, taken as-is from the spreadsheet
COLUMN_REGION = "Region"                           # free-text, scanned for keywords below
COLUMN_ADDRESS_LINE1 = "Work Address - Line 1"     # matched against ADDRESS_CATEGORIES
COLUMN_CITY = "City"
COLUMN_DISPLAY_POSTAL = "Postal or ZIP code"        # shown in the popup

# Only keep rows whose postal code's first letter maps to one of these
# provinces (standard Canadian postal-code prefix: T -> Alberta, V ->
# British Columbia). Set to None to disable the province filter entirely.
ALLOWED_PROVINCES = {"BC", "AB"}

# First letter of a Canadian postal code -> province/territory abbreviation.
PROVINCE_BY_POSTAL_PREFIX = {
    "A": "NL", "B": "NS", "C": "PE", "E": "NB",
    "G": "QC", "H": "QC", "J": "QC",
    "K": "ON", "L": "ON", "M": "ON", "N": "ON", "P": "ON",
    "R": "MB", "S": "SK", "T": "AB", "V": "BC",
    "X": "NT/NU", "Y": "YT",
}

# Region classification: value -> list of keywords to search for (case-insensitive,
# substring match) inside COLUMN_REGION. First match wins; order matters if a
# row's text could match more than one.
REGION_KEYWORDS = {
    "ClaimsPro": ["claimspro"],
    "SCM": ["scm"],
    "Pario": ["pario"],
}

# Marker color per region. ClaimsPro & SCM intentionally share "Blue".
REGION_COLORS = {
    "ClaimsPro": "#7e9cd1",  # Blue
    "SCM": "#7e9cd1",        # Blue
    "Pario": "#55489d",      # Purple
}

# Work Address - Line 1 categories to filter on. Matching is case-insensitive
# and ignores extra whitespace, but otherwise looks for these as substrings
# of the spreadsheet's address value (so "Suite 112, 6093 Iona Drive, Burnaby"
# still matches "Suite 112, 6093 Iona Drive"). This is also the column name
# matched in the spreadsheet -- COLUMN_ADDRESS_LINE1 above -- but the map's
# own UI labels this filter with ADDRESS_FILTER_LABEL below instead.
ADDRESS_CATEGORIES = [
    "8333 Eastlake Drive Suite 202",
    "1849 Welch Street",
    "Suite 112, 6093 Iona Drive",
]

# Label shown in the map's UI (filter panel heading) for the Work Address
# filter. Purely cosmetic -- COLUMN_ADDRESS_LINE1 (the spreadsheet column
# name being matched) is unaffected.
ADDRESS_FILTER_LABEL = "Work Address"

# Marker/line/legend color per work address, and reflected on the map for
# the destination pins themselves.
ADDRESS_COLORS = {
    "8333 Eastlake Drive Suite 202": "#EF6528",
    "1849 Welch Street": "#4DB595",
    "Suite 112, 6093 Iona Drive": "#7D7370",
}
UNKNOWN_ADDRESS_COLOR = "#555555"

# Label used for rows that don't match any known Region keyword / Address category.
UNKNOWN_REGION_LABEL = "Unclassified"
UNKNOWN_ADDRESS_LABEL = "Other / Unmatched Address"
UNKNOWN_COLOR = "gray"

MAP_TITLE = "Postal codes map - North Vancouver"
MAP_START_LOCATION = [53.7267, -119.0]   # rough BC/AB midpoint
MAP_START_ZOOM = 5

# Background color for the sidebar's title header block.
SIDEBAR_HEADER_BG_COLOR = "#3A4458"
SIDEBAR_HEADER_TEXT_COLOR = "#ffffff"

# Typeface used for all text on the map (filter/sidebar/KPI panels, popups).
# Loaded from Google Fonts; falls back to a normal sans-serif stack if that
# CDN is unreachable when the map is viewed (a missing web font degrades
# gracefully, unlike a missing script, so this isn't vendored locally like
# the JS/CSS libraries are).
FONT_FAMILY = "'Open Sans', Arial, sans-serif"
GOOGLE_FONT_CSS_URL = "https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap"

# Default pin display mode when the map first loads: "individual" (every
# worker pin shown separately, as before) or "cluster" (grouped via
# Leaflet.markercluster). The viewer can switch with an in-page toggle.
DEFAULT_PIN_DISPLAY_MODE = "individual"

# ---------------------------------------------------------------------------
# Drive-time CONFIG (Google Maps Platform: Geocoding API + Distance Matrix
# API, both must be enabled with billing set up on your Google Cloud
# project). Drive time is computed once, at generation time, from each
# worker's postal code (COLUMN_LAT/COLUMN_LONG) to their classified Work
# Address - Line 1 (its canonical text is geocoded once and cached; rows
# classified as UNKNOWN_ADDRESS_LABEL have no fixed destination and are
# skipped). Results are baked into the static page -- there is no live
# API access from the viewer's browser.
# ---------------------------------------------------------------------------

# Supply the key one of three ways (checked in this order), so you never
# have to commit a real key into this file:
#   1. Paste it here, e.g. GOOGLE_MAPS_API_KEY = "AIza..."
#   2. Set the GOOGLE_MAPS_API_KEY environment variable before running.
#   3. Leave both blank and run interactively (Colab/Jupyter) -- you'll be
#      prompted for it in a text box (masked input).
GOOGLE_MAPS_API_KEY = ""

# Both directions/times below are always computed and baked into the
# exported map.html -- the viewer picks between them with a toggle in the
# deployed page itself (there is no build-time prompt for this; a static
# export can't call the API live, so both must be precomputed). "AM" means
# the morning commute (Worker's Postal -> Work Address); "PM" means the
# afternoon commute (Work Address -> Worker's Postal, i.e. the reverse
# direction -- drive time/distance can differ from the AM leg due to
# one-way streets, ramps, and time-of-day traffic).
DEPARTURE_TIME_OPTIONS = {
    "AM": {"label": "7:30 AM", "hour": 7, "minute": 30},
    "PM": {"label": "4:30 PM", "hour": 16, "minute": 30},
}

# Which direction the exported map shows by default when first opened
# ("AM" or "PM"); the viewer can switch with the in-page toggle afterward.
DEFAULT_DIRECTION = "AM"

# Drive times are modeled for the next occurrence of this weekday (0=Monday
# ... 6=Sunday) at the chosen time, in each worker's own province timezone
# -- far enough in the future for Google's traffic-model prediction, and
# representative of a normal weekday commute rather than a specific date.
DRIVE_TIME_TARGET_WEEKDAY = 0  # Monday

PROVINCE_TIMEZONES = {
    "BC": "America/Vancouver",
    "AB": "America/Edmonton",
}

# Distance Matrix API request batching: origins per request x 1 destination
# per request must stay comfortably under the API's per-request element cap.
DISTANCE_MATRIX_MAX_ORIGINS_PER_REQUEST = 25

# Thresholds (minutes) for the "% of pins under N minutes" KPIs.
DRIVE_TIME_THRESHOLDS_MINUTES = [15, 20, 30]

# Style for the straight worker-to-address lines drawn on the map (colored
# per-row to match its region). Destination markers use Leaflet's plain
# default pin (no color option, to avoid depending on the Leaflet.awesome-
# markers plugin -- see build_map).
DISTANCE_LINE_WEIGHT = 1.5
DISTANCE_LINE_OPACITY = 0.55

# Leaflet.markercluster powers the "Clustered" pin display mode (there is
# no clustering in core Leaflet). If this plugin fails to load in the
# viewer's browser, the map falls back to individual pins automatically
# (see the defensive check in the injected JS) rather than breaking.
LEAFLET_MARKERCLUSTER_VERSION = "1.5.3"
LEAFLET_MARKERCLUSTER_JS_URL = (
    f"https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/"
    f"{LEAFLET_MARKERCLUSTER_VERSION}/leaflet.markercluster.js"
)
LEAFLET_MARKERCLUSTER_CSS_URLS = [
    f"https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/"
    f"{LEAFLET_MARKERCLUSTER_VERSION}/MarkerCluster.css",
    f"https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/"
    f"{LEAFLET_MARKERCLUSTER_VERSION}/MarkerCluster.Default.css",
]

OUTPUT_DIR = "postal_code_map_output"
OUTPUT_ZIP = "postal_code_map_output.zip"
OUTPUT_HTML_NAME = "map.html"

# Folium normally links Leaflet/jQuery/Bootstrap/Font Awesome from external
# CDNs. If the machine opening map.html has no internet access, or a
# network/firewall blocks any of those CDNs, the map silently fails to
# render. When True, those library files are downloaded once at generation
# time and bundled into the zip's assets/ folder so map.html works with no
# CDN dependency for the map itself (map tile images still need internet,
# same as any web map). Requires internet access wherever you *run* this
# script (e.g. Colab), not wherever the map is later opened.
VENDOR_LIBS_LOCALLY = True


# ---------------------------------------------------------------------------
# Step 1: Load & clean the spreadsheet
# ---------------------------------------------------------------------------

def _normalize_column_name(name):
    """Normalize a header for matching: lowercase, unify curly quotes/dashes,
    collapse whitespace, and strip surrounding punctuation."""
    text = str(name).strip().lower()
    text = text.replace("’", "'").replace("‘", "'")
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text


def load_spreadsheet(path):
    """Read the input .xlsx into a DataFrame and check required columns exist.

    Column matching is normalized (case/whitespace/apostrophe/dash-insensitive)
    so headers like "worker's postal ", "Worker’s Postal", or "WORKER'S POSTAL"
    all resolve to the configured COLUMN_GEOCODE_POSTAL, etc. Matched columns
    are renamed to the exact configured names so the rest of the pipeline can
    rely on them as-is.
    """
    df = pd.read_excel(path)

    required = [
        COLUMN_GEOCODE_POSTAL,
        COLUMN_LAT,
        COLUMN_LONG,
        COLUMN_REGION,
        COLUMN_ADDRESS_LINE1,
        COLUMN_CITY,
        COLUMN_DISPLAY_POSTAL,
    ]

    actual_by_normalized = {}
    for col in df.columns:
        key = _normalize_column_name(col)
        actual_by_normalized.setdefault(key, col)

    rename_map = {}
    missing = []
    for wanted in required:
        key = _normalize_column_name(wanted)
        actual = actual_by_normalized.get(key)
        if actual is None:
            missing.append(wanted)
        elif actual != wanted:
            rename_map[actual] = wanted

    if missing:
        raise ValueError(
            f"Spreadsheet is missing required column(s): {missing}. "
            f"Found columns: {list(df.columns)}"
        )

    if rename_map:
        df = df.rename(columns=rename_map)

    df = df.dropna(subset=[COLUMN_GEOCODE_POSTAL]).copy()
    df[COLUMN_GEOCODE_POSTAL] = df[COLUMN_GEOCODE_POSTAL].astype(str).str.strip()
    return df


# ---------------------------------------------------------------------------
# Step 2: Read pin coordinates directly from the spreadsheet
# ---------------------------------------------------------------------------

def prepare_coordinates(df):
    """Use the spreadsheet's own COLUMN_LAT/COLUMN_LONG values as each pin's
    location (no geocoding needed), and derive a province code from the
    first letter of the postal code to apply ALLOWED_PROVINCES."""
    df = df.copy()
    df["latitude"] = pd.to_numeric(df[COLUMN_LAT], errors="coerce")
    df["longitude"] = pd.to_numeric(df[COLUMN_LONG], errors="coerce")

    valid_range = (
        df["latitude"].between(-90, 90) & df["longitude"].between(-180, 180)
    )
    before = len(df)
    df = df[df["latitude"].notna() & df["longitude"].notna() & valid_range].copy()
    dropped = before - len(df)
    if dropped:
        warnings.warn(
            f"Dropped {dropped} row(s) with missing or invalid {COLUMN_LAT}/{COLUMN_LONG} values."
        )

    first_letter = df[COLUMN_GEOCODE_POSTAL].str.strip().str.upper().str[0]
    df["province_code"] = first_letter.map(PROVINCE_BY_POSTAL_PREFIX)

    if ALLOWED_PROVINCES:
        before = len(df)
        df = df[df["province_code"].isin(ALLOWED_PROVINCES)].copy()
        dropped = before - len(df)
        if dropped:
            warnings.warn(
                f"Dropped {dropped} row(s) outside of allowed provinces {ALLOWED_PROVINCES}."
            )

    return df


# ---------------------------------------------------------------------------
# Step 3: Classify Region and Work Address into filter categories
# ---------------------------------------------------------------------------

def _normalize(text):
    return re.sub(r"\s+", " ", str(text)).strip().lower()


def classify_region(raw_value):
    text = _normalize(raw_value)
    for region, keywords in REGION_KEYWORDS.items():
        if any(keyword.lower() in text for keyword in keywords):
            return region
    return UNKNOWN_REGION_LABEL


def classify_address(raw_value):
    text = _normalize(raw_value)
    for address in ADDRESS_CATEGORIES:
        if _normalize(address) in text:
            return address
    return UNKNOWN_ADDRESS_LABEL


def classify_rows(df):
    df["region_category"] = df[COLUMN_REGION].apply(classify_region)
    df["address_category"] = df[COLUMN_ADDRESS_LINE1].apply(classify_address)
    df["marker_color"] = df["region_category"].map(REGION_COLORS).fillna(UNKNOWN_COLOR)
    return df


# ---------------------------------------------------------------------------
# Step 4: Compute drive times (Google Maps Geocoding + Distance Matrix APIs)
# ---------------------------------------------------------------------------

GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
GOOGLE_DISTANCE_MATRIX_URL = "https://maps.googleapis.com/maps/api/distancematrix/json"


def _get_google_maps_api_key():
    if GOOGLE_MAPS_API_KEY:
        return GOOGLE_MAPS_API_KEY

    env_key = os.environ.get("GOOGLE_MAPS_API_KEY")
    if env_key:
        return env_key

    if _running_in_colab() or _running_in_notebook():
        import getpass

        key = getpass.getpass("Enter your Google Maps API key (Geocoding + Distance Matrix): ").strip()
        if key:
            return key

    raise RuntimeError(
        "No Google Maps API key found. Set GOOGLE_MAPS_API_KEY at the top of "
        "generate_map.py, set a GOOGLE_MAPS_API_KEY environment variable, or "
        "run this interactively (Colab/Jupyter) to be prompted for one."
    )


def _load_timezone(tz_name):
    try:
        return ZoneInfo(tz_name)
    except ZoneInfoNotFoundError:
        # Some minimal environments (certain Windows/Docker setups) ship
        # Python without the system IANA tz database. Install the tzdata
        # package (pure-data, no compiled deps) and retry once.
        import subprocess
        import sys

        print("System timezone database not found; installing 'tzdata'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "tzdata"])
        return ZoneInfo(tz_name)


def _next_departure_epoch(hour, minute, tz_name, target_weekday=DRIVE_TIME_TARGET_WEEKDAY):
    """Unix timestamp for the next occurrence of target_weekday at
    hour:minute in the given IANA timezone, always strictly in the future
    (at least one full day out), for Google's traffic-model prediction."""
    tz = _load_timezone(tz_name)
    now = datetime.datetime.now(tz)
    days_ahead = (target_weekday - now.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7  # always at least a week out, never "later today"
    target_date = now.date() + datetime.timedelta(days=days_ahead)
    target_dt = datetime.datetime(
        target_date.year, target_date.month, target_date.day, hour, minute, tzinfo=tz
    )
    return int(target_dt.timestamp())


def _geocode_address_text(address_text, api_key):
    """Geocode a free-text address (used for the fixed office addresses, not
    postal codes) via the Geocoding API, biased to Canada."""
    params = {"address": address_text, "region": "ca", "key": api_key}
    url = GOOGLE_GEOCODE_URL + "?" + urllib.parse.urlencode(params)
    with urllib.request.urlopen(url, timeout=20) as resp:
        payload = json.loads(resp.read().decode("utf-8"))

    status = payload.get("status")
    if status == "OK":
        location = payload["results"][0]["geometry"]["location"]
        return location["lat"], location["lng"]

    raise RuntimeError(
        f"Google Geocoding API returned status={status!r} for address "
        f"'{address_text}' ({payload.get('error_message', 'no further detail')}). "
        f"Check that the Geocoding API is enabled and billing is set up for "
        f"this API key in Google Cloud Console."
    )


def geocode_work_addresses(df, api_key):
    """Geocode each distinct, classified Work Address - Line 1 category
    present in df (skipping UNKNOWN_ADDRESS_LABEL, which has no fixed
    canonical text). Returns {address_category: (lat, lng)}."""
    categories = sorted(
        c for c in df["address_category"].unique() if c != UNKNOWN_ADDRESS_LABEL
    )
    return {category: _geocode_address_text(category, api_key) for category in categories}


def _distance_matrix_request(origins, destinations, departure_epoch, api_key):
    """Raw Distance Matrix API call: origins and destinations are lists of
    (lat, lng). Returns the parsed 'rows' list (one row per origin, each
    with an 'elements' list of one entry per destination) -- no batching,
    caller must keep origins*destinations under the per-request cap."""
    params = {
        "origins": "|".join(f"{lat},{lng}" for lat, lng in origins),
        "destinations": "|".join(f"{lat},{lng}" for lat, lng in destinations),
        "departure_time": str(departure_epoch),
        "traffic_model": "best_guess",
        "mode": "driving",
        "key": api_key,
    }
    url = GOOGLE_DISTANCE_MATRIX_URL + "?" + urllib.parse.urlencode(params)
    with urllib.request.urlopen(url, timeout=30) as resp:
        payload = json.loads(resp.read().decode("utf-8"))

    status = payload.get("status")
    if status != "OK":
        raise RuntimeError(
            f"Google Distance Matrix API returned status={status!r} "
            f"({payload.get('error_message', 'no further detail')}). Check that "
            f"the Distance Matrix API is enabled and billing is set up for "
            f"this API key in Google Cloud Console."
        )
    return payload["rows"]


def _element_to_minutes_km(element):
    if element.get("status") != "OK":
        return None, None
    duration_seconds = (
        element["duration_in_traffic"]["value"]
        if "duration_in_traffic" in element
        else element["duration"]["value"]
    )
    return duration_seconds / 60.0, element["distance"]["value"] / 1000.0


def _distance_matrix_many_origins_one_destination(origin_latlngs, destination_latlng, departure_epoch, api_key):
    """Drive time/distance from each of many origins to a single destination
    (the AM, Worker -> Address direction). Returns a list of (minutes, km)
    in the same order as origin_latlngs, batching internally to stay under
    DISTANCE_MATRIX_MAX_ORIGINS_PER_REQUEST origins per request."""
    results = []
    for start in range(0, len(origin_latlngs), DISTANCE_MATRIX_MAX_ORIGINS_PER_REQUEST):
        batch = origin_latlngs[start : start + DISTANCE_MATRIX_MAX_ORIGINS_PER_REQUEST]
        rows = _distance_matrix_request(batch, [destination_latlng], departure_epoch, api_key)
        for row in rows:
            results.append(_element_to_minutes_km(row["elements"][0]))
        time.sleep(0.05)
    return results


def _distance_matrix_one_origin_many_destinations(origin_latlng, destination_latlngs, departure_epoch, api_key):
    """Drive time/distance from a single origin to each of many destinations
    (the PM, Address -> Worker direction). Returns a list of (minutes, km)
    in the same order as destination_latlngs, batching internally to stay
    under DISTANCE_MATRIX_MAX_ORIGINS_PER_REQUEST destinations per request."""
    results = []
    for start in range(0, len(destination_latlngs), DISTANCE_MATRIX_MAX_ORIGINS_PER_REQUEST):
        batch = destination_latlngs[start : start + DISTANCE_MATRIX_MAX_ORIGINS_PER_REQUEST]
        rows = _distance_matrix_request([origin_latlng], batch, departure_epoch, api_key)
        elements = rows[0]["elements"]
        for element in elements:
            results.append(_element_to_minutes_km(element))
        time.sleep(0.05)
    return results


def compute_drive_times(df, api_key):
    """Add drive_minutes_am/distance_km_am (Worker's Postal -> Work Address,
    modeled at 7:30 AM) and drive_minutes_pm/distance_km_pm (Work Address ->
    Worker's Postal, modeled at 4:30 PM) columns, plus dest_lat/dest_lng,
    by querying the Distance Matrix API once per (address_category,
    province, direction) group (batched), using each province's own local
    timezone. Rows with UNKNOWN_ADDRESS_LABEL or an unresolvable province
    timezone get NaN in both directions and are excluded from drive-time
    KPIs regardless of which direction is toggled."""
    df = df.copy()
    for direction in DEPARTURE_TIME_OPTIONS:
        df[f"drive_minutes_{direction.lower()}"] = pd.NA
        df[f"distance_km_{direction.lower()}"] = pd.NA
    df["dest_lat"] = pd.NA
    df["dest_lng"] = pd.NA

    address_coords = geocode_work_addresses(df, api_key)
    for category, (lat, lng) in address_coords.items():
        mask = df["address_category"] == category
        df.loc[mask, "dest_lat"] = lat
        df.loc[mask, "dest_lng"] = lng

    routable = df[df["address_category"] != UNKNOWN_ADDRESS_LABEL]
    skipped_no_address = len(df) - len(routable)
    if skipped_no_address:
        warnings.warn(
            f"{skipped_no_address} row(s) have an unmatched Work Address - Line 1 "
            f"and were skipped for drive-time calculation (both directions)."
        )

    skipped_no_tz = 0
    for (category, province), group in routable.groupby(["address_category", "province_code"]):
        tz_name = PROVINCE_TIMEZONES.get(province)
        if tz_name is None:
            skipped_no_tz += len(group)
            continue

        destination = address_coords[category]
        origins = list(zip(group["latitude"], group["longitude"]))

        am = DEPARTURE_TIME_OPTIONS["AM"]
        am_epoch = _next_departure_epoch(am["hour"], am["minute"], tz_name)
        am_results = _distance_matrix_many_origins_one_destination(
            origins, destination, am_epoch, api_key
        )
        for idx, (minutes, km) in zip(group.index, am_results):
            df.at[idx, "drive_minutes_am"] = minutes
            df.at[idx, "distance_km_am"] = km

        pm = DEPARTURE_TIME_OPTIONS["PM"]
        pm_epoch = _next_departure_epoch(pm["hour"], pm["minute"], tz_name)
        pm_results = _distance_matrix_one_origin_many_destinations(
            destination, origins, pm_epoch, api_key
        )
        for idx, (minutes, km) in zip(group.index, pm_results):
            df.at[idx, "drive_minutes_pm"] = minutes
            df.at[idx, "distance_km_pm"] = km

    if skipped_no_tz:
        warnings.warn(
            f"{skipped_no_tz} row(s) have a province with no configured timezone "
            f"in PROVINCE_TIMEZONES and were skipped for drive-time calculation."
        )

    for direction in DEPARTURE_TIME_OPTIONS:
        suffix = direction.lower()
        df[f"drive_minutes_{suffix}"] = pd.to_numeric(df[f"drive_minutes_{suffix}"], errors="coerce")
        df[f"distance_km_{suffix}"] = pd.to_numeric(df[f"distance_km_{suffix}"], errors="coerce")

        no_route = df[f"drive_minutes_{suffix}"].isna().sum() - skipped_no_address - skipped_no_tz
        if no_route > 0:
            warnings.warn(
                f"{no_route} row(s) could not be routed by the Distance Matrix API "
                f"for the {DEPARTURE_TIME_OPTIONS[direction]['label']} direction "
                f"(no driving route found)."
            )

    return df


# ---------------------------------------------------------------------------
# Step 5: Build the popup text
# ---------------------------------------------------------------------------

def _nullable_float(value):
    return float(value) if pd.notna(value) else None


def build_popup_html(row):
    address_line = ", ".join(
        str(part).strip()
        for part in [
            row[COLUMN_ADDRESS_LINE1],
            row[COLUMN_CITY],
            row[COLUMN_DISPLAY_POSTAL],
        ]
        if pd.notna(part) and str(part).strip()
    )
    def drive_line(direction):
        minutes = row.get(f"drive_minutes_{direction.lower()}")
        km = row.get(f"distance_km_{direction.lower()}")
        return f"{minutes:.1f} min ({km:.1f} km)" if pd.notna(minutes) else "no route computed"

    am_label = DEPARTURE_TIME_OPTIONS["AM"]["label"]
    pm_label = DEPARTURE_TIME_OPTIONS["PM"]["label"]

    return (
        f"<b>Worker's Postal:</b> {escape(str(row[COLUMN_GEOCODE_POSTAL]))}<br>"
        f"<b>Region:</b> {escape(str(row['region_category']))}<br>"
        f"<b>Address:</b> {escape(address_line)}<br>"
        f"<b>{escape(am_label)} (Home → Work):</b> {escape(drive_line('AM'))}<br>"
        f"<b>{escape(pm_label)} (Work → Home):</b> {escape(drive_line('PM'))}"
    )


# ---------------------------------------------------------------------------
# Step 6: Build the Folium map with combinable Region / Address checkboxes
# ---------------------------------------------------------------------------

def build_map(df):
    import folium
    from folium import Element

    m = folium.Map(location=MAP_START_LOCATION, zoom_start=MAP_START_ZOOM, tiles="OpenStreetMap")
    m.get_root().html.add_child(Element(f"<title>{escape(MAP_TITLE)}</title>"))

    head_extras = [f'<link rel="stylesheet" href="{GOOGLE_FONT_CSS_URL}"/>']
    head_extras += [f'<link rel="stylesheet" href="{url}"/>' for url in LEAFLET_MARKERCLUSTER_CSS_URLS]
    head_extras.append(f'<script src="{LEAFLET_MARKERCLUSTER_JS_URL}"></script>')
    head_extras.append(
        f"<style>body, .leaflet-container, .leaflet-popup-content {{"
        f"font-family: {FONT_FAMILY} !important; }}</style>"
    )
    for tag in head_extras:
        m.get_root().header.add_child(Element(tag))

    region_labels = list(REGION_KEYWORDS.keys())
    if (df["region_category"] == UNKNOWN_REGION_LABEL).any():
        region_labels.append(UNKNOWN_REGION_LABEL)

    address_labels = list(ADDRESS_CATEGORIES)
    if (df["address_category"] == UNKNOWN_ADDRESS_LABEL).any():
        address_labels.append(UNKNOWN_ADDRESS_LABEL)

    # One destination marker per work address that actually has coordinates
    # (i.e. was successfully geocoded), keyed by address category so the
    # filter JS can toggle it alongside that address's worker pins/lines.
    # Colored per ADDRESS_COLORS and drawn as a CircleMarker (bigger/bordered
    # to stand out from worker pins) rather than folium's default Marker
    # icon, since a custom color there would require the Leaflet.awesome-
    # markers plugin -- the same secondary-CDN dependency that has already
    # caused pins to silently disappear once in this project.
    dest_markers = {}
    has_dest = df["dest_lat"].notna() & df["dest_lng"].notna()
    for category, group in df[has_dest].groupby("address_category"):
        dest_lat = float(group["dest_lat"].iloc[0])
        dest_lng = float(group["dest_lng"].iloc[0])
        color = ADDRESS_COLORS.get(category, UNKNOWN_ADDRESS_COLOR)
        dest_marker = folium.CircleMarker(
            location=[dest_lat, dest_lng],
            radius=11,
            color="#ffffff",
            weight=2,
            fill=True,
            fill_color=color,
            fill_opacity=0.95,
            popup=folium.Popup(f"<b>Work Address:</b> {escape(category)}", max_width=320),
        )
        dest_marker.add_to(m)
        dest_markers[category] = {
            "var": dest_marker.get_name(),
            "lat": dest_lat,
            "lng": dest_lng,
        }

    marker_meta = []
    for _, row in df.iterrows():
        # CircleMarker only needs core Leaflet (no icon plugin), so a pin
        # always renders even if a secondary CDN (e.g. Leaflet.awesome-markers)
        # is unreachable.
        marker = folium.CircleMarker(
            location=[row["latitude"], row["longitude"]],
            radius=8,
            color=row["marker_color"],
            weight=2,
            fill=True,
            fill_color=row["marker_color"],
            fill_opacity=0.85,
            popup=folium.Popup(build_popup_html(row), max_width=320),
        )
        marker.add_to(m)

        has_dest_coords = pd.notna(row.get("dest_lat")) and pd.notna(row.get("dest_lng"))

        line_var = None
        if has_dest_coords:
            # One straight line per worker; geometry is direction-independent
            # (same two endpoints either way), so it isn't duplicated per AM/PM.
            line = folium.PolyLine(
                locations=[
                    [row["latitude"], row["longitude"]],
                    [row["dest_lat"], row["dest_lng"]],
                ],
                color=row["marker_color"],
                weight=DISTANCE_LINE_WEIGHT,
                opacity=DISTANCE_LINE_OPACITY,
            )
            line.add_to(m)
            line_var = line.get_name()

        marker_meta.append(
            {
                "var": marker.get_name(),
                "line_var": line_var,
                "region": row["region_category"],
                "address": row["address_category"],
                "postal": row[COLUMN_GEOCODE_POSTAL],
                "lat": row["latitude"],
                "lng": row["longitude"],
                "drive_minutes_am": _nullable_float(row.get("drive_minutes_am")),
                "distance_km_am": _nullable_float(row.get("distance_km_am")),
                "drive_minutes_pm": _nullable_float(row.get("drive_minutes_pm")),
                "distance_km_pm": _nullable_float(row.get("distance_km_pm")),
            }
        )

    _add_controls(m, marker_meta, region_labels, address_labels, dest_markers)
    return m


# Zoom level used when clicking a sidebar entry to jump to its pin.
SIDEBAR_ZOOM_LEVEL = 14


_WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


class _JsVarRef:
    """Marks a string as a raw JS identifier (e.g. a Leaflet layer's
    variable name) rather than a value to be JSON-string-escaped."""

    def __init__(self, name):
        self.name = name


def _add_controls(m, marker_meta, region_labels, address_labels, dest_markers):
    """Inject the custom Leaflet controls: an AM/PM departure-time toggle,
    two checkbox groups (Region, Work Address), a left-hand sidebar listing
    every pin by Worker's Postal (with its own per-worker include/exclude
    checkbox), and a KPI panel. A pin's map visibility requires its region
    AND address to be checked AND its own worker checkbox to be checked;
    the sidebar row itself only follows the Region/Address filters
    (unchecking a worker deselects it without hiding it from the list).
    The KPI panel and each sidebar row's drive time reflect whichever
    direction (AM/PM) is currently toggled, computed from whichever pins
    are currently included (filters + individually checked) and have a
    valid drive time for that direction. Clicking a pin's postal-code
    label zooms/pans to it. Both directions' data are already baked into
    marker_meta at generation time -- the toggle only ever switches
    between them client-side, no API access from the viewer's browser."""
    from folium import Element

    map_var = m.get_name()

    def color_swatch_html(color):
        return (
            f'<span style="display:inline-block; width:11px; height:11px; '
            f'border-radius:50%; background:{escape(color)}; border:1px solid #333; '
            f'margin-right:6px; vertical-align:middle;"></span>'
        )

    def checkbox_html(group, values, color_map=None):
        items = []
        for v in values:
            safe_id = f"{group}_{re.sub(r'[^a-zA-Z0-9]', '_', v)}"
            swatch = color_swatch_html(color_map[v]) if color_map else ""
            items.append(
                f'<label style="display:block;font-weight:normal;margin:2px 0;">'
                f'<input type="checkbox" class="{group}-filter" value="{escape(v)}" '
                f'id="{safe_id}" checked> {swatch}{escape(v)}</label>'
            )
        return "\n".join(items)

    region_colors = {label: REGION_COLORS.get(label, UNKNOWN_COLOR) for label in region_labels}
    address_colors = {label: ADDRESS_COLORS.get(label, UNKNOWN_ADDRESS_COLOR) for label in address_labels}

    weekday_name = _WEEKDAY_NAMES[DRIVE_TIME_TARGET_WEEKDAY % 7]

    def direction_radio_html(direction, info):
        checked = "checked" if direction == DEFAULT_DIRECTION else ""
        arrow = "Home → Work" if direction == "AM" else "Work → Home"
        return (
            f'<label style="display:block;font-weight:normal;margin:2px 0;">'
            f'<input type="radio" name="postal-map-direction" class="direction-toggle" '
            f'value="{direction}" {checked}> {escape(info["label"])} ({escape(arrow)})</label>'
        )

    direction_radios_html = "\n".join(
        direction_radio_html(direction, info) for direction, info in DEPARTURE_TIME_OPTIONS.items()
    )

    kpi_rows_html = "".join(
        f'<div class="postal-map-kpi-row"><span>Share under {t} min</span>'
        f'<span id="postal-map-kpi-pct-{t}">—</span></div>'
        for t in DRIVE_TIME_THRESHOLDS_MINUTES
    )

    pin_display_radios_html = "\n".join(
        f'<label style="display:block;font-weight:normal;margin:2px 0;">'
        f'<input type="radio" name="postal-map-pin-display" class="pin-display-toggle" '
        f'value="{value}" {"checked" if value == DEFAULT_PIN_DISPLAY_MODE else ""}> {label}</label>'
        for value, label in [("individual", "Individual pins"), ("cluster", "Clustered")]
    )

    control_html = f"""
    <style>#postal-map-stack button {{ font-family: {FONT_FAMILY}; }}</style>
    <div id="postal-map-stack" style="
        position: fixed; top: 10px; left: 10px; z-index: 9999;
        display: flex; flex-direction: column; gap: 10px;
        width: 270px; max-height: 96vh;
        font-family: {FONT_FAMILY}; font-size: 13px;">

      <div id="postal-map-filter-panel" style="
          background: white; padding: 10px 14px; border: 2px solid #444;
          border-radius: 6px; flex: 0 0 auto;
          max-height: 46vh; overflow-y: auto; box-shadow: 2px 2px 6px rgba(0,0,0,0.3);">
        <div style="font-weight:bold; margin-top:0;">Departure Time</div>
        {direction_radios_html}
        <div style="color:#777; font-size:11px; margin-top:2px;">
          Modeled for next {escape(weekday_name)}, each worker's local province time.
        </div>

        <div style="font-weight:bold; margin-top:10px;">Pin Display</div>
        {pin_display_radios_html}

        <div style="font-weight:bold; margin-top:10px;">Region</div>
        {checkbox_html("region", region_labels, color_map=region_colors)}
        <button id="region-select-all" style="margin-top:4px;">All</button>
        <button id="region-select-none">None</button>

        <div style="font-weight:bold; margin-top:10px;">{escape(ADDRESS_FILTER_LABEL)}</div>
        {checkbox_html("address", address_labels, color_map=address_colors)}
        <button id="address-select-all" style="margin-top:4px;">All</button>
        <button id="address-select-none">None</button>
      </div>

      <div id="postal-map-sidebar" style="
          background: white; border: 2px solid #444; border-radius: 6px;
          flex: 1 1 auto; min-height: 120px; display: flex; flex-direction: column;
          overflow: hidden; box-shadow: 2px 2px 6px rgba(0,0,0,0.3);">
        <div style="background:{SIDEBAR_HEADER_BG_COLOR}; color:{SIDEBAR_HEADER_TEXT_COLOR};
            font-weight:bold; padding:10px 14px; border-radius:4px 4px 0 0;">
          {escape(MAP_TITLE)}
        </div>
        <div style="font-weight:bold; padding:10px 14px 4px 14px;">
          Pins (Worker's Postal)
        </div>
        <div id="postal-map-pin-count" style="padding:0 14px 4px 14px; color:#555;"></div>
        <div style="padding:0 14px 6px 14px;">
          <button id="worker-select-all">All</button>
          <button id="worker-select-none">None</button>
        </div>
        <div id="postal-map-pin-list" style="overflow-y:auto; flex:1; min-height:0; padding:0 6px 8px 6px;"></div>
      </div>

      <div id="postal-map-kpi-panel" style="
          background: white; border: 2px solid #444; border-radius: 6px;
          flex: 0 0 auto; max-height: 34vh; overflow-y: auto;
          box-shadow: 2px 2px 6px rgba(0,0,0,0.3); padding: 10px 14px;">
        <div style="font-weight:bold; margin-bottom:4px;">Drive Time KPIs</div>
        <div id="postal-map-kpi-direction-note" style="color:#555; font-size:11px; margin-bottom:8px;"></div>
        <style>
          #postal-map-kpi-panel .postal-map-kpi-row {{
            display:flex; justify-content:space-between; gap:8px; margin:2px 0;
          }}
          #postal-map-kpi-panel .postal-map-kpi-row span:last-child {{ font-weight:bold; }}
        </style>
        <div class="postal-map-kpi-row"><span>Base size</span><span id="postal-map-kpi-base">—</span></div>
        <div class="postal-map-kpi-row"><span>Average time</span><span id="postal-map-kpi-avg-time">—</span></div>
        <div class="postal-map-kpi-row"><span>Median time</span><span id="postal-map-kpi-median-time">—</span></div>
        {kpi_rows_html}
        <div class="postal-map-kpi-row"><span>Average distance</span><span id="postal-map-kpi-avg-dist">—</span></div>
        <div class="postal-map-kpi-row"><span>Closest</span><span id="postal-map-kpi-closest">—</span></div>
        <div class="postal-map-kpi-row"><span>Farthest</span><span id="postal-map-kpi-farthest">—</span></div>
      </div>
    </div>
    """

    def _js_literal(value):
        """A JS object-literal fragment for one value: json.dumps handles
        strings/numbers/null safely; a bare (unquoted) var name is passed
        through as-is so it refers to the actual Leaflet object variable,
        not a string of its name."""
        if isinstance(value, _JsVarRef):
            return value.name
        return json.dumps(value)

    marker_meta_js = ",\n".join(
        "    {"
        + ", ".join(
            f"{key}: {_js_literal(value)}"
            for key, value in {
                "var": _JsVarRef(meta["var"]),
                "lineVar": _JsVarRef(meta["line_var"]) if meta["line_var"] else None,
                "region": meta["region"],
                "address": meta["address"],
                "postal": str(meta["postal"]),
                "lat": float(meta["lat"]),
                "lng": float(meta["lng"]),
                "driveMinutesAM": meta["drive_minutes_am"],
                "distanceKmAM": meta["distance_km_am"],
                "driveMinutesPM": meta["drive_minutes_pm"],
                "distanceKmPM": meta["distance_km_pm"],
            }.items()
        )
        + "}"
        for meta in marker_meta
    )

    dest_markers_js = ", ".join(
        f"{json.dumps(address)}: {{var: {info['var']}}}"
        for address, info in dest_markers.items()
    )

    region_order_js = ", ".join(json.dumps(r) for r in region_labels)
    address_order_js = ", ".join(json.dumps(a) for a in address_labels)
    region_colors_js = ", ".join(
        f"{json.dumps(r)}: {json.dumps(region_colors[r])}" for r in region_labels
    )
    thresholds_js = ", ".join(str(t) for t in DRIVE_TIME_THRESHOLDS_MINUTES)
    direction_labels_js = ", ".join(
        f"{json.dumps(direction)}: {json.dumps(info['label'])}"
        for direction, info in DEPARTURE_TIME_OPTIONS.items()
    )
    default_direction_js = json.dumps(DEFAULT_DIRECTION)

    filter_js = f"""
    (function() {{
      function init() {{
        var markerInfo = [
{marker_meta_js}
        ];
        var destMarkers = {{{dest_markers_js}}};
        var regionOrder = [{region_order_js}];
        var addressOrder = [{address_order_js}];
        var regionColors = {{{region_colors_js}}};
        var kpiThresholds = [{thresholds_js}];
        var directionLabels = {{{direction_labels_js}}};
        var currentDirection = {default_direction_js};
        var pinDisplayMode = {json.dumps(DEFAULT_PIN_DISPLAY_MODE)};
        // Leaflet.markercluster has no core-Leaflet fallback; if that plugin
        // failed to load (e.g. its CDN was unreachable), clusterGroup stays
        // null and the map just behaves as "individual pins" always, rather
        // than throwing.
        var clusterGroup = (typeof L.markerClusterGroup === 'function') ? L.markerClusterGroup() : null;
        if (!clusterGroup) {{
          document.querySelectorAll('.pin-display-toggle').forEach(function(radio) {{
            if (radio.value === 'cluster') {{ radio.disabled = true; }}
          }});
        }}

        markerInfo.forEach(function(info) {{
          info.workerChecked = true;
          info.addedDirect = false;
          info.addedToCluster = false;
        }});

        function driveMinutesFor(info) {{
          return currentDirection === 'AM' ? info.driveMinutesAM : info.driveMinutesPM;
        }}
        function distanceKmFor(info) {{
          return currentDirection === 'AM' ? info.distanceKmAM : info.distanceKmPM;
        }}

        var pinListEl = document.getElementById('postal-map-pin-list');
        var pinCountEl = document.getElementById('postal-map-pin-count');
        var directionNoteEl = document.getElementById('postal-map-kpi-direction-note');

        function groupKey(region) {{
          return 'postal-map-group-' + region.replace(/[^a-zA-Z0-9]/g, '_');
        }}

        function formatMinutes(minutes) {{
          return minutes === null || minutes === undefined ? 'no route' : minutes.toFixed(1) + ' min';
        }}

        // Group pins by region (in the same order as the Region filter) so
        // the sidebar mirrors the map's color coding. Each region gets a
        // color-swatch header followed by its own pins, built as one
        // contiguous block so later regions don't get interleaved.
        var groupHeaderEls = {{}};
        var groupRowsEls = {{}};
        regionOrder.forEach(function(region) {{
          var header = document.createElement('div');
          header.id = groupKey(region);
          header.style.display = 'flex';
          header.style.alignItems = 'center';
          header.style.fontWeight = 'bold';
          header.style.margin = '8px 4px 2px 4px';
          var swatch = document.createElement('span');
          swatch.style.display = 'inline-block';
          swatch.style.width = '11px';
          swatch.style.height = '11px';
          swatch.style.borderRadius = '50%';
          swatch.style.background = regionColors[region] || '#999';
          swatch.style.border = '1px solid #333';
          swatch.style.marginRight = '6px';
          header.appendChild(swatch);
          var label = document.createElement('span');
          label.textContent = region;
          header.appendChild(label);
          pinListEl.appendChild(header);
          groupHeaderEls[region] = header;
          groupRowsEls[region] = [];
        }});

        markerInfo.forEach(function(info) {{
          var row = document.createElement('div');
          row.className = 'postal-map-pin-row';
          row.style.display = 'flex';
          row.style.alignItems = 'center';
          row.style.padding = '4px 8px 4px 22px';
          row.style.margin = '2px 0';
          row.style.borderRadius = '4px';
          row.title = 'Region: ' + info.region + ' | ' + info.address;
          row.addEventListener('mouseenter', function() {{ row.style.background = '#eef3fb'; }});
          row.addEventListener('mouseleave', function() {{ row.style.background = ''; }});

          var checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.checked = true;
          checkbox.className = 'worker-toggle';
          checkbox.style.marginRight = '6px';
          checkbox.addEventListener('change', function() {{
            info.workerChecked = checkbox.checked;
            recomputeAll();
          }});
          row.appendChild(checkbox);

          var labelSpan = document.createElement('span');
          labelSpan.style.cursor = 'pointer';
          labelSpan.style.flex = '1';
          labelSpan.addEventListener('click', function() {{
            {map_var}.setView([info.lat, info.lng], {SIDEBAR_ZOOM_LEVEL});
            info.var.openPopup();
          }});
          row.appendChild(labelSpan);
          info.labelSpanEl = labelSpan;

          var rowsForRegion = groupRowsEls[info.region];
          if (rowsForRegion) {{
            // Insert right after the last row already placed for this
            // region (or right after its header, if this is the first).
            var previous = rowsForRegion.length
              ? rowsForRegion[rowsForRegion.length - 1]
              : groupHeaderEls[info.region];
            previous.parentNode.insertBefore(row, previous.nextSibling);
            rowsForRegion.push(row);
          }} else {{
            pinListEl.appendChild(row);
          }}
          info.rowEl = row;
          info.checkboxEl = checkbox;
        }});

        function checkedValues(cls) {{
          var boxes = document.querySelectorAll('.' + cls);
          var vals = [];
          boxes.forEach(function(b) {{ if (b.checked) vals.push(b.value); }});
          return vals;
        }}

        function median(sortedNums) {{
          var n = sortedNums.length;
          if (n === 0) {{ return null; }}
          var mid = Math.floor(n / 2);
          return n % 2 ? sortedNums[mid] : (sortedNums[mid - 1] + sortedNums[mid]) / 2;
        }}

        function updateKpiPanel(included) {{
          var byId = function(id) {{ return document.getElementById(id); }};
          var base = included.length;
          byId('postal-map-kpi-base').textContent = base + ' ' + (base === 1 ? 'employee' : 'employees');

          if (base === 0) {{
            byId('postal-map-kpi-avg-time').textContent = '—';
            byId('postal-map-kpi-median-time').textContent = '—';
            byId('postal-map-kpi-avg-dist').textContent = '—';
            byId('postal-map-kpi-closest').textContent = '—';
            byId('postal-map-kpi-farthest').textContent = '—';
            kpiThresholds.forEach(function(t) {{
              var el = byId('postal-map-kpi-pct-' + t);
              if (el) {{ el.textContent = '—'; }}
            }});
            return;
          }}

          var times = included.map(driveMinutesFor);
          var dists = included.map(distanceKmFor);
          var sortedTimes = times.slice().sort(function(a, b) {{ return a - b; }});
          var avgTime = times.reduce(function(a, b) {{ return a + b; }}, 0) / base;
          var avgDist = dists.reduce(function(a, b) {{ return a + b; }}, 0) / base;

          byId('postal-map-kpi-avg-time').textContent = avgTime.toFixed(1) + ' min';
          byId('postal-map-kpi-median-time').textContent = median(sortedTimes).toFixed(1) + ' min';
          byId('postal-map-kpi-avg-dist').textContent = avgDist.toFixed(1) + ' km';

          kpiThresholds.forEach(function(t) {{
            var el = byId('postal-map-kpi-pct-' + t);
            if (!el) {{ return; }}
            var count = times.filter(function(x) {{ return x <= t; }}).length;
            el.textContent = (100 * count / base).toFixed(0) + '%';
          }});

          var closest = included.reduce(function(a, b) {{
            return driveMinutesFor(b) < driveMinutesFor(a) ? b : a;
          }});
          var farthest = included.reduce(function(a, b) {{
            return driveMinutesFor(b) > driveMinutesFor(a) ? b : a;
          }});
          byId('postal-map-kpi-closest').textContent = closest.postal + ' (' + driveMinutesFor(closest).toFixed(1) + ' min)';
          byId('postal-map-kpi-farthest').textContent = farthest.postal + ' (' + driveMinutesFor(farthest).toFixed(1) + ' min)';
        }}

        function recomputeAll() {{
          var regions = checkedValues('region-filter');
          var addresses = checkedValues('address-filter');
          var visibleCount = 0;
          var included = [];

          if (directionNoteEl) {{
            directionNoteEl.textContent = 'Showing ' + directionLabels[currentDirection] +
              ' (' + (currentDirection === 'AM' ? 'Home → Work' : 'Work → Home') + ').';
          }}

          var useCluster = pinDisplayMode === 'cluster' && clusterGroup;
          if (clusterGroup) {{
            if (useCluster) {{
              if (!{map_var}.hasLayer(clusterGroup)) {{ clusterGroup.addTo({map_var}); }}
            }} else {{
              if ({map_var}.hasLayer(clusterGroup)) {{ {map_var}.removeLayer(clusterGroup); }}
            }}
          }}

          markerInfo.forEach(function(info) {{
            var filterMatch = regions.indexOf(info.region) !== -1 &&
                               addresses.indexOf(info.address) !== -1;
            var onMap = filterMatch && info.workerChecked;

            if (onMap && useCluster) {{
              if (info.addedDirect) {{ {map_var}.removeLayer(info.var); info.addedDirect = false; }}
              if (!info.addedToCluster) {{ clusterGroup.addLayer(info.var); info.addedToCluster = true; }}
            }} else if (onMap) {{
              if (info.addedToCluster) {{ clusterGroup.removeLayer(info.var); info.addedToCluster = false; }}
              if (!info.addedDirect) {{ info.var.addTo({map_var}); info.addedDirect = true; }}
            }} else {{
              if (info.addedDirect) {{ {map_var}.removeLayer(info.var); info.addedDirect = false; }}
              if (info.addedToCluster) {{ clusterGroup.removeLayer(info.var); info.addedToCluster = false; }}
            }}

            if (onMap) {{
              if (info.lineVar && !{map_var}.hasLayer(info.lineVar)) {{ info.lineVar.addTo({map_var}); }}
              visibleCount++;
            }} else {{
              if (info.lineVar && {map_var}.hasLayer(info.lineVar)) {{ {map_var}.removeLayer(info.lineVar); }}
            }}
            info.rowEl.style.display = filterMatch ? '' : 'none';
            info.labelSpanEl.textContent = info.postal + ' (' + formatMinutes(driveMinutesFor(info)) + ')';

            var minutes = driveMinutesFor(info);
            if (onMap && minutes !== null && minutes !== undefined) {{
              included.push(info);
            }}
          }});

          regionOrder.forEach(function(region) {{
            var anyVisible = (groupRowsEls[region] || []).some(function(row) {{
              return row.style.display !== 'none';
            }});
            var headerEl = groupHeaderEls[region];
            if (headerEl) {{ headerEl.style.display = anyVisible ? 'flex' : 'none'; }}
          }});

          addressOrder.forEach(function(address) {{
            var dest = destMarkers[address];
            if (!dest) {{ return; }}
            var show = addresses.indexOf(address) !== -1;
            if (show) {{
              if (!{map_var}.hasLayer(dest.var)) {{ dest.var.addTo({map_var}); }}
            }} else {{
              if ({map_var}.hasLayer(dest.var)) {{ {map_var}.removeLayer(dest.var); }}
            }}
          }});

          pinCountEl.textContent = visibleCount + ' of ' + markerInfo.length + ' shown';
          updateKpiPanel(included);
        }}

        document.querySelectorAll('.direction-toggle').forEach(function(radio) {{
          radio.addEventListener('change', function() {{
            if (radio.checked) {{
              currentDirection = radio.value;
              recomputeAll();
            }}
          }});
        }});

        document.querySelectorAll('.pin-display-toggle').forEach(function(radio) {{
          radio.addEventListener('change', function() {{
            if (radio.checked) {{
              pinDisplayMode = radio.value;
              recomputeAll();
            }}
          }});
        }});

        document.querySelectorAll('.region-filter, .address-filter').forEach(function(b) {{
          b.addEventListener('change', recomputeAll);
        }});

        function bindBulk(id, cls, checked) {{
          var el = document.getElementById(id);
          if (!el) return;
          el.addEventListener('click', function() {{
            document.querySelectorAll('.' + cls).forEach(function(b) {{ b.checked = checked; }});
            recomputeAll();
          }});
        }}
        bindBulk('region-select-all', 'region-filter', true);
        bindBulk('region-select-none', 'region-filter', false);
        bindBulk('address-select-all', 'address-filter', true);
        bindBulk('address-select-none', 'address-filter', false);

        function bindWorkerBulk(id, checked) {{
          var el = document.getElementById(id);
          if (!el) return;
          el.addEventListener('click', function() {{
            markerInfo.forEach(function(info) {{
              info.workerChecked = checked;
              info.checkboxEl.checked = checked;
            }});
            recomputeAll();
          }});
        }}
        bindWorkerBulk('worker-select-all', true);
        bindWorkerBulk('worker-select-none', false);

        recomputeAll();
      }}

      if (document.readyState === 'complete' || document.readyState === 'interactive') {{
        setTimeout(init, 200);
      }} else {{
        document.addEventListener('DOMContentLoaded', init);
      }}
    }})();
    """

    m.get_root().html.add_child(Element(control_html))
    m.get_root().script.add_child(Element(filter_js))


# ---------------------------------------------------------------------------
# Step 7: Package output into a zip for static hosting
# ---------------------------------------------------------------------------

_CDN_REF_PATTERN = re.compile(
    r'(?P<attr>src|href)="(?P<url>https://[^"]+\.(?:js|css))"'
)


def _localize_cdn_assets(html_text, output_dir):
    """Download the external Leaflet/jQuery/Bootstrap/Font Awesome files that
    folium links from CDNs and rewrite the HTML to reference local copies
    under assets/, so map.html doesn't depend on those CDNs being reachable
    when someone later opens it. Falls back to leaving the CDN URL in place
    (with a warning) for anything that fails to download."""
    assets_dir = os.path.join(output_dir, "assets")
    os.makedirs(assets_dir, exist_ok=True)

    seen = {}

    def replace(match):
        attr, url = match.group("attr"), match.group("url")
        if url not in seen:
            ext = os.path.splitext(url)[1]
            digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:10]
            local_name = f"{digest}{ext}"
            local_path = os.path.join(assets_dir, local_name)
            try:
                with urllib.request.urlopen(url, timeout=20) as resp:
                    data = resp.read()
                with open(local_path, "wb") as f:
                    f.write(data)
                seen[url] = f"assets/{local_name}"
            except Exception as exc:  # noqa: BLE001 - best-effort vendoring
                warnings.warn(f"Could not download {url} for offline use ({exc}); "
                               f"map.html will still reference it from the CDN.")
                seen[url] = url
        return f'{attr}="{seen[url]}"'

    return _CDN_REF_PATTERN.sub(replace, html_text)


def package_output(m, output_dir=OUTPUT_DIR, output_zip=OUTPUT_ZIP):
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    html_path = os.path.join(output_dir, OUTPUT_HTML_NAME)
    m.save(html_path)

    if VENDOR_LIBS_LOCALLY:
        with open(html_path, "r", encoding="utf-8") as f:
            html_text = f.read()
        html_text = _localize_cdn_assets(html_text, output_dir)
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_text)

    # Most static hosts (Netlify, GitHub Pages, S3, ...) serve index.html by
    # default at the site root. Without it, visiting the root URL shows the
    # host's own "not found" page even though map.html is deployed and fine
    # -- easy to mistake for "the map doesn't show". Ship a duplicate copy
    # named index.html so the site works whether visitors land on the root
    # URL or on /map.html directly.
    if OUTPUT_HTML_NAME != "index.html":
        shutil.copyfile(html_path, os.path.join(output_dir, "index.html"))

    readme_path = os.path.join(output_dir, "README.txt")
    with open(readme_path, "w") as f:
        f.write(
            "This folder contains the interactive map (map.html, duplicated as\n"
            "index.html) plus its supporting library files (assets/). Upload\n"
            "map.html, index.html, and the assets/ folder together to any static\n"
            "web host (GitHub Pages, S3 static site, Netlify, etc.), keeping\n"
            "them in the same relative layout.\n"
            "\n"
            "index.html is what most hosts serve automatically at your site's\n"
            "root URL (e.g. https://yoursite.netlify.app/) -- if you visit the\n"
            "root URL and only map.html exists, the host's own \"not found\"\n"
            "page shows instead of the map. Visiting /map.html directly always\n"
            "works too.\n"
            "\n"
            "The map itself doesn't need internet access to load (the Leaflet/\n"
            "jQuery/Bootstrap/Font Awesome files are bundled locally), but the\n"
            "background map tiles are still fetched live from OpenStreetMap, so\n"
            "whoever views the map needs internet access for those to appear.\n"
        )

    if os.path.exists(output_zip):
        os.remove(output_zip)
    with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(output_dir):
            for fname in files:
                full_path = os.path.join(root, fname)
                arcname = os.path.relpath(full_path, output_dir)
                zf.write(full_path, arcname=arcname)

    return output_zip


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def generate_map_from_excel(input_path, output_zip=OUTPUT_ZIP):
    df = load_spreadsheet(input_path)
    df = prepare_coordinates(df)
    df = classify_rows(df)

    api_key = _get_google_maps_api_key()
    df = compute_drive_times(df, api_key)

    m = build_map(df)
    zip_path = package_output(m, output_zip=output_zip)
    print(f"Done. {len(df)} pins plotted. Output written to: {zip_path}")
    return zip_path


def _running_in_colab():
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


def _running_in_notebook():
    try:
        from IPython import get_ipython

        return get_ipython() is not None
    except ImportError:
        return False


def _colab_main():
    """Convenience entry point for Google Colab: uploads the spreadsheet,
    runs the pipeline, and downloads the resulting zip."""
    from google.colab import files

    print("Upload your .xlsx spreadsheet...")
    uploaded = files.upload()
    input_path = next(iter(uploaded))

    zip_path = generate_map_from_excel(input_path)

    print("Downloading output zip...")
    files.download(zip_path)


def _main():
    if _running_in_colab():
        _colab_main()
        return

    if _running_in_notebook():
        # Plain Jupyter (not Colab): sys.argv holds kernel launch args, not a
        # file path, so ask directly instead.
        input_path = input("Path to your .xlsx spreadsheet: ").strip()
        if not input_path:
            raise SystemExit("No path provided.")
        generate_map_from_excel(input_path)
        return

    import sys

    if len(sys.argv) < 2:
        raise SystemExit("Usage: python generate_map.py <path-to-spreadsheet.xlsx>")
    generate_map_from_excel(sys.argv[1])


if __name__ == "__main__":
    _main()
