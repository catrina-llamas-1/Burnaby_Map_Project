"""
Postal Code Map Generator
==========================

Plots worker postal codes (British Columbia & Alberta) on an interactive
Leaflet/Folium map, with two independent-but-combinable checkbox filters:

  1) Region     -> classified from a keyword column into ClaimsPro / SCM / Pario
  2) Work Address - Line 1 -> one of three known office addresses

ClaimsPro and SCM share one color (blue). Pario gets its own color (purple).
Both filters can be used together (AND logic) or on their own.

Designed to run top-to-bottom as a single Google Colab cell (or as a normal
Python script). Everything you're likely to want to tweak lives in the
CONFIG block below.

Input:  an .xlsx spreadsheet with (at least) these columns:
          - Worker's Postal              (postal code used to place the pin)
          - Region                       (free text scanned for keywords)
          - Work Address - Line 1
          - City
          - Postal or ZIP code           (display-only postal/zip for the address)

Output: a .zip file containing map.html (and a small README), ready to
        upload to any static host (GitHub Pages, S3, Netlify, etc).
"""

import hashlib
import os
import re
import shutil
import urllib.request
import zipfile
import warnings
from html import escape

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG - edit these to match your spreadsheet / branding
# ---------------------------------------------------------------------------

# Column names as they appear in the spreadsheet.
COLUMN_GEOCODE_POSTAL = "Worker's Postal"          # used to place the pin on the map
COLUMN_REGION = "Region"                           # free-text, scanned for keywords below
COLUMN_ADDRESS_LINE1 = "Work Address - Line 1"     # matched against ADDRESS_CATEGORIES
COLUMN_CITY = "City"
COLUMN_DISPLAY_POSTAL = "Postal or ZIP code"        # shown in the popup, not used for geocoding

# Only keep rows whose postal code resolves to one of these provinces.
# Set to None to disable the province filter entirely.
ALLOWED_PROVINCES = {"BC", "AB"}

# Region classification: value -> list of keywords to search for (case-insensitive,
# substring match) inside COLUMN_REGION. First match wins; order matters if a
# row's text could match more than one.
REGION_KEYWORDS = {
    "ClaimsPro": ["claimspro"],
    "SCM": ["scm"],
    "Pario": ["pario"],
}

# Marker color per region. ClaimsPro & SCM intentionally share "blue".
REGION_COLORS = {
    "ClaimsPro": "blue",
    "SCM": "blue",
    "Pario": "purple",
}

# Work Address - Line 1 categories to filter on. Matching is case-insensitive
# and ignores extra whitespace, but otherwise looks for these as substrings
# of the spreadsheet's address value (so "Suite 112, 6093 Iona Drive, Burnaby"
# still matches "Suite 112, 6093 Iona Drive").
ADDRESS_CATEGORIES = [
    "8333 Eastlake Drive Suite 202",
    "1849 Welch Street",
    "Suite 112, 6093 Iona Drive",
]

# Label used for rows that don't match any known Region keyword / Address category.
UNKNOWN_REGION_LABEL = "Unclassified"
UNKNOWN_ADDRESS_LABEL = "Other / Unmatched Address"
UNKNOWN_COLOR = "gray"

MAP_TITLE = "Worker Postal Code Map"
MAP_START_LOCATION = [53.7267, -119.0]   # rough BC/AB midpoint
MAP_START_ZOOM = 5

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
# Step 2: Geocode postal codes (Canada, offline via pgeocode)
# ---------------------------------------------------------------------------

def geocode_postal_codes(df):
    """Add latitude/longitude/province columns based on COLUMN_GEOCODE_POSTAL."""
    try:
        import pgeocode
    except ModuleNotFoundError:
        import subprocess
        import sys

        print("pgeocode not found; installing it now...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "pgeocode"])
        import pgeocode

    nomi = pgeocode.Nominatim("ca")
    codes = df[COLUMN_GEOCODE_POSTAL].str.replace(" ", "", regex=False).str.upper()
    lookup = nomi.query_postal_code(codes.tolist())

    df["latitude"] = lookup["latitude"].values
    df["longitude"] = lookup["longitude"].values
    df["province_code"] = lookup["state_code"].values

    before = len(df)
    df = df.dropna(subset=["latitude", "longitude"]).copy()
    dropped = before - len(df)
    if dropped:
        warnings.warn(f"Dropped {dropped} row(s) with un-geocodable postal codes.")

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
# Step 4: Build the popup text
# ---------------------------------------------------------------------------

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
    return (
        f"<b>Region:</b> {escape(str(row['region_category']))}<br>"
        f"<b>Address:</b> {escape(address_line)}"
    )


# ---------------------------------------------------------------------------
# Step 5: Build the Folium map with combinable Region / Address checkboxes
# ---------------------------------------------------------------------------

def build_map(df):
    import folium
    from folium import Element

    m = folium.Map(location=MAP_START_LOCATION, zoom_start=MAP_START_ZOOM, tiles="OpenStreetMap")
    m.get_root().html.add_child(Element(f"<title>{escape(MAP_TITLE)}</title>"))

    region_labels = list(REGION_KEYWORDS.keys())
    if (df["region_category"] == UNKNOWN_REGION_LABEL).any():
        region_labels.append(UNKNOWN_REGION_LABEL)

    address_labels = list(ADDRESS_CATEGORIES)
    if (df["address_category"] == UNKNOWN_ADDRESS_LABEL).any():
        address_labels.append(UNKNOWN_ADDRESS_LABEL)

    marker_meta = []  # [{"var": js_var_name, "region": ..., "address": ...}, ...]

    for _, row in df.iterrows():
        marker = folium.Marker(
            location=[row["latitude"], row["longitude"]],
            popup=folium.Popup(build_popup_html(row), max_width=320),
            icon=folium.Icon(color=row["marker_color"], icon="user", prefix="fa"),
        )
        marker.add_to(m)
        marker_meta.append(
            {
                "var": marker.get_name(),
                "region": row["region_category"],
                "address": row["address_category"],
            }
        )

    _add_filter_control(m, marker_meta, region_labels, address_labels)
    return m


def _add_filter_control(m, marker_meta, region_labels, address_labels):
    """Inject a custom Leaflet control with two checkbox groups (Region, Work
    Address). A marker is shown only if its region AND its address are both
    checked, so the two filters combine (AND) or can be used on their own by
    leaving every box in the other group checked."""
    from folium import Element

    map_var = m.get_name()

    def checkbox_html(group, values):
        items = []
        for v in values:
            safe_id = f"{group}_{re.sub(r'[^a-zA-Z0-9]', '_', v)}"
            items.append(
                f'<label style="display:block;font-weight:normal;margin:2px 0;">'
                f'<input type="checkbox" class="{group}-filter" value="{escape(v)}" '
                f'id="{safe_id}" checked> {escape(v)}</label>'
            )
        return "\n".join(items)

    control_html = f"""
    <div id="postal-map-filter-panel" style="
        position: fixed; top: 10px; right: 10px; z-index: 9999;
        background: white; padding: 10px 14px; border: 2px solid #444;
        border-radius: 6px; font-family: Arial, sans-serif; font-size: 13px;
        max-height: 90vh; overflow-y: auto; box-shadow: 2px 2px 6px rgba(0,0,0,0.3);">
      <div style="font-weight:bold; margin-bottom:6px;">{escape(MAP_TITLE)}</div>

      <div style="font-weight:bold; margin-top:6px;">Region</div>
      {checkbox_html("region", region_labels)}
      <button id="region-select-all" style="margin-top:4px;">All</button>
      <button id="region-select-none">None</button>

      <div style="font-weight:bold; margin-top:10px;">Work Address - Line 1</div>
      {checkbox_html("address", address_labels)}
      <button id="address-select-all" style="margin-top:4px;">All</button>
      <button id="address-select-none">None</button>
    </div>
    """

    marker_meta_js = ",\n".join(
        '    {{var: {var}, region: "{region}", address: "{address}"}}'.format(
            var=meta["var"],
            region=meta["region"].replace('"', '\\"'),
            address=meta["address"].replace('"', '\\"'),
        )
        for meta in marker_meta
    )

    filter_js = f"""
    <script>
    (function() {{
      function init() {{
        var markerInfo = [
{marker_meta_js}
        ];

        function checkedValues(cls) {{
          var boxes = document.querySelectorAll('.' + cls);
          var vals = [];
          boxes.forEach(function(b) {{ if (b.checked) vals.push(b.value); }});
          return vals;
        }}

        function applyFilter() {{
          var regions = checkedValues('region-filter');
          var addresses = checkedValues('address-filter');
          markerInfo.forEach(function(info) {{
            var show = regions.indexOf(info.region) !== -1 &&
                       addresses.indexOf(info.address) !== -1;
            if (show) {{
              if (!{map_var}.hasLayer(info.var)) {{ info.var.addTo({map_var}); }}
            }} else {{
              if ({map_var}.hasLayer(info.var)) {{ {map_var}.removeLayer(info.var); }}
            }}
          }});
        }}

        document.querySelectorAll('.region-filter, .address-filter').forEach(function(b) {{
          b.addEventListener('change', applyFilter);
        }});

        function bindBulk(id, cls, checked) {{
          var el = document.getElementById(id);
          if (!el) return;
          el.addEventListener('click', function() {{
            document.querySelectorAll('.' + cls).forEach(function(b) {{ b.checked = checked; }});
            applyFilter();
          }});
        }}
        bindBulk('region-select-all', 'region-filter', true);
        bindBulk('region-select-none', 'region-filter', false);
        bindBulk('address-select-all', 'address-filter', true);
        bindBulk('address-select-none', 'address-filter', false);
      }}

      if (document.readyState === 'complete' || document.readyState === 'interactive') {{
        setTimeout(init, 200);
      }} else {{
        document.addEventListener('DOMContentLoaded', init);
      }}
    }})();
    </script>
    """

    m.get_root().html.add_child(Element(control_html))
    m.get_root().script.add_child(Element(filter_js))


# ---------------------------------------------------------------------------
# Step 6: Package output into a zip for static hosting
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
    df = geocode_postal_codes(df)
    df = classify_rows(df)
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
