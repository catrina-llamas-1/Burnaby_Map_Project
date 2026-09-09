# Worker Postal Code Map

Generates an interactive HTML map of worker postal codes (BC & Alberta),
filterable by **Region** (ClaimsPro / SCM / Pario) and **Work Address - Line 1**,
usable together or independently. The Region filter shows a color legend
next to each checkbox (blue for ClaimsPro/SCM, purple for Pario). A
left-hand sidebar lists every pin by its `Worker's Postal` value, grouped
under the same color-coded Region headings, stays in sync with the active
filters, and clicking an entry zooms/pans the map to that pin and opens its
popup. Output is a zip you can host as a static site.

## Google Maps API key (required for geocoding)

Postal codes are geocoded with the **Google Maps Geocoding API** (full
postal-code precision, not just a 3-character area centroid). You need a
Google Cloud project with the **Geocoding API** enabled and billing set up,
and an API key from it. Each *unique* postal code in your spreadsheet is one
billed request — duplicates are cached and don't cost extra.

Supply the key one of three ways (checked in this order):
1. Paste it into `GOOGLE_MAPS_API_KEY` in the `CONFIG` block at the top of
   `generate_map.py`.
2. Set a `GOOGLE_MAPS_API_KEY` environment variable before running.
3. Leave both blank and run interactively (Colab/Jupyter) — you'll be
   prompted for it with a masked input, so nothing sensitive ends up saved
   in the notebook.

If the key is missing, invalid, or the Geocoding API/billing isn't enabled,
the script fails immediately with a clear error instead of silently
producing a map with no pins.

## Run in Google Colab

1. Upload `generate_map.py` to your Colab session (or paste its contents into a cell).
2. Install dependencies in a cell:
   ```
   !pip install -q pandas openpyxl folium
   ```
3. Run the script:
   ```
   !python generate_map.py
   ```
   or, if pasted into a cell, just run the cell — it will prompt you for
   your Google Maps API key (if not already configured) and to upload your
   `.xlsx` file, then automatically download `postal_code_map_output.zip`
   when done.

## Run locally

```
pip install -r requirements.txt
python generate_map.py path/to/spreadsheet.xlsx
```

## Required spreadsheet columns

- `Worker's Postal` — postal code used to place the pin
- `Region` — free text; scanned for keywords (`ClaimsPro`, `SCM`, `Pario`)
- `Work Address - Line 1`
- `City`
- `Postal or ZIP code` — display-only, shown in the popup

## Customizing

All configuration (column names, region keywords, colors, address categories,
map title/start location) lives in the `CONFIG` block at the top of
`generate_map.py`.

## Output

`postal_code_map_output.zip` containing `map.html` (duplicated as
`index.html`, since most static hosts serve that at the site root) and an
`assets/` folder with Leaflet/jQuery/Bootstrap/Font Awesome bundled locally,
so the map doesn't depend on any CDN being reachable when someone opens it
later. Upload `map.html`, `index.html`, and `assets/` together, keeping
their relative layout, to any static host (GitHub Pages, S3, Netlify, etc.).
Background map tiles are still fetched live from OpenStreetMap, so viewers
need internet access for those to load — same as any web map.
