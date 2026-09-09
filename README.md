# Worker Postal Code Map

Generates an interactive HTML map of worker postal codes (BC & Alberta),
filterable by **Region** (ClaimsPro / SCM / Pario) and **Work Address - Line 1**,
usable together or independently. The Region filter shows a color legend
next to each checkbox. A left-hand sidebar lists every pin by its
`Worker's Postal` value, grouped under the same color-coded Region headings,
stays in sync with the active filters, and clicking an entry zooms/pans the
map to that pin and opens its popup. Output is a zip you can host as a
static site.

Pins are placed using the `Lat`/`Long` coordinates already present in the
spreadsheet — no geocoding step, no API key needed.

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
   or, if pasted into a cell, just run the cell — it will prompt you to
   upload your `.xlsx` file, then automatically download
   `postal_code_map_output.zip` when done.

## Run locally

```
pip install -r requirements.txt
python generate_map.py path/to/spreadsheet.xlsx
```

## Required spreadsheet columns

- `Worker's Postal` — postal code, shown in the popup/sidebar (its first
  letter is also used to determine province for the BC/AB filter)
- `Lat`, `Long` — coordinates used to place the pin directly
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
