# Worker Postal Code Map

Generates an interactive HTML map of worker postal codes (BC & Alberta),
filterable by **Region** (ClaimsPro / SCM / Pario) and **Work Address - Line 1**,
usable together or independently. Output is a zip you can host as a static site.

## Run in Google Colab

1. Upload `generate_map.py` to your Colab session (or paste its contents into a cell).
2. Install dependencies in a cell:
   ```
   !pip install -q pandas openpyxl pgeocode folium
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

`postal_code_map_output.zip` containing `map.html` and an `assets/` folder
with its Leaflet/jQuery/Bootstrap/Font Awesome files bundled locally, so the
map doesn't depend on any CDN being reachable when someone opens it later.
Upload both `map.html` and `assets/` together, keeping their relative layout,
to any static host (GitHub Pages, S3, Netlify, etc.). Background map tiles
are still fetched live from OpenStreetMap, so viewers need internet access
for those to load — same as any web map.
