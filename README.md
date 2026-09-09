# Worker Postal Code Map

Generates an interactive HTML map of worker postal codes (BC & Alberta),
filterable by **Region** (ClaimsPro / SCM / Pario) and **Work Address - Line 1**,
usable together or independently. The Region filter shows a color legend
next to each checkbox. Output is a zip you can host as a static site.

Pins are placed using the `Lat`/`Long` coordinates already present in the
spreadsheet — no geocoding needed to plot them.

## Drive times, straight-line distance & KPIs

Each worker's drive time and distance to their classified Work Address is
computed with the **Google Maps Distance Matrix API**, for a departure time
you choose: **7:30 AM** or **4:30 PM** (modeled for next Monday, in each
worker's own province timezone, so results reflect typical weekday traffic
rather than a specific date). A straight line is drawn on the map between
each worker's pin and their work address.

- **Sidebar** — lists every worker by `Worker's Postal` (with their drive
  time), grouped under color-coded Region headings, each with its own
  checkbox to include/exclude that worker. Unchecking a worker removes
  their pin and line from the map and from the KPI panel, without removing
  them from the list. Clicking a worker's postal code zooms/pans to it.
- **KPI panel** — recomputes live from whichever workers are currently
  included (Region + Work Address filters, and individual checkboxes):
  base size, average drive time, median drive time, % of workers under 15
  / 20 / 30 minutes, average distance (km), and the closest/farthest
  worker by drive time.

## Google Maps API key (required for drive times)

You need a Google Cloud project with the **Geocoding API** and **Distance
Matrix API** both enabled and billing set up, and an API key from it (the
Geocoding API is used once per distinct work address; results are cached).

Supply the key one of three ways (checked in this order):
1. Paste it into `GOOGLE_MAPS_API_KEY` in the `CONFIG` block at the top of
   `generate_map.py`.
2. Set a `GOOGLE_MAPS_API_KEY` environment variable before running.
3. Leave both blank and run interactively (Colab/Jupyter) — you'll be
   prompted for it in a text box (masked input), and separately asked to
   choose 7:30 AM or 4:30 PM for the departure time.

If the key is missing, invalid, or either API/billing isn't enabled, the
script fails immediately with a clear error rather than silently producing
a map with no drive times.

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
   your Google Maps API key and departure time (if not already configured)
   and to upload your `.xlsx` file, then automatically download
   `postal_code_map_output.zip` when done.

## Run locally

```
pip install -r requirements.txt
python generate_map.py path/to/spreadsheet.xlsx
```

## Required spreadsheet columns

- `Worker's Postal` — postal code, shown in the popup/sidebar (its first
  letter is also used to determine province for the BC/AB filter and for
  picking the right timezone for drive-time calculations)
- `Lat`, `Long` — coordinates used to place the pin directly
- `Region` — free text; scanned for keywords (`ClaimsPro`, `SCM`, `Pario`)
- `Work Address - Line 1`
- `City`
- `Postal or ZIP code` — display-only, shown in the popup

## Customizing

All configuration (column names, region keywords, colors, address
categories, departure time options, drive-time KPI thresholds, map
title/start location) lives in the `CONFIG` block at the top of
`generate_map.py`.

## Output

`postal_code_map_output.zip` containing `map.html` (duplicated as
`index.html`, since most static hosts serve that at the site root) and an
`assets/` folder with Leaflet/jQuery/Bootstrap/Font Awesome bundled locally,
so the map doesn't depend on any CDN being reachable when someone opens it
later. Upload `map.html`, `index.html`, and `assets/` together, keeping
their relative layout, to any static host (GitHub Pages, S3, Netlify, etc.).
Background map tiles are still fetched live from OpenStreetMap, so viewers
need internet access for those to load — same as any web map. Drive times
and distances are computed once when you run the script and baked into the
page; the deployed map does not call the Google Maps API itself.
