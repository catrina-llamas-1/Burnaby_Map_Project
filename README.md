# Worker Postal Code Map

Generates an interactive HTML map of worker postal codes (BC & Alberta),
filterable by **Region** (ClaimsPro / SCM / Pario) and **Work Address - Line 1**,
usable together or independently. The Region filter shows a color legend
next to each checkbox. Output is a static site, ready to deploy to
Firebase Hosting (or any other static host).

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

## Background map tiles

Background map tiles come from OpenStreetMap's own tile servers, which
are volunteer-run and require every request to properly identify itself
(see [OSM's tile usage policy](https://operations.osmfoundation.org/policies/tiles/)).
A normal browser viewing the exported map from a real hosted URL already
sends a `Referer` header identifying the page, which satisfies this — but
opening `map.html` as a local `file://` path sends no `Referer` and can
get flagged as unidentified/bulk traffic, resulting in "Access blocked"
tiles.

**Always view/test the exported map via a real URL** — either the static
host it's deployed to, or a local static server for testing
(`python -m http.server` from inside the unzipped output, then visit
`http://localhost:8000/`) — never by double-clicking `map.html` directly.

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

`postal_code_map_output/` (and the equivalent `postal_code_map_output.zip`)
containing `map.html` (duplicated as `index.html`, since most static hosts
serve that at the site root) and an `assets/` folder with Leaflet/jQuery/
Bootstrap/Font Awesome bundled locally, so the map doesn't depend on any
CDN being reachable when someone opens it later. Background map tiles are
still fetched live from OpenStreetMap, so viewers need internet access for
those to load (see "Background map tiles" above for the Referer/hosting
caveat). Drive times and distances are computed once when you run the
script and baked into the page; the deployed map does not call the Google
Maps API itself.

Both are gitignored by default, since the output can contain real
employee postal codes/addresses — don't commit it to a public repo.

## Deploying to Firebase Hosting

This mirrors the setup used by `catrina-llamas-1/Companies-amenities-map`
(same Firebase project, same deploy pattern): upload a spreadsheet to the
repo, push, and GitHub Actions builds and deploys automatically. No
stored service-account key — auth uses Workload Identity Federation.

- `firebase.json` points Firebase Hosting's public directory at
  `postal_code_map_output/`, the generator's default output folder.
- `.firebaserc` links this repo to the `stony-plain-rd-companies-map`
  Firebase project (shared with Companies-amenities-map), but deploys to
  its own Hosting **site**, `commuter-analysis-map-demo`, via the
  `burnaby-map` target — not that project's default site, which
  Companies-amenities-map already occupies. Live at
  <https://commuter-analysis-map-demo.web.app>.
- `data/` holds the one input spreadsheet — see `data/README.md`. Upload
  a new one there (GitHub web UI: **Add file → Upload files**) to update
  the live map.
- `build_for_deploy.py` finds that spreadsheet and runs `generate_map.py`
  against it, producing `postal_code_map_output/`.

### Updating the live map

Just replace the spreadsheet in `data/` and push/merge to
`feature/firebase-hosting` (or open a PR touching `data/` for a 7-day
preview link first). GitHub Actions does the rest. The workflows only
trigger on changes under `data/`, not every push — each build calls the
paid Google Geocoding/Distance Matrix APIs, so unrelated pushes (docs,
code changes) don't burn API calls.

### One-time setup (already done once for this project; redo per-repo)

1. **GCP Workload Identity Federation**: run
   `scripts/setup_github_deploy.sh` once, in [Google Cloud
   Shell](https://shell.cloud.google.com), signed in as an owner of the
   `stony-plain-rd-companies-map` Firebase project. It reuses that
   project's existing `github-deploy` service account (created for
   Companies-amenities-map) and adds a second OIDC provider trusting only
   this repo, so it doesn't touch or collide with that repo's existing
   access. It prints two values at the end.
2. Add those as **repository variables** (not secrets) in
   `catrina-llamas-1/Burnaby_Map_Project` → Settings → Secrets and
   variables → Actions → Variables tab:
   - `GCP_WORKLOAD_IDENTITY_PROVIDER`
   - `GCP_SERVICE_ACCOUNT`
3. Add a **repository secret** (Secrets tab — this one's sensitive) named
   `GOOGLE_MAPS_API_KEY`: a Google Maps Platform key with the Geocoding
   API and Distance Matrix API enabled and billing set up. Unlike
   Companies-amenities-map, this project computes drive times and needs
   its own key.

Until both are set, pushes to `data/` still build the map (so you can
verify `build_for_deploy.py` works) but skip the deploy step with a
warning — see the "Check Firebase setup" step in
`.github/workflows/firebase-hosting-merge.yml`.

### Manual deploy (no GitHub Actions)

1. Run `generate_map.py` to produce `postal_code_map_output/` (or
   `python build_for_deploy.py` to build from whatever's in `data/`).
2. Install the Firebase CLI if you haven't: `npm install -g firebase-tools`.
3. From the repo root: `firebase login`, then `firebase deploy`.

### Other static hosts

The output is still a portable static site, so any other static host
(GitHub Pages, S3, Netlify, etc.) works too if you'd rather use one of
those instead — just upload `map.html`, `index.html`, and `assets/`
together, keeping their relative layout.
