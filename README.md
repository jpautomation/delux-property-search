# Delux Property Search

A Python tool that reads Rightmove alert emails, scores each property against weighted criteria, and generates an interactive map with adjustable score weights and filters.

## What it does

- Reads Rightmove alert emails via IMAP (Gmail)
- Scrapes each listing for address, price, bedrooms, floor area, land area, description, key features, and SSTC status
- Scores each property 0–100 across seven criteria:
  - **Tranquillity** — ONS Rural-Urban Classification (via postcodes.io)
  - **Property type** — configurable preference scores
  - **Flood risk** — Environment Agency flood zone API
  - **Price** — relative to your maximum budget
  - **Bedrooms** — relative to your target
  - **Broadband** — Ofcom Connected Nations gigabit/FTTP/SFBB availability
  - **Drive time** — OpenRouteService routing API to a reference location
- Extracts metadata flags from listing descriptions (views, stone built, garage, log burner, Aga, annexe, woodland, land, holiday potential, etc.)
- Looks up floor area from the EPC register
- Deduplicates across runs; detects price changes
- Saves results to `results.csv`
- Generates `map.html` — an interactive Leaflet map with:
  - Adjustable score weight sliders (live recolour)
  - Absolute filters (max price, min bedrooms, min floor area, min land)

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

`rasterio` is optional (only needed if you have a CPRE tranquillity raster file):
```bash
conda install rasterio
```

### 2. Configure credentials

```bash
cp config.example.json config.json
```

Edit `config.json`:

| Key | What to put |
|-----|-------------|
| `email.username` | Your Gmail address |
| `email.app_password` | Gmail app password — [create one here](https://myaccount.google.com/apppasswords) (requires 2FA enabled) |
| `email.imap_server` | `imap.gmail.com` (or your provider) |
| `openrouteservice_api_key` | Free key from [openrouteservice.org](https://openrouteservice.org) |
| `epc_api_key` | Bearer token from [get-energy-performance-data.communities.gov.uk](https://get-energy-performance-data.communities.gov.uk) → sign in → My account |
| `rightmove_sender` | The email address you forward Rightmove alerts from |
| `thresholds.reference_lat/lon` | Coordinates of your reference location (e.g. workplace) for drive time scoring |

### 3. Download Ofcom broadband data

Download the fixed-line postcode-level coverage files from the [Ofcom Connected Nations](https://www.ofcom.org.uk/research-and-data/telecoms-research/connected-nations) data page. Place the CSV files in a folder called `postcode_files/` next to the script.

Files are named `202507_fixed_pc_coverage_r01_{AREA}.csv` where `{AREA}` is the postcode area (e.g. `CW`, `SK`, `ST`). Download only the areas covering your search region.

### 4. Set up Rightmove email alerts

- Go to [rightmove.co.uk](https://rightmove.co.uk), run a search, and save it as an email alert
- Forward the alert emails to the Gmail account you configured

### 5. Add manual properties (optional)

Paste Rightmove URLs into `manual_urls.txt`, one per line:
```
https://www.rightmove.co.uk/properties/123456789
```

## Run

```bash
python property_scorer.py
```

The script records the date of each run in `last_fetch_date.txt`. Subsequent runs fetch only emails received since that date, regardless of whether those emails have been read elsewhere.

On the first run (no stored date), the last 90 days are fetched. To go further back:

```bash
python property_scorer.py --since 2026-01-01
```

Results are saved to `results.csv`. Open `map.html` in a browser for the interactive map.

## Scoring weights

Adjust the `weights` section of `config.json` to reflect your priorities. Values are relative — the script normalises them so they don't need to sum to 1.

## Data sources

| Source | Used for |
|--------|----------|
| [postcodes.io](https://postcodes.io) | Geocoding, ONS Rural-Urban Classification |
| [Nominatim / OpenStreetMap](https://nominatim.org) | Address-level geocoding |
| [Environment Agency Flood API](https://environment.data.gov.uk) | Flood zone scoring |
| [Ofcom Connected Nations](https://www.ofcom.org.uk/research-and-data/telecoms-research/connected-nations) | Broadband availability |
| [OpenRouteService](https://openrouteservice.org) | Driving time |
| [EPC Register](https://get-energy-performance-data.communities.gov.uk) | Floor area |
| [Rightmove](https://rightmove.co.uk) | Property listings (via your own alert emails) |
