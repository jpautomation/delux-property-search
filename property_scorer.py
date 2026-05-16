#!/usr/bin/env python3
"""
Delux Property Search
Reads Rightmove alert emails and scores each property against weighted
criteria, producing a ranked list and an interactive map.

Setup
-----
1. Install dependencies:
       pip install imap-tools requests beautifulsoup4 numpy pandas tabulate pyproj
       conda install rasterio        (easiest on Windows)

2. Enable IMAP in Outlook:
       outlook.com → Settings → Mail → Sync email → Enable IMAP

3. Create an app password (if MFA is on):
       account.microsoft.com → Security → Advanced security → App passwords

4. Download data files (place in same folder as this script):

   Ofcom Broadband data (postcode-level, split by area)
       https://www.ofcom.org.uk/siteassets/resources/documents/research-and-data/multi-sector/infrastructure-research/connected-nations-2025/202507_fixed_pc_coverage_r01.zip
       Extract the zip — produces one CSV per postcode area (CW, SK, SY, etc.).
       Place the extracted folder as: postcode_files/  (next to this script)

   OpenRouteService API key (free)
       https://openrouteservice.org/dev/#/signup
       Add key to CONFIG below.
"""

import os
import re
import json
import time
import logging
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup
from imap_tools import MailBox, AND
from tabulate import tabulate

try:
    import rasterio
    from pyproj import Transformer
    RASTERIO_AVAILABLE = True
except ImportError:
    RASTERIO_AVAILABLE = False

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────────────────────────────────────

_config_path = os.path.join(os.path.dirname(__file__), "config.json")
if not os.path.exists(_config_path):
    raise FileNotFoundError(
        "config.json not found. Copy config.example.json to config.json and fill in your credentials."
    )
with open(_config_path, encoding="utf-8") as _f:
    CONFIG = json.load(_f)

# ──────────────────────────────────────────────────────────────────────────────
# DATA CLASSES
# ──────────────────────────────────────────────────────────────────────────────

@dataclass
class Property:
    url: str
    address: str = ""
    price: Optional[int] = None
    bedrooms: Optional[int] = None
    property_type: str = ""
    postcode: str = ""
    lat: Optional[float] = None
    lon: Optional[float] = None
    rural_urban: str = ""
    status: str = ""
    floor_area: Optional[float] = None
    land_acres: Optional[float] = None
    key_features: list = field(default_factory=list)
    description: str = ""
    flag_views: bool = False
    flag_stone_built: bool = False
    flag_garage: bool = False
    flag_outbuildings: bool = False
    flag_log_burner: bool = False
    flag_aga: bool = False
    flag_no_chain: bool = False
    flag_annexe: bool = False
    flag_woodland: bool = False
    flag_planning: bool = False
    flag_period_character: bool = False
    flag_paddock: bool = False
    flag_land: bool = False
    flag_holiday: bool = False
    scores: dict = field(default_factory=dict)
    total_score: float = 0.0
    notes: list = field(default_factory=list)

# ──────────────────────────────────────────────────────────────────────────────
# MANUAL PROPERTIES
# ──────────────────────────────────────────────────────────────────────────────

def load_manual_properties(config: dict) -> list:
    path = os.path.join(os.path.dirname(__file__), "manual_urls.txt")
    if not os.path.exists(path):
        return []
    properties = []
    seen_ids = set()
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            m = re.search(r"rightmove\.co\.uk/properties/(\d+)", line)
            if not m:
                logging.warning(f"manual_urls.txt: skipping unrecognised line: {line}")
                continue
            prop_id = m.group(1)
            if prop_id in seen_ids:
                continue
            seen_ids.add(prop_id)
            url = f"https://www.rightmove.co.uk/properties/{prop_id}"
            properties.append(Property(url=url))
    if properties:
        print(f"Loaded {len(properties)} manual properties from manual_urls.txt")
    return properties

# EMAIL READER
# ──────────────────────────────────────────────────────────────────────────────

def fetch_rightmove_emails(config: dict) -> list:
    properties = []
    seen_urls = set()
    email_cfg = config["email"]

    with MailBox(email_cfg["imap_server"]).login(
        email_cfg["username"], email_cfg["app_password"]
    ) as mailbox:
        subject = config.get("rightmove_subject", "")
        if config.get("fetch_all_email"):
            criteria = AND(from_=config["rightmove_sender"], subject=subject) if subject else AND(from_=config["rightmove_sender"])
        else:
            from datetime import date, timedelta
            criteria = AND(from_=config["rightmove_sender"], subject=subject,
                           date_gte=date.today() - timedelta(days=30)) if subject else AND(
                from_=config["rightmove_sender"], date_gte=date.today() - timedelta(days=30))

        for msg in mailbox.fetch(criteria):
            html = msg.html or ""
            soup = BeautifulSoup(html, "html.parser")

            for a in soup.find_all("a", href=True):
                href = a["href"]
                match = re.search(r"rightmove\.co\.uk/properties/(\d+)", href)
                if not match:
                    continue
                prop_id = match.group(1)
                canonical_url = f"https://www.rightmove.co.uk/properties/{prop_id}"
                if canonical_url in seen_urls:
                    continue
                seen_urls.add(canonical_url)

                price, beds, ptype, address = _extract_from_email_context(a, soup)
                properties.append(Property(
                    url=canonical_url,
                    price=price,
                    bedrooms=beds,
                    property_type=ptype,
                    address=address,
                ))

    logging.info(f"Found {len(properties)} unique properties in emails")
    return properties


def _extract_from_email_context(link_tag, soup) -> tuple:
    price = beds = None
    ptype = address = ""

    # Walk up the DOM until we find a node with enough text to be the property card
    node = link_tag
    text = ""
    for _ in range(12):
        parent = node.find_parent(["td", "tr", "div", "table"])
        if not parent:
            break
        candidate = parent.get_text(" ", strip=True)
        if len(candidate) > 80 and "for sale" in candidate.lower():
            text = candidate
            break
        node = parent

    if not text:
        return price, beds, ptype, address

    m = re.search(r"£([\d,]+)", text)
    if m:
        price = int(m.group(1).replace(",", ""))

    m = re.search(r"(\d+)\s*bedroom", text, re.IGNORECASE)
    if m:
        beds = int(m.group(1))

    type_keywords = [
        "farmhouse", "barn conversion", "detached", "semi-detached",
        "end of terrace", "terraced", "flat", "apartment", "bungalow",
        "cottage", "character property", "maisonette",
    ]
    for t in type_keywords:
        if t.lower() in text.lower():
            ptype = t
            break

    # Address sits between "for sale" and "Marketed by" (or end of string)
    m = re.search(r"for sale\s+(.+?)(?:\s+Marketed by|\s*$)", text, re.IGNORECASE)
    if m:
        address = m.group(1).strip()[:100]

    return price, beds, ptype, address

# ──────────────────────────────────────────────────────────────────────────────
# RIGHTMOVE LISTING ENRICHMENT (postcode + fill missing fields)
# ──────────────────────────────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    )
}

def enrich_from_listing(prop: Property) -> None:
    try:
        resp = requests.get(prop.url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
    except Exception as e:
        prop.notes.append(f"Listing fetch failed: {e}")
        return

    soup = BeautifulSoup(resp.text, "html.parser")

    # JSON-LD structured data (most reliable)
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            addr = data.get("address", {})
            pc = addr.get("postalCode", "")
            if pc:
                prop.postcode = pc.strip().upper()
            if not prop.address:
                prop.address = addr.get("streetAddress", "")
            break
        except (json.JSONDecodeError, AttributeError):
            continue

    # Fallback: address from <h1> or page title
    if not prop.address:
        h1 = soup.find("h1")
        if h1:
            prop.address = h1.get_text(strip=True)
        elif soup.title:
            m = re.search(r"for sale in (.+?)(?:\s*\||\s*for £|\s*$)", soup.title.string or "", re.IGNORECASE)
            if m:
                prop.address = m.group(1).strip()

    # Fallback: regex postcode scan — pick least-common real postcode
    # (nav/search context repeats a different postcode many times; property postcode appears once)
    if not prop.postcode:
        from collections import Counter
        candidates = re.findall(r"\b([A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2})\b", resp.text)
        real = [p for p in candidates if not re.fullmatch(r"[A-Z]\d[A-Z]\s*\d[A-Z][A-Z]", p)]
        if real:
            prop.postcode = min(Counter(real), key=lambda p: Counter(real)[p]).strip().upper()

    page_text = soup.get_text(" ")

    if not prop.bedrooms:
        m = re.search(r"(\d+)\s*bed", page_text, re.IGNORECASE)
        if m:
            prop.bedrooms = int(m.group(1))

    if not prop.property_type:
        for t in [
            "farmhouse", "barn conversion", "detached", "semi-detached",
            "end of terrace", "terraced", "flat", "bungalow", "cottage",
        ]:
            if t.lower() in page_text.lower():
                prop.property_type = t
                break

    if not prop.price:
        m = re.search(r"£([\d,]+)", page_text)
        if m:
            prop.price = int(m.group(1).replace(",", ""))

    # Floor area from listing (sq ft or m²)
    if prop.floor_area is None:
        m = re.search(r"([\d,]+)\s*sq\.?\s*ft", page_text, re.IGNORECASE)
        if m:
            prop.floor_area = round(int(m.group(1).replace(",", "")) * 0.0929, 1)  # convert to m²
        else:
            m = re.search(r"([\d,]+)\s*m²", page_text)
            if m:
                prop.floor_area = float(m.group(1).replace(",", ""))

    # Land area (acres or hectares → stored as acres)
    _WORD_NUMS = {
        "half": 0.5, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
        "eleven": 11, "twelve": 12, "fifteen": 15, "twenty": 20,
    }
    if prop.land_acres is None:
        m = re.search(r"([\d.]+)\s*acres?", page_text, re.IGNORECASE)
        if m:
            prop.land_acres = round(float(m.group(1)), 2)
        else:
            m = re.search(r"([\d.]+)\s*hectares?", page_text, re.IGNORECASE)
            if m:
                prop.land_acres = round(float(m.group(1)) * 2.471, 2)
        if prop.land_acres is None:
            m = re.search(
                r"(" + "|".join(_WORD_NUMS) + r")\s+acres?",
                page_text, re.IGNORECASE
            )
            if m:
                prop.land_acres = _WORD_NUMS[m.group(1).lower()]

    # Key features
    if not prop.key_features:
        for heading in soup.find_all(["h2", "h3", "h4", "strong"]):
            if "key feature" in heading.get_text(strip=True).lower():
                ul = heading.find_next(["ul", "ol"])
                if ul:
                    prop.key_features = [li.get_text(strip=True) for li in ul.find_all("li")]
                break

    # Full description text
    if not prop.description:
        for heading in soup.find_all(["h2", "h3", "h4"]):
            if heading.get_text(strip=True).lower() == "description":
                nxt = heading.find_next_sibling()
                if nxt:
                    prop.description = re.sub(r"\s+", " ", nxt.get_text(" ", strip=True))
                break

    # Metadata flags from description + key features
    _search_text = (prop.description + " " + " ".join(prop.key_features)).lower()
    _FLAG_PATTERNS = {
        "views":           ["view", "panoramic", "far-reaching", "far reaching", "outlook", "vista"],
        "stone_built":     ["stone built", "stone-built", "stone cottage", "stone farmhouse", "stone property", "stone construction"],
        "garage":          ["garage", "carport"],
        "outbuildings":    ["outbuilding", "workshop", "stable", "barn", "store room", "store building"],
        "log_burner":      ["log burner", "log-burner", "wood burner", "woodburner", "wood burning stove", "wood-burning"],
        "aga":             ["aga", "rayburn", "range cooker"],
        "no_chain":        ["no chain", "no onward chain", "chain free", "chain-free"],
        "annexe":          ["annexe", "annex", "granny flat", "holiday let", "holiday barn", "self-contained", "self contained"],
        "woodland":        ["woodland", "copse", "orchard", "mature trees"],
        "planning":        ["planning permission", "subject to planning", "development potential", "development opportunity", "planning consent"],
        "period_character":["period ", "listed building", "grade ii", "grade 2", "heritage", "historic", "victorian", "georgian", "edwardian"],
        "paddock":         ["paddock", "equestrian", "horses", "stabling", "stables"],
        "land":            ["acres", "hectares", "paddock", "grassland", "farmland", "smallholding", "plot of land", "land extending"],
        "holiday":         ["holiday let", "holiday cottage", "holiday barn", "holiday rental", "holiday home", "tourist", "airbnb", "income potential", "letting potential"],
    }
    for flag, patterns in _FLAG_PATTERNS.items():
        setattr(prop, f"flag_{flag}", any(p in _search_text for p in patterns))

    # Status badge — search only the ksc_lozenge element, not full page text
    # (full page text always contains "under offer" in the sidebar nav)
    badge_text = ""
    badge = soup.find(class_=re.compile(r"ksc_lozenge"))
    if badge:
        badge_text = badge.get_text(strip=True).lower()
    for phrase, label in [
        ("sold stc",     "SSTC"),
        ("under offer",  "Under offer"),
        ("let agreed",   "Let agreed"),
    ]:
        if phrase in badge_text:
            prop.status = label
            prop.notes.append(label)
            break

    time.sleep(1.5)  # polite pacing

# ──────────────────────────────────────────────────────────────────────────────
# GEOCODING
# ──────────────────────────────────────────────────────────────────────────────

def geocode_postcode(prop: Property) -> bool:
    if not prop.postcode:
        prop.notes.append("No postcode — cannot geocode")
        return False

    # Try Nominatim first (full address → more accurate for rural properties)
    if prop.address:
        try:
            query = f"{prop.address}, {prop.postcode}, UK"
            r = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={"q": query, "format": "json", "limit": 1},
                headers={"User-Agent": "property-scorer/1.0"},
                timeout=10,
            )
            results = r.json()
            if results:
                prop.lat = float(results[0]["lat"])
                prop.lon = float(results[0]["lon"])
                # Still fetch ruc11 from postcodes.io (Nominatim doesn't have it)
                _fetch_ruc11(prop)
                return True
        except Exception:
            pass  # fall through to postcode centroid

    # Fallback: postcode centroid via postcodes.io
    try:
        r = requests.get(
            f"https://api.postcodes.io/postcodes/{prop.postcode.replace(' ', '')}",
            timeout=10,
        )
        data = r.json()
        if data.get("status") == 200:
            prop.lat = data["result"]["latitude"]
            prop.lon = data["result"]["longitude"]
            prop.rural_urban = data["result"].get("ruc11", "")
            return True
        prop.notes.append(f"Postcode lookup failed: {data.get('error', 'unknown')}")
    except Exception as e:
        prop.notes.append(f"Geocode error: {e}")
    return False


def _fetch_ruc11(prop: Property) -> None:
    try:
        r = requests.get(
            f"https://api.postcodes.io/postcodes/{prop.postcode.replace(' ', '')}",
            timeout=10,
        )
        data = r.json()
        if data.get("status") == 200:
            prop.rural_urban = data["result"].get("ruc11", "")
    except Exception:
        pass


def lookup_epc_floor_area(prop: Property, config: dict) -> None:
    """Try EPC register for floor area (m²). Skips if Rightmove scrape already got it."""
    if prop.floor_area is not None:
        return
    token = config.get("epc_api_key", "")
    if not token or not prop.postcode:
        return
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    _BASE = "https://api.get-energy-performance-data.communities.gov.uk/api"
    try:
        # Step 1: search postcode for certificate number
        r = requests.get(
            f"{_BASE}/domestic/search",
            params={"postcode": prop.postcode.replace(" ", ""), "page_size": 10},
            headers=headers,
            timeout=10,
        )
        rows = r.json().get("data", [])
        if not rows:
            return
        # Pick best address match
        addr_lower = prop.address.lower()
        def _similarity(row):
            epc_addr = " ".join([
                row.get("addressLine1") or "", row.get("addressLine2") or "",
                row.get("addressLine3") or "", row.get("addressLine4") or "",
            ]).lower()
            return sum(1 for w in addr_lower.split() if len(w) > 2 and w in epc_addr)
        best = max(rows, key=_similarity)
        cert_num = best.get("certificateNumber")
        if not cert_num:
            return
        # Step 2: fetch full certificate for floor area
        r2 = requests.get(
            f"{_BASE}/certificate",
            params={"certificate_number": cert_num},
            headers=headers,
            timeout=10,
        )
        cert = r2.json().get("data", r2.json())
        area = cert.get("total_floor_area")
        if area:
            prop.floor_area = float(area)
    except Exception as e:
        logging.debug(f"EPC lookup failed for {prop.postcode}: {e}")


# ──────────────────────────────────────────────────────────────────────────────
# SCORING FUNCTIONS (each returns 0–100)
# ──────────────────────────────────────────────────────────────────────────────

_RUC_SCORES = {
    "hamlet":     100,
    "isolated":   100,
    "village":     80,
    "rural town":  60,
    "rural":       70,
    "sparse":      45,
    "city":        15,
    "conurbation": 10,
    "urban":       25,
}

def score_tranquillity(prop: Property, config: dict) -> float:
    # Prefer CPRE raster if available
    if RASTERIO_AVAILABLE and prop.lat and prop.lon:
        raster_path = os.path.join(os.path.dirname(__file__), config["cpre_raster_path"])
        if os.path.exists(raster_path):
            try:
                with rasterio.open(raster_path) as src:
                    transformer = Transformer.from_crs("EPSG:4326", src.crs.to_epsg(), always_xy=True)
                    x, y = transformer.transform(prop.lon, prop.lat)
                    row, col = src.index(x, y)
                    value = float(src.read(1)[row, col])
                    if value != src.nodata and not np.isnan(value):
                        return float(np.clip(value, 0, 100))
            except Exception as e:
                prop.notes.append(f"Tranquillity raster error: {e}")

    # Fallback: ONS Rural-Urban Classification from postcodes.io
    ruc = prop.rural_urban.lower()
    if not ruc:
        return 50.0
    for keyword, score in _RUC_SCORES.items():
        if keyword in ruc:
            return float(score)
    return 50.0


def score_property_type(prop: Property, config: dict) -> float:
    ptype = prop.property_type.lower()
    for key, val in config["property_type_scores"].items():
        if key in ptype:
            return float(val)
    return 30.0


def score_bedrooms(prop: Property, config: dict) -> float:
    beds = prop.bedrooms
    t = config["thresholds"]
    if beds is None:
        return 50.0
    if beds < t["min_bedrooms"]:
        return 0.0
    target = t["target_bedrooms"]
    if beds == target:
        return 100.0
    if beds > target:
        return max(70.0, 100.0 - (beds - target) * 10)
    span = target - t["min_bedrooms"]
    return 60.0 + (beds - t["min_bedrooms"]) * (40.0 / max(1, span))


def score_price(prop: Property, config: dict) -> float:
    if prop.price is None:
        return 50.0
    max_p = config["thresholds"]["max_price_gbp"]
    if prop.price > max_p:
        return 0.0
    ratio = prop.price / max_p
    return float(np.clip(100 - ratio * 50, 50, 100))


def score_flood_risk(prop: Property, config: dict) -> float:
    if not prop.lat or not prop.lon:
        return 50.0
    try:
        url = (
            "https://environment.data.gov.uk/arcgis/rest/services/EA/"
            "FloodMapForPlanningRiversAndSea/FeatureServer/0/query"
        )
        params = {
            "geometry":       f"{prop.lon},{prop.lat}",
            "geometryType":   "esriGeometryPoint",
            "inSR":           "4326",
            "spatialRel":     "esriSpatialRelIntersects",
            "outFields":      "flood_zone",
            "returnGeometry": "false",
            "f":              "json",
        }
        r = requests.get(url, params=params, timeout=15)
        features = r.json().get("features", [])
        if not features:
            return 100.0  # not in any mapped flood zone
        zone = str(features[0].get("attributes", {}).get("flood_zone", "")).upper()
        return {"ZONE 1": 100.0, "1": 100.0,
                "ZONE 2": 40.0,  "2": 40.0,
                "ZONE 3": 0.0,   "3": 0.0}.get(zone, 50.0)
    except Exception as e:
        prop.notes.append(f"Flood risk error: {e}")
        return 50.0


# Cache loaded Ofcom area dataframes to avoid re-reading the same file
_ofcom_cache: dict = {}

def score_broadband(prop: Property, config: dict) -> float:
    if not prop.postcode:
        return 50.0

    # Extract postcode area (leading letters, e.g. "CW" from "CW10 0AA")
    m = re.match(r'^([A-Z]{1,2})', prop.postcode.upper())
    if not m:
        prop.notes.append("Cannot determine postcode area")
        return 50.0
    area = m.group(1)

    folder = os.path.join(os.path.dirname(__file__), config["ofcom_pc_folder"])
    csv_path = os.path.join(folder, f"202507_fixed_pc_coverage_r01_{area}.csv")
    if not os.path.exists(csv_path):
        prop.notes.append(f"Ofcom file not found for area {area}")
        return 50.0

    try:
        if area not in _ofcom_cache:
            df = pd.read_csv(csv_path, dtype=str, low_memory=False)
            df["_pc"] = df["postcode"].str.replace(" ", "").str.upper()
            _ofcom_cache[area] = df

        df = _ofcom_cache[area]
        pc = prop.postcode.replace(" ", "").upper()
        row = df[df["_pc"] == pc]
        if row.empty:
            prop.notes.append("Postcode not in Ofcom data")
            return 50.0

        r = row.iloc[0]
        gigabit = float(r.get("Gigabit availability (% premises)", 0) or 0)
        ufbb    = float(r.get("UFBB (100Mbit/s) availability (% premises)", 0) or 0)
        sfbb    = float(r.get("SFBB availability (% premises)", 0) or 0)

        # Score: best available tier wins; gigabit=100, 100Mbit=80, 30Mbit=60
        score = max(gigabit * 1.0, ufbb * 0.8, sfbb * 0.6)
        return float(np.clip(score, 0, 100))

    except Exception as e:
        prop.notes.append(f"Broadband error: {e}")
        return 50.0


def score_drive_time(prop: Property, config: dict) -> float:
    if not prop.lat or not prop.lon:
        return 50.0
    key = config["openrouteservice_api_key"]
    if "YOUR_ORS" in key:
        prop.notes.append("ORS key not configured")
        return 50.0
    try:
        r = requests.post(
            "https://api.openrouteservice.org/v2/directions/driving-car",
            headers={"Authorization": key, "Content-Type": "application/json"},
            json={
                "coordinates": [
                    [prop.lon, prop.lat],
                    [config["thresholds"]["reference_lon"],
                     config["thresholds"]["reference_lat"]],
                ]
            },
            timeout=15,
        )
        duration_s = r.json()["routes"][0]["summary"]["duration"]
        minutes = duration_s / 60
        max_m = config["thresholds"]["max_drive_minutes"]
        return float(np.clip(100 * (1 - minutes / max_m), 0, 100))
    except Exception as e:
        prop.notes.append(f"Drive time error: {e}")
        return 50.0

# ──────────────────────────────────────────────────────────────────────────────
# AGGREGATION
# ──────────────────────────────────────────────────────────────────────────────

SCORERS = {
    "tranquillity":  score_tranquillity,
    "property_type": score_property_type,
    "flood_risk":    score_flood_risk,
    "price":         score_price,
    "bedrooms":      score_bedrooms,
    "broadband":     score_broadband,
    "drive_time":    score_drive_time,
}

def score_property(prop: Property, config: dict) -> None:
    weights = config["weights"]
    total = 0.0
    for key, scorer in SCORERS.items():
        s = scorer(prop, config)
        prop.scores[key] = round(s, 1)
        total += s * weights[key]
    prop.total_score = round(total, 1)

# ──────────────────────────────────────────────────────────────────────────────
# OUTPUT
# ──────────────────────────────────────────────────────────────────────────────

def print_results(properties: list, config: dict) -> None:
    props = sorted(properties, key=lambda p: p.total_score, reverse=True)
    w = config["weights"]

    headers = [
        "#", "Address", "Price", "Beds", "Type", "TOTAL",
        f"Tranq\n{int(w['tranquillity']*100)}%",
        f"Type\n{int(w['property_type']*100)}%",
        f"Flood\n{int(w['flood_risk']*100)}%",
        f"Price\n{int(w['price']*100)}%",
        f"Beds\n{int(w['bedrooms']*100)}%",
        f"BB\n{int(w['broadband']*100)}%",
        f"Drive\n{int(w['drive_time']*100)}%",
        "URL",
    ]

    rows = []
    for i, p in enumerate(props, 1):
        s = p.scores
        rows.append([
            i,
            (p.address or p.postcode or "?")[:35],
            f"£{p.price:,}" if p.price else "?",
            p.bedrooms or "?",
            (p.property_type or "?")[:14],
            p.total_score,
            s.get("tranquillity", "-"),
            s.get("property_type", "-"),
            s.get("flood_risk", "-"),
            s.get("price", "-"),
            s.get("bedrooms", "-"),
            s.get("broadband", "-"),
            s.get("drive_time", "-"),
            p.url,
        ])

    print("\n" + tabulate(rows, headers=headers, tablefmt="rounded_outline"))
    print()

    for p in props:
        if p.notes:
            label = (p.address or p.url)[:30]
            print(f"  [{label}] {'; '.join(p.notes)}")

CSV_COLUMNS = [
    "url", "address", "price", "bedrooms", "property_type", "postcode",
    "lat", "lon", "status", "floor_area", "price_per_sqm", "land_acres",
    "total_score", "tranquillity", "property_type_score", "flood_risk",
    "price_score", "bedrooms_score", "broadband", "drive_time",
    "key_features", "description",
    "views", "stone_built", "garage", "outbuildings", "log_burner", "aga",
    "no_chain", "annexe", "woodland", "planning", "period_character", "paddock",
    "land", "holiday",
    "notes",
]

def load_seen(config: dict) -> dict:
    """Returns {prop_id: stored_price} for all previously scored properties."""
    path = os.path.join(os.path.dirname(__file__), config["results_csv"])
    if not os.path.exists(path):
        return {}
    try:
        df = pd.read_csv(path, usecols=["url", "price"], dtype=str)
        result = {}
        for _, row in df.iterrows():
            m = re.search(r"/properties/(\d+)", str(row.get("url", "")))
            if m:
                try:
                    result[m.group(1)] = int(float(row["price"])) if pd.notna(row["price"]) else None
                except (ValueError, TypeError):
                    result[m.group(1)] = None
        return result
    except Exception:
        return {}

def save_results(properties: list, config: dict) -> None:
    if not properties:
        return
    path = os.path.join(os.path.dirname(__file__), config["results_csv"])

    def prop_to_row(p):
        s = p.scores
        return {
            "url":                 p.url,
            "address":             p.address,
            "price":               p.price,
            "bedrooms":            p.bedrooms,
            "property_type":       p.property_type,
            "postcode":            p.postcode,
            "lat":                 p.lat,
            "lon":                 p.lon,
            "status":              p.status,
            "floor_area":          p.floor_area,
            "price_per_sqm":       round(p.price / p.floor_area) if p.price and p.floor_area else "",
            "land_acres":          p.land_acres,
            "total_score":         p.total_score,
            "tranquillity":        s.get("tranquillity", ""),
            "property_type_score": s.get("property_type", ""),
            "flood_risk":          s.get("flood_risk", ""),
            "price_score":         s.get("price", ""),
            "bedrooms_score":      s.get("bedrooms", ""),
            "broadband":           s.get("broadband", ""),
            "drive_time":          s.get("drive_time", ""),
            "key_features":        " | ".join(p.key_features),
            "description":         p.description,
            "views":               "Y" if p.flag_views else "",
            "stone_built":         "Y" if p.flag_stone_built else "",
            "garage":              "Y" if p.flag_garage else "",
            "outbuildings":        "Y" if p.flag_outbuildings else "",
            "log_burner":          "Y" if p.flag_log_burner else "",
            "aga":                 "Y" if p.flag_aga else "",
            "no_chain":            "Y" if p.flag_no_chain else "",
            "annexe":              "Y" if p.flag_annexe else "",
            "woodland":            "Y" if p.flag_woodland else "",
            "planning":            "Y" if p.flag_planning else "",
            "period_character":    "Y" if p.flag_period_character else "",
            "paddock":             "Y" if p.flag_paddock else "",
            "land":                "Y" if p.flag_land else "",
            "holiday":             "Y" if p.flag_holiday else "",
            "notes":               "; ".join(p.notes),
        }

    new_rows = {
        re.search(r"/properties/(\d+)", p.url).group(1): prop_to_row(p)
        for p in properties
        if re.search(r"/properties/(\d+)", p.url)
    }

    if os.path.exists(path):
        existing = pd.read_csv(path, dtype=str)
        existing["_id"] = existing["url"].str.extract(r"/properties/(\d+)")
        existing = existing[~existing["_id"].isin(new_rows.keys())].drop(columns=["_id"])
        updated = pd.concat([existing, pd.DataFrame(new_rows.values(), columns=CSV_COLUMNS)], ignore_index=True)
    else:
        updated = pd.DataFrame(new_rows.values(), columns=CSV_COLUMNS)

    updated.to_csv(path, index=False, encoding="utf-8")
    logging.info(f"Saved {len(new_rows)} properties to {path}")

# ──────────────────────────────────────────────────────────────────────────────
# MAP GENERATION
# ──────────────────────────────────────────────────────────────────────────────

def generate_map(config: dict) -> None:
    csv_path = os.path.join(os.path.dirname(__file__), config["results_csv"])
    if not os.path.exists(csv_path):
        return

    df = pd.read_csv(csv_path, dtype=str)

    # Geocode any rows missing lat/lon
    def _has_coords(row):
        try:
            v = float(row.get("lat", ""))
            return v == v
        except (TypeError, ValueError):
            return False

    for idx, row in df.iterrows():
        if _has_coords(row):
            continue
        pc = str(row.get("postcode", "")).replace(" ", "")
        addr = str(row.get("address", ""))
        if not pc or pc == "nan":
            continue
        geocoded = False
        if addr and addr != "nan":
            try:
                r = requests.get(
                    "https://nominatim.openstreetmap.org/search",
                    params={"q": f"{addr}, {pc}, UK", "format": "json", "limit": 1},
                    headers={"User-Agent": "property-scorer/1.0"},
                    timeout=10,
                )
                results = r.json()
                if results:
                    df.at[idx, "lat"] = str(results[0]["lat"])
                    df.at[idx, "lon"] = str(results[0]["lon"])
                    geocoded = True
            except Exception:
                pass
        if not geocoded:
            try:
                r = requests.get(f"https://api.postcodes.io/postcodes/{pc}", timeout=10)
                result = r.json().get("result", {})
                if result:
                    df.at[idx, "lat"] = str(result["latitude"])
                    df.at[idx, "lon"] = str(result["longitude"])
            except Exception:
                pass

    def _sf(val, default=None):
        try:
            f = float(val)
            return default if f != f else f
        except (TypeError, ValueError):
            return default

    def _ss(val, default=""):
        s = str(val) if val is not None else ""
        return default if s in ("", "nan", "None") else s

    # Build property data array for JS
    props_data = []
    for _, row in df.iterrows():
        lat = _sf(row.get("lat"))
        lon = _sf(row.get("lon"))
        if lat is None or lon is None:
            continue
        price_f = _sf(row.get("price"))
        price_str = f"£{int(price_f):,}" if price_f else "?"
        fa = _sf(row.get("floor_area"))
        ppsm = _sf(row.get("price_per_sqm"))
        kf = _ss(row.get("key_features", ""))
        kf_short = "; ".join(kf.split(" | ")[:4]) if kf else ""
        props_data.append({
            "lat":                lat,
            "lon":                lon,
            "address":            _ss(row.get("address", ""))[:60],
            "price":              price_str,
            "beds":               _ss(row.get("bedrooms", ""), "?"),
            "ptype":              _ss(row.get("property_type", "")).capitalize(),
            "status":             _ss(row.get("status", "")),
            "url":                _ss(row.get("url", "")),
            "floor_area":         fa,
            "price_per_sqm":      ppsm,
            "land_acres":         _sf(row.get("land_acres")),
            "price_num":          price_f or 0,
            "beds_num":           _sf(row.get("bedrooms"), 0) or 0,
            "fa_num":             fa or 0,
            "land_num":           _sf(row.get("land_acres"), 0) or 0,
            "key_features":       kf_short,
            "tranquillity":       _sf(row.get("tranquillity"), 0),
            "property_type_score":_sf(row.get("property_type_score"), 0),
            "flood_risk":         _sf(row.get("flood_risk"), 0),
            "price_score":        _sf(row.get("price_score"), 0),
            "bedrooms_score":     _sf(row.get("bedrooms_score"), 0),
            "broadband":          _sf(row.get("broadband"), 0),
            "drive_time":         _sf(row.get("drive_time"), 0),
        })

    props_json = json.dumps(props_data)
    centre_lat = config["thresholds"]["reference_lat"]
    centre_lon = config["thresholds"]["reference_lon"]
    w = config["weights"]
    def _wi(v): return int(round(v * 100))

    # Filter slider ranges derived from data
    prices  = [p["price_num"] for p in props_data if p["price_num"]]
    beds    = [p["beds_num"]  for p in props_data if p["beds_num"]]
    areas   = [p["fa_num"]    for p in props_data if p["fa_num"]]
    lands   = [p["land_num"]  for p in props_data if p["land_num"]]
    filt_max_price   = int(max(prices) / 50000 + 1) * 50000 if prices else 2000000
    filt_max_beds    = int(max(beds))  if beds  else 6
    filt_max_area    = int(max(areas) / 25 + 1) * 25 if areas else 500
    filt_max_land    = round(max(lands) + 0.5) if lands else 20
    filt_price_step  = 25000

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Property Scores</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
  body {{ margin:0; display:flex; height:100vh; font-family:sans-serif; }}
  #panel {{
    width:270px; min-width:270px; background:#f8f8f8; padding:12px;
    overflow-y:auto; box-shadow:2px 0 6px rgba(0,0,0,.15); z-index:1000;
  }}
  #panel h3 {{ margin:0 0 10px; font-size:14px; }}
  .srow {{ margin-bottom:8px; font-size:12px; }}
  .srow label {{ display:flex; justify-content:space-between; margin-bottom:2px; }}
  .srow input[type=range] {{ width:100%; }}
  #total {{ font-size:12px; color:#888; margin:6px 0 10px; }}
  #map {{ flex:1; }}
  .legend {{ background:#fff; padding:8px 10px; border-radius:6px; line-height:1.8em; font-size:12px; }}
  .legend span {{ display:inline-block; width:12px; height:12px; border-radius:50%; margin-right:5px; vertical-align:middle; }}
</style>
</head>
<body>
<div id="panel">
  <h3>Score Weights</h3>
  <div class="srow"><label><span>Tranquillity</span><b id="lv_tranquillity">{_wi(w['tranquillity'])}</b></label>
    <input type="range" id="w_tranquillity" min="0" max="100" value="{_wi(w['tranquillity'])}" oninput="update()"></div>
  <div class="srow"><label><span>Property type</span><b id="lv_property_type">{_wi(w['property_type'])}</b></label>
    <input type="range" id="w_property_type" min="0" max="100" value="{_wi(w['property_type'])}" oninput="update()"></div>
  <div class="srow"><label><span>Flood risk</span><b id="lv_flood_risk">{_wi(w['flood_risk'])}</b></label>
    <input type="range" id="w_flood_risk" min="0" max="100" value="{_wi(w['flood_risk'])}" oninput="update()"></div>
  <div class="srow"><label><span>Price</span><b id="lv_price">{_wi(w['price'])}</b></label>
    <input type="range" id="w_price" min="0" max="100" value="{_wi(w['price'])}" oninput="update()"></div>
  <div class="srow"><label><span>Bedrooms</span><b id="lv_bedrooms">{_wi(w['bedrooms'])}</b></label>
    <input type="range" id="w_bedrooms" min="0" max="100" value="{_wi(w['bedrooms'])}" oninput="update()"></div>
  <div class="srow"><label><span>Broadband</span><b id="lv_broadband">{_wi(w['broadband'])}</b></label>
    <input type="range" id="w_broadband" min="0" max="100" value="{_wi(w['broadband'])}" oninput="update()"></div>
  <div class="srow"><label><span>Drive time</span><b id="lv_drive_time">{_wi(w['drive_time'])}</b></label>
    <input type="range" id="w_drive_time" min="0" max="100" value="{_wi(w['drive_time'])}" oninput="update()"></div>
  <div id="total"></div>
  <button onclick="reset()" style="font-size:12px;width:100%;padding:4px;">Reset defaults</button>

  <h3 style="margin:14px 0 8px;">Filters</h3>
  <div class="srow"><label><span>Max price</span><b id="lv_price_max">£{filt_max_price:,}</b></label>
    <input type="range" id="f_price_max" min="0" max="{filt_max_price}" step="{filt_price_step}" value="{filt_max_price}" oninput="update()"></div>
  <div class="srow"><label><span>Min bedrooms</span><b id="lv_beds_min">any</b></label>
    <input type="range" id="f_beds_min" min="0" max="{filt_max_beds}" step="1" value="0" oninput="update()"></div>
  <div class="srow"><label><span>Min floor area (m²)</span><b id="lv_area_min">any</b></label>
    <input type="range" id="f_area_min" min="0" max="{filt_max_area}" step="25" value="0" oninput="update()"></div>
  <div class="srow"><label><span>Min land (acres)</span><b id="lv_land_min">any</b></label>
    <input type="range" id="f_land_min" min="0" max="{filt_max_land}" step="0.5" value="0" oninput="update()"></div>
  <button onclick="resetFilters()" style="font-size:12px;width:100%;padding:4px;margin-top:4px;">Reset filters</button>
</div>
<div id="map"></div>
<script>
var DEFAULTS = {{tranquillity:{_wi(w['tranquillity'])},property_type:{_wi(w['property_type'])},flood_risk:{_wi(w['flood_risk'])},price:{_wi(w['price'])},bedrooms:{_wi(w['bedrooms'])},broadband:{_wi(w['broadband'])},drive_time:{_wi(w['drive_time'])}}};
var properties = {props_json};

var map = L.map('map').setView([{centre_lat},{centre_lon}], 11);
L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
  attribution: '© OpenStreetMap contributors'
}}).addTo(map);
L.marker([{centre_lat},{centre_lon}]).bindPopup('<b>Congleton</b>').addTo(map);

var layerGroup = L.layerGroup().addTo(map);

var KEYS = ['tranquillity','property_type','flood_risk','price','bedrooms','broadband','drive_time'];

function getWeights() {{
  var w = {{}};
  KEYS.forEach(function(k) {{ w[k] = parseInt(document.getElementById('w_'+k).value); }});
  return w;
}}

function calcScore(prop, w) {{
  var wsum = KEYS.reduce(function(a,k){{return a+w[k];}}, 0);
  if (wsum === 0) return 0;
  return (w.tranquillity * prop.tranquillity +
          w.property_type * prop.property_type_score +
          w.flood_risk * prop.flood_risk +
          w.price * prop.price_score +
          w.bedrooms * prop.bedrooms_score +
          w.broadband * prop.broadband +
          w.drive_time * prop.drive_time) / wsum;
}}

function scoreColor(s) {{
  if (s >= 70) return '#27ae60';
  if (s >= 55) return '#f39c12';
  return '#e74c3c';
}}

function fmt(v, digits) {{
  if (v === null || v === undefined) return null;
  return digits === 0 ? Math.round(v) : Math.round(v * Math.pow(10,digits)) / Math.pow(10,digits);
}}

function getFilters() {{
  var priceMax = parseFloat(document.getElementById('f_price_max').value);
  var bedsMin  = parseFloat(document.getElementById('f_beds_min').value);
  var areaMin  = parseFloat(document.getElementById('f_area_min').value);
  var landMin  = parseFloat(document.getElementById('f_land_min').value);
  return {{priceMax:priceMax, bedsMin:bedsMin, areaMin:areaMin, landMin:landMin}};
}}

function passesFilters(p, f) {{
  if (f.priceMax > 0 && p.price_num > 0 && p.price_num > f.priceMax) return false;
  if (f.bedsMin  > 0 && p.beds_num  > 0 && p.beds_num  < f.bedsMin)  return false;
  if (f.areaMin  > 0 && p.fa_num    > 0 && p.fa_num    < f.areaMin)  return false;
  if (f.landMin  > 0 && p.land_num  >= 0 && p.land_num < f.landMin)  return false;
  return true;
}}

function renderMarkers(w, f) {{
  layerGroup.clearLayers();
  var shown = 0;
  properties.forEach(function(p) {{
    if (!passesFilters(p, f)) return;
    shown++;
    var sstc = p.status === 'SSTC' || p.status === 'Under offer';
    var score = sstc ? null : Math.round(calcScore(p, w) * 10) / 10;
    var color = sstc ? '#888888' : scoreColor(score);

    var popup = '<b>' + p.address + '</b><br>'
      + p.ptype + ' &middot; ' + p.beds + ' bed &middot; ' + p.price + '<br>';
    if (p.floor_area) {{
      popup += fmt(p.floor_area,0) + ' m²';
      if (p.price_per_sqm) popup += ' &middot; £' + fmt(p.price_per_sqm,0).toLocaleString() + '/m²';
      popup += '<br>';
    }}
    if (p.land_acres) popup += p.land_acres + ' acres<br>';
    if (p.status) popup += '<b style="color:#c0392b">' + p.status + '</b><br>';
    popup += '<b>Score: ' + (score !== null ? score : '—') + '</b><br>';
    popup += 'Tranq:' + p.tranquillity + ' Flood:' + p.flood_risk
           + ' BB:' + p.broadband + ' Drive:' + p.drive_time + '<br>';
    if (p.key_features) popup += '<i style="font-size:11px">' + p.key_features + '</i><br>';
    popup += '<a href="' + p.url + '" target="_blank">View on Rightmove</a>';

    L.circleMarker([p.lat, p.lon], {{
      radius:12, color:color, fillColor:color, fillOpacity:0.85, weight:2
    }}).bindPopup(popup).addTo(layerGroup);
  }});
  document.getElementById('total').textContent = 'Weights total: ' + KEYS.reduce(function(a,k){{return a+w[k];}},0) + '  |  Showing: ' + shown + ' / ' + properties.length;
}}

function update() {{
  var w = getWeights();
  var f = getFilters();
  KEYS.forEach(function(k) {{
    document.getElementById('lv_'+k).textContent = w[k];
  }});
  var pm = f.priceMax; document.getElementById('lv_price_max').textContent = '£' + pm.toLocaleString();
  document.getElementById('lv_beds_min').textContent  = f.bedsMin  > 0 ? f.bedsMin  + '+' : 'any';
  document.getElementById('lv_area_min').textContent  = f.areaMin  > 0 ? f.areaMin  + ' m²' : 'any';
  document.getElementById('lv_land_min').textContent  = f.landMin  > 0 ? f.landMin  + ' ac' : 'any';
  renderMarkers(w, f);
}}

function reset() {{
  KEYS.forEach(function(k) {{ document.getElementById('w_'+k).value = DEFAULTS[k]; }});
  update();
}}

function resetFilters() {{
  document.getElementById('f_price_max').value = document.getElementById('f_price_max').max;
  document.getElementById('f_beds_min').value  = 0;
  document.getElementById('f_area_min').value  = 0;
  document.getElementById('f_land_min').value  = 0;
  update();
}}

// Legend
var legend = L.control({{position:'bottomright'}});
legend.onAdd = function() {{
  var d = L.DomUtil.create('div','legend');
  d.innerHTML = '<b>Score</b><br>'
    + '<span style="background:#27ae60"></span> &ge; 70<br>'
    + '<span style="background:#f39c12"></span> 55&ndash;69<br>'
    + '<span style="background:#e74c3c"></span> &lt; 55<br>'
    + '<span style="background:#888888"></span> SSTC / Under offer';
  return d;
}};
legend.addTo(map);

update();
</script>
</body>
</html>"""

    map_path = os.path.join(os.path.dirname(__file__), "map.html")
    with open(map_path, "w", encoding="utf-8") as f:
        f.write(html)
    logging.info(f"Map saved to {map_path}")

# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s",
                        stream=sys.stdout)

    seen = load_seen(CONFIG)
    if seen:
        logging.info(f"{len(seen)} properties already in results.csv")

    print("Reading Rightmove alert emails...")
    properties = fetch_rightmove_emails(CONFIG)

    # Merge manual URLs (deduplicate by property ID)
    manual = load_manual_properties(CONFIG)
    existing_ids = {re.search(r"/properties/(\d+)", p.url).group(1)
                    for p in properties if re.search(r"/properties/(\d+)", p.url)}
    for p in manual:
        m = re.search(r"/properties/(\d+)", p.url)
        if m and m.group(1) not in existing_ids:
            properties.append(p)
            existing_ids.add(m.group(1))

    # Enrich all — needed to detect price changes on already-seen properties
    print(f"Enriching {len(properties)} properties from listing pages...")
    for i, prop in enumerate(properties, 1):
        m = re.search(r"/properties/(\d+)", prop.url)
        prop_id = m.group(1) if m else None
        if prop_id and prop_id in seen and seen[prop_id] is not None:
            # Quick check: skip enrichment if we'll likely skip this property
            # (full enrich happens below only for new/price-changed)
            prop._seen_price = seen[prop_id]
        else:
            prop._seen_price = None
        print(f"  [{i}/{len(properties)}] {prop.url}")
        enrich_from_listing(prop)
        geocode_postcode(prop)
        lookup_epc_floor_area(prop, CONFIG)

    # Filter: keep new properties and those with a changed price
    to_score = []
    for prop in properties:
        m = re.search(r"/properties/(\d+)", prop.url)
        prop_id = m.group(1) if m else None
        if prop_id not in seen:
            to_score.append(prop)
        elif prop.price and prop.price != seen.get(prop_id):
            logging.info(f"Price changed for {prop.url}: £{seen[prop_id]:,} → £{prop.price:,}")
            to_score.append(prop)
        else:
            logging.info(f"Skipping (unchanged): {prop.url}")
    properties = to_score

    if not properties:
        print("No new or price-changed properties found.")
        generate_map(CONFIG)
        return

    print("Scoring...")
    for prop in properties:
        score_property(prop, CONFIG)

    print_results(properties, CONFIG)
    try:
        save_results(properties, CONFIG)
    except PermissionError:
        print("WARNING: results.csv is open in another program — close it and re-run to save.")
    generate_map(CONFIG)


if __name__ == "__main__":
    main()
