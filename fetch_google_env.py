#!/usr/bin/env python3
"""
IKDR Kawasaki environmental study: fetch everything we can from Google Maps Platform
for each patient postal code (US ZIP / Canadian FSA).

One-off script. Only postal codes are sent to Google, never patient IDs.

Per location it collects:
  geocode      Geocoding API     centroid lat/lng, bounds
  elevation    Elevation API     ground elevation (m)
  images       Static Maps API   roadmap + satellite + "analysis" styled map at 1, 2, 3 and 10 km,
                                 no markers; % green / % blue pixels; distance to nearest water,
                                 highway and freeway (m), measured on the analysis map
  traffic      Routes API        typical traffic Tue/Thu/Sat 09:00 & 17:00 local, 4-level scale
  airquality   Air Quality API   current conditions + hourly history of the last 30 days
  pollen       Pollen API        5-day forecast (the API has no history)
  weather      Weather API       current conditions, last 24 h, 10-day forecast
  solar        Solar API         nearest building's solar potential (sunshine hours/year, ...)

Usage
  export GOOGLE_MAPS_API_KEY=...            # or put it in a .env file next to this script
  python fetch_google_env.py --input ikdr_all_postal_codes.xlsx --fsa-regions FSA-merge-Info.xlsx --estimate
  python fetch_google_env.py --input ... --fsa-regions ... --limit 3        # try 3 locations first
  python fetch_google_env.py --input ... --fsa-regions ...                  # full run

Every API response is cached under <out>/cache, so a re-run (after a crash, or to add a step)
never pays for the same request twice.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import logging
import math
import os
import re
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests
from PIL import Image
from tqdm import tqdm

# =============================================================================
# Settings (agree these with the team)
# =============================================================================
ALTITUDES_M = (1000, 2000, 3000, 10000)
# Image footprint (m per side) = altitude * SPAN_FACTOR. 1.0 = a "1,000 m" image covers
# 1 km x 1 km of ground (the convention of Google Maps web URLs, "@lat,lng,1000m").
SPAN_FACTOR = 1.0
OUTPUT_IMAGE_PX = 1024                  # every saved image is resampled to this size
DISTANCE_FALLBACK_SPAN_M = 50000        # extra analysis image when nothing is found within 10 km
CA_LEVEL = "fsa"                        # Canadian codes: "fsa" (first 3 chars) or "full"

TRAFFIC_SLOTS = [("Tue", 1, 9), ("Tue", 1, 17), ("Thu", 3, 9),
                 ("Thu", 3, 17), ("Sat", 5, 9), ("Sat", 5, 17)]   # (label, weekday Mon=0, hour)
TRAFFIC_PROBE_HALF_M = 400              # two probe routes (W->E, S->N), 800 m long, through the centroid
TRAFFIC_LEVEL_CUTS = (1.10, 1.30, 1.60) # traffic/free-flow time ratio -> 1 fast,2 moderate,3 slow,4 heavy

AQ_HISTORY_HOURS = 720                  # API maximum (30 days)
POLLEN_DAYS = 5                         # API maximum
WEATHER_HISTORY_HOURS = 24              # API maximum
WEATHER_FORECAST_DAYS = 10
SOLAR_QUALITY = "LOW"                   # accept lowest imagery quality = best coverage

# Pixel classes for satellite and default roadmap images (HSV)
GREEN_HUE, BLUE_HUE = (70, 170), (180, 260)
MIN_SAT, MIN_VAL = 0.15, 0.15
# Analysis map: pure colours drawn by the style below
PALETTE = {"other": (255, 255, 255), "green": (0, 255, 0), "water": (0, 0, 255),
           "highway": (255, 0, 0), "freeway": (255, 0, 255)}
PALETTE_TOL = 60.0

ANALYSIS_STYLE = [
    "element:labels|visibility:off",
    "element:geometry|color:0xffffff",
    "feature:road|visibility:off",
    "feature:transit|visibility:off",
    "feature:administrative|visibility:off",
    "feature:landscape.natural.landcover|element:geometry|color:0x00ff00",
    "feature:poi.park|element:geometry|color:0x00ff00",
    "feature:water|element:geometry|color:0x0000ff",
    "feature:road.highway|element:geometry|visibility:on",
    "feature:road.highway|element:geometry|color:0xff0000",
    "feature:road.highway.controlled_access|element:geometry|visibility:on",
    "feature:road.highway.controlled_access|element:geometry|color:0xff00ff",
]

URL = {
    "geocode": "https://maps.googleapis.com/maps/api/geocode/json",
    "elevation": "https://maps.googleapis.com/maps/api/elevation/json",
    "static": "https://maps.googleapis.com/maps/api/staticmap",
    "routes": "https://routes.googleapis.com/directions/v2:computeRoutes",
    "aq_current": "https://airquality.googleapis.com/v1/currentConditions:lookup",
    "aq_history": "https://airquality.googleapis.com/v1/history:lookup",
    "pollen": "https://pollen.googleapis.com/v1/forecast:lookup",
    "wx_current": "https://weather.googleapis.com/v1/currentConditions:lookup",
    "wx_history": "https://weather.googleapis.com/v1/history/hours:lookup",
    "wx_forecast": "https://weather.googleapis.com/v1/forecast/days:lookup",
    "solar": "https://solar.googleapis.com/v1/buildingInsights:findClosest",
}
STEPS = ["geocode", "elevation", "images", "traffic", "airquality", "pollen", "weather", "solar"]
# Default = only what Cedric's list needs and the key has enabled. Add the others with --steps.
DEFAULT_STEPS = ["images", "traffic", "airquality", "pollen"]

log = logging.getLogger("ikdr")


# =============================================================================
# HTTP client: disk cache + retries + call counter
# =============================================================================
class Google:
    RETRY = {429, 500, 502, 503, 504}

    def __init__(self, key: str, cache_dir: Path, session: requests.Session | None = None,
                 min_interval: float = 0.05):
        self.key, self.cache_dir = key, cache_dir
        self.s = session or requests.Session()
        self.min_interval, self._last = min_interval, 0.0
        self.sent, self.hits = Counter(), Counter()

    def _path(self, api, method, url, params, body, tag, binary):
        blob = json.dumps([method, url, sorted((params or {}).items(), key=str), body, tag],
                          sort_keys=True, default=str)
        return self.cache_dir / api / (hashlib.sha256(blob.encode()).hexdigest()[:32]
                                       + (".png" if binary else ".json"))

    def call(self, api, method, url, params=None, body=None, *, headers=None,
             binary=False, tag=None):
        """-> (ok, data). data = bytes for images, dict otherwise ({'error': ...} on failure)."""
        path = self._path(api, method, url, params, body, tag, binary)
        err_path = path.with_suffix(".err.json")
        if path.exists():
            self.hits[api] += 1
            return True, path.read_bytes() if binary else json.loads(path.read_text())
        if err_path.exists():  # deterministic errors (400/404) are cached too
            self.hits[api] += 1
            return False, json.loads(err_path.read_text())

        params, headers = dict(params or {}), dict(headers or {})
        if "maps.googleapis.com/maps/api" in url:
            params["key"] = self.key
        else:
            headers["X-Goog-Api-Key"] = self.key

        delay = 2.0
        for attempt in range(6):
            wait = self.min_interval - (time.monotonic() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.monotonic()
            try:
                r = self.s.request(method, url, params=params, json=body, headers=headers, timeout=60)
            except requests.RequestException as e:
                status, err = -1, f"network: {type(e).__name__}"
            else:
                self.sent[api] += 1
                status = r.status_code
                if status < 400:
                    ok, data = self._parse(r, binary)
                    if ok:
                        path.parent.mkdir(parents=True, exist_ok=True)
                        path.write_bytes(data) if binary else path.write_text(json.dumps(data))
                        return True, data
                    status, err = 400, data["error"]
                else:
                    err = _err_text(r)
            if status in self.RETRY or status == -1:
                if attempt < 5:
                    time.sleep(delay)
                    delay = min(delay * 2, 60)
                    continue
            break
        if status in (400, 404):
            err_path.parent.mkdir(parents=True, exist_ok=True)
            err_path.write_text(json.dumps({"error": err, "status": status}))
        if status in (401, 403):
            raise SystemExit(f"{api}: HTTP {status} {err}\n-> Is this API enabled for the key "
                             f"and is billing set up? Nothing was cached for this request.")
        log.warning("%s: HTTP %s %s", api, status, err)
        return False, {"error": err, "status": status}

    @staticmethod
    def _parse(r, binary):
        if binary:
            if not r.headers.get("Content-Type", "").startswith("image/"):
                return False, {"error": r.text[:300]}
            return True, r.content
        data = r.json() if r.content else {}
        st = data.get("status") if isinstance(data, dict) else None   # geocode/elevation style
        if st and st not in ("OK", "ZERO_RESULTS"):
            if st in ("REQUEST_DENIED", "OVER_DAILY_LIMIT"):
                raise SystemExit(f"{st}: {data.get('error_message')}")
            return False, {"error": f"{st} {data.get('error_message', '')}".strip()}
        return True, data


def _err_text(r):
    try:
        e = r.json().get("error", {})
        return f"{e.get('status', '')} {e.get('message', '')}".strip() if isinstance(e, dict) else str(e)
    except ValueError:
        return r.text[:300]


# =============================================================================
# Locations
# =============================================================================
CA_RE = re.compile(r"^[ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z](\d[ABCEGHJ-NPRSTV-Z]\d)?$")


def find_col(df, *names, required=True):
    low = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n in low:
            return low[n]
    if required:
        raise SystemExit(f"Could not find any of {names} in columns {list(df.columns)}")


def normalize_postal(raw):
    """-> (country, code, problem). Excel drops leading zeros of US ZIPs: restore them."""
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None, None, "missing"
    if isinstance(raw, float) and raw.is_integer():
        raw = int(raw)
    s = re.sub(r"[\s\-]", "", str(raw)).upper()
    if s in ("", "NAN", "NA", "N/A", "NONE", "UNKNOWN", "0"):
        return None, None, "missing"
    if s.isdigit():
        if len(s) <= 5:
            return "US", s.zfill(5), None
        if len(s) == 9:
            return "US", s[:5], None
    if CA_RE.match(s):
        return "CA", s, None
    return None, None, f"unrecognised: {raw}"


def read_table(path):
    return pd.read_csv(path, dtype=object) if str(path).lower().endswith(".csv") \
        else pd.read_excel(path, dtype=object)


def load_regions(path):
    if not path:
        return {}
    df = read_table(path)
    reg, names = find_col(df, "reg", "region"), find_col(df, "regnames", "fsas")
    pop = find_col(df, "population", required=False)
    out = {}
    for _, r in df.dropna(subset=[reg, names]).iterrows():
        for fsa in str(r[names]).split(","):
            if fsa.strip():
                out[fsa.strip().upper()] = (str(r[reg]).strip(), r[pop] if pop else None)
    return out


def build_locations(input_path, regions_path, ca_level):
    df = read_table(input_path)
    id_c = find_col(df, "ikdr_regid", "regid", "id")
    zip_c = find_col(df, "zip", "postal_code", "postal code", "postalcode")
    regions = load_regions(regions_path)
    rows = []
    for _, r in df.iterrows():
        country, code, problem = normalize_postal(r[zip_c])
        key = query = None
        if country:
            query = code[:3] if (country == "CA" and ca_level == "fsa") else code
            key = f"{country}_{query}"
        reg, pop = regions.get(code[:3], (None, None)) if country == "CA" else (None, None)
        rows.append({"IKDR_REGID": r[id_c], "zip_raw": r[zip_c], "country": country,
                     "postal_code": code, "location_key": key, "query_code": query,
                     "fsa_region": reg, "fsa_region_population": pop, "postal_problem": problem})
    patients = pd.DataFrame(rows)
    locs = (patients.dropna(subset=["location_key"])
            .groupby("location_key", as_index=False)
            .agg(country=("country", "first"), query_code=("query_code", "first"),
                 fsa_region=("fsa_region", "first"), n_patients=("IKDR_REGID", "size")))
    return patients, locs


# =============================================================================
# Geometry helpers
# =============================================================================
R_EARTH = 6371008.8


def zoom_for_span(lat, span_m):
    """Highest integer zoom whose 640 px image covers span_m with >=10% margin (cut off
    later, removing the Google logo). -> (zoom, centre crop fraction)."""
    mpp0 = 156543.03392 * math.cos(math.radians(lat))
    z = max(0, min(21, math.floor(math.log2(640 * mpp0 / (span_m / 0.9)))))
    return z, span_m / (640 * mpp0 / 2 ** z)


def offset(lat, lng, east_m, north_m):
    return (lat + math.degrees(north_m / R_EARTH),
            lng + math.degrees(east_m / (R_EARTH * math.cos(math.radians(lat)))))


def haversine(a, b):
    p1, p2 = math.radians(a[0]), math.radians(b[0])
    h = (math.sin((p2 - p1) / 2) ** 2
         + math.cos(p1) * math.cos(p2) * math.sin(math.radians(b[1] - a[1]) / 2) ** 2)
    return 2 * R_EARTH * math.asin(math.sqrt(h))


def decode_polyline(s):
    pts, i, lat, lng = [], 0, 0, 0
    while i < len(s):
        vals = []
        for _ in range(2):
            shift = res = 0
            while True:
                b = ord(s[i]) - 63
                i += 1
                res |= (b & 0x1F) << shift
                shift += 5
                if b < 0x20:
                    break
            vals.append(~(res >> 1) if res & 1 else res >> 1)
        lat, lng = lat + vals[0], lng + vals[1]
        pts.append((lat / 1e5, lng / 1e5))
    return pts


# =============================================================================
# Steps
# =============================================================================
def step_geocode(g, locs):
    recs = []
    for r in tqdm(locs.itertuples(), total=len(locs), desc="geocode"):
        ok, d = g.call("geocode", "GET", URL["geocode"],
                       {"components": f"postal_code:{r.query_code}|country:{r.country}"})
        rec = {"location_key": r.location_key}
        res = d.get("results", []) if ok else []
        if not ok:
            rec["geocode_status"] = f"error: {d.get('error')}"
        elif not res:
            rec["geocode_status"] = "zero_results"
        else:
            postal = {"postal_code", "postal_code_prefix"}
            best = next((x for x in res if postal & set(x.get("types", []))), res[0])
            geo = best["geometry"]
            b = geo.get("bounds") or geo.get("viewport") or {}
            rec.update({
                "geocode_status": "ok" if postal & set(best.get("types", [])) else "non_postal_match",
                "lat": geo["location"]["lat"], "lng": geo["location"]["lng"],
                "partial_match": bool(best.get("partial_match")),
                "formatted_address": best.get("formatted_address"), "place_id": best.get("place_id"),
                "bounds_ne_lat": b.get("northeast", {}).get("lat"),
                "bounds_ne_lng": b.get("northeast", {}).get("lng"),
                "bounds_sw_lat": b.get("southwest", {}).get("lat"),
                "bounds_sw_lng": b.get("southwest", {}).get("lng"),
            })
        recs.append(rec)
    return pd.DataFrame(recs)


def step_elevation(g, pts):
    recs = []
    for i in range(0, len(pts), 200):
        chunk = pts.iloc[i:i + 200]
        ok, d = g.call("elevation", "GET", URL["elevation"],
                       {"locations": "|".join(f"{r.lat:.6f},{r.lng:.6f}" for r in chunk.itertuples())})
        res = d.get("results", []) if ok else []
        for j, r in enumerate(chunk.itertuples()):
            e = res[j] if j < len(res) else {}
            recs.append({"location_key": r.location_key, "elevation_m": e.get("elevation"),
                         "elevation_resolution_m": e.get("resolution")})
    return pd.DataFrame(recs)


# ---- images -----------------------------------------------------------------
def fetch_image(g, lat, lng, span, kind, out_path):
    """Fetch, crop to exactly `span` metres, resample, save. -> PIL image or None."""
    if out_path.exists():
        return Image.open(out_path).convert("RGB")
    zoom, frac = zoom_for_span(lat, span)
    params = {"center": f"{lat:.6f},{lng:.6f}", "zoom": zoom, "size": "640x640", "scale": 2,
              "format": "png", "maptype": "satellite" if kind == "satellite" else "roadmap"}
    if kind == "analysis":
        params["style"] = ANALYSIS_STYLE
    ok, data = g.call("static_maps", "GET", URL["static"], params, binary=True)
    if not ok:
        return None
    im = Image.open(io.BytesIO(data)).convert("RGB")
    w, h = im.size
    cw, ch = round(w * frac), round(h * frac)
    l, t = (w - cw) // 2, (h - ch) // 2
    im = im.crop((l, t, l + cw, t + ch)).resize(
        (OUTPUT_IMAGE_PX, OUTPUT_IMAGE_PX), Image.NEAREST if kind == "analysis" else Image.LANCZOS)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im.save(out_path)
    return im


def hsv_fractions(im):
    hsv = np.asarray(im.convert("HSV"), dtype=float) / 255.0
    hue, sat, val = hsv[..., 0] * 360, hsv[..., 1], hsv[..., 2]
    colored = (sat >= MIN_SAT) & (val >= MIN_VAL)
    green = colored & (hue >= GREEN_HUE[0]) & (hue <= GREEN_HUE[1])
    blue = colored & (hue >= BLUE_HUE[0]) & (hue <= BLUE_HUE[1])
    return round(100 * green.mean(), 3), round(100 * blue.mean(), 3)


def analysis_masks(im):
    a = np.asarray(im, dtype=float)
    names = list(PALETTE)
    cols = np.array([PALETTE[n] for n in names], dtype=float)
    d = np.linalg.norm(a[..., None, :] - cols[None, None], axis=-1)
    idx = d.argmin(-1)
    idx[d.min(-1) > PALETTE_TOL] = names.index("other")
    return {n: idx == i for i, n in enumerate(names)}


def nearest_m(mask, span):
    """Distance (m) from image centre to nearest True pixel, or None if not reliably inside
    the image (only distances within the inscribed circle are guaranteed to be the nearest)."""
    ys, xs = np.nonzero(mask)
    if not len(xs):
        return None
    c = (mask.shape[0] - 1) / 2
    dist = math.sqrt(((xs - c) ** 2 + (ys - c) ** 2).min()) * span / mask.shape[0]
    return dist if dist <= span / 2 else None


def step_images(g, pts, out_dir):
    img_rows, loc_rows = [], []
    for r in tqdm(pts.itertuples(), total=len(pts), desc="images"):
        found = {}  # feature -> (distance, span used)
        spans = [a * SPAN_FACTOR for a in ALTITUDES_M]
        for alt, span in zip(ALTITUDES_M, spans):
            row = {"location_key": r.location_key, "altitude_m": alt, "span_m": span,
                   "m_per_px": span / OUTPUT_IMAGE_PX, "zoom": zoom_for_span(r.lat, span)[0]}
            for kind in ("roadmap", "satellite", "analysis"):
                path = out_dir / "images" / r.location_key / f"{kind}_{alt}m.png"
                im = fetch_image(g, r.lat, r.lng, span, kind, path)
                row[f"{kind}_file"] = str(path.relative_to(out_dir)) if im else None
                if im is None:
                    continue
                if kind == "analysis":
                    m = analysis_masks(im)
                    row["analysis_green_pct"] = round(100 * m["green"].mean(), 3)
                    row["analysis_water_pct"] = round(100 * m["water"].mean(), 3)
                    row["analysis_highway_pct"] = round(100 * (m["highway"] | m["freeway"]).mean(), 3)
                    _distances(m, span, found)
                else:
                    row[f"{kind}_green_pct"], row[f"{kind}_blue_pct"] = hsv_fractions(im)
            img_rows.append(row)
        if len(found) < 3:   # something not within 5 km: one wider image
            span = DISTANCE_FALLBACK_SPAN_M
            path = out_dir / "images" / r.location_key / f"analysis_{span}m_distance.png"
            im = fetch_image(g, r.lat, r.lng, span, "analysis", path)
            if im is not None:
                _distances(analysis_masks(im), span, found)
        loc = {"location_key": r.location_key}
        for feat in ("water", "highway", "freeway"):
            d, span = found.get(feat, (None, None))
            loc[f"dist_{feat}_m"] = round(d, 1) if d is not None else None
            loc[f"dist_{feat}_resolution_m"] = round(span / OUTPUT_IMAGE_PX, 1) if span else None
            if d is None:
                loc[f"dist_{feat}_note"] = f"> {DISTANCE_FALLBACK_SPAN_M // 2} m or not fetched"
        loc_rows.append(loc)
    return pd.DataFrame(img_rows), pd.DataFrame(loc_rows)


def _distances(masks, span, found):
    for feat, mask in (("water", masks["water"]),
                       ("highway", masks["highway"] | masks["freeway"]),
                       ("freeway", masks["freeway"])):
        if feat not in found:
            d = nearest_m(mask, span)
            if d is not None:
                found[feat] = (d, span)


# ---- traffic ----------------------------------------------------------------
_tf = None


def local_tz(lat, lng):
    global _tf
    if _tf is None:
        from timezonefinder import TimezoneFinder
        _tf = TimezoneFinder()
    return ZoneInfo(_tf.timezone_at(lat=lat, lng=lng) or "UTC")


def departure_utc(anchor: date, weekday, hour, tz):
    """First `weekday` strictly after the anchor date, at `hour` local time, as RFC 3339 UTC."""
    days = (weekday - anchor.weekday()) % 7 or 7
    local = datetime.combine(anchor + timedelta(days=days), datetime.min.time()).replace(hour=hour, tzinfo=tz)
    return local.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), local.isoformat()


def traffic_level(ratio):
    if ratio is None or math.isnan(ratio):
        return None
    return 1 + sum(ratio > c for c in TRAFFIC_LEVEL_CUTS)


def step_traffic(g, pts, anchor):
    fields = ("routes.duration,routes.staticDuration,routes.distanceMeters,"
              "routes.polyline.encodedPolyline,routes.travelAdvisory.speedReadingIntervals")
    recs = []
    for r in tqdm(pts.itertuples(), total=len(pts), desc="traffic"):
        tz = local_tz(r.lat, r.lng)
        probes = {"WE": (offset(r.lat, r.lng, -TRAFFIC_PROBE_HALF_M, 0), offset(r.lat, r.lng, TRAFFIC_PROBE_HALF_M, 0)),
                  "SN": (offset(r.lat, r.lng, 0, -TRAFFIC_PROBE_HALF_M), offset(r.lat, r.lng, 0, TRAFFIC_PROBE_HALF_M))}
        for label, wd, hour in TRAFFIC_SLOTS:
            dep, dep_local = departure_utc(anchor, wd, hour, tz)
            dur = static = 0.0
            seg = Counter()
            n_ok = 0
            for (o, d) in probes.values():
                body = {"origin": {"location": {"latLng": {"latitude": o[0], "longitude": o[1]}}},
                        "destination": {"location": {"latLng": {"latitude": d[0], "longitude": d[1]}}},
                        "travelMode": "DRIVE", "routingPreference": "TRAFFIC_AWARE_OPTIMAL",
                        "departureTime": dep, "extraComputations": ["TRAFFIC_ON_POLYLINE"],
                        "polylineQuality": "HIGH_QUALITY"}
                ok, res = g.call("routes", "POST", URL["routes"], body=body,
                                 headers={"X-Goog-FieldMask": fields})
                routes = res.get("routes", []) if ok else []
                if not routes:
                    continue
                rt = routes[0]
                dur += float(rt.get("duration", "0s").rstrip("s"))
                static += float(rt.get("staticDuration", "0s").rstrip("s"))
                n_ok += 1
                pts_ = decode_polyline(rt.get("polyline", {}).get("encodedPolyline", ""))
                for iv in rt.get("travelAdvisory", {}).get("speedReadingIntervals", []):
                    a, b = iv.get("startPolylinePointIndex", 0), iv.get("endPolylinePointIndex", 0)
                    seg[iv.get("speed", "UNKNOWN")] += sum(
                        haversine(pts_[k], pts_[k + 1]) for k in range(a, min(b, len(pts_) - 1)))
            ratio = dur / static if static > 0 else float("nan")
            tot = sum(seg.values()) or float("nan")
            recs.append({"location_key": r.location_key, "slot": f"{label}_{hour:02d}00",
                         "departure_local": dep_local, "probe_routes_ok": n_ok,
                         "duration_s": dur or None, "free_flow_s": static or None,
                         "congestion_ratio": round(ratio, 3) if not math.isnan(ratio) else None,
                         "traffic_level_1to4": traffic_level(ratio),
                         "pct_normal": round(100 * seg["NORMAL"] / tot, 1) if seg else None,
                         "pct_slow": round(100 * seg["SLOW"] / tot, 1) if seg else None,
                         "pct_jam": round(100 * seg["TRAFFIC_JAM"] / tot, 1) if seg else None})
    return pd.DataFrame(recs)


# ---- air quality ------------------------------------------------------------
AQ_EXTRA = ["LOCAL_AQI", "POLLUTANT_CONCENTRATION", "DOMINANT_POLLUTANT_CONCENTRATION"]


def _aq_flat(info):
    row = {}
    for ix in info.get("indexes", []):
        c = ix.get("code", "idx")
        row[f"aqi_{c}"] = ix.get("aqi")
        row[f"aqi_{c}_category"] = ix.get("category")
        row[f"aqi_{c}_dominant"] = ix.get("dominantPollutant")
    for p in info.get("pollutants", []):
        conc = p.get("concentration", {})
        row[f"{p.get('code')}_{conc.get('units', '')}".lower()] = conc.get("value")
    return row


def step_airquality(g, pts, tag):
    cur, hist = [], []
    for r in tqdm(pts.itertuples(), total=len(pts), desc="air quality"):
        loc = {"latitude": r.lat, "longitude": r.lng}
        ok, d = g.call("air_quality", "POST", URL["aq_current"], tag=tag, body={
            "location": loc, "universalAqi": True, "extraComputations": AQ_EXTRA, "languageCode": "en"})
        cur.append({"location_key": r.location_key, "aq_status": "ok" if ok else d.get("error"),
                    "dateTime": d.get("dateTime") if ok else None, **(_aq_flat(d) if ok else {})})
        token = None
        while True:
            body = {"location": loc, "hours": AQ_HISTORY_HOURS, "pageSize": 168, "universalAqi": True,
                    "extraComputations": AQ_EXTRA, "languageCode": "en"}
            if token:
                body["pageToken"] = token
            ok, d = g.call("air_quality", "POST", URL["aq_history"], body=body, tag=tag)
            if not ok:
                break
            for h in d.get("hoursInfo", []):
                hist.append({"location_key": r.location_key, "dateTime": h.get("dateTime"), **_aq_flat(h)})
            token = d.get("nextPageToken")
            if not token:
                break
    return pd.DataFrame(cur), pd.DataFrame(hist)


# ---- pollen -----------------------------------------------------------------
def step_pollen(g, pts, tag):
    rows = []
    for r in tqdm(pts.itertuples(), total=len(pts), desc="pollen"):
        ok, d = g.call("pollen", "GET", URL["pollen"], tag=tag, params={
            "location.latitude": r.lat, "location.longitude": r.lng,
            "days": POLLEN_DAYS, "plantsDescription": "false"})
        if not ok:
            rows.append({"location_key": r.location_key, "pollen_status": d.get("error")})
            continue
        for day in d.get("dailyInfo", []):
            dt = day.get("date", {})
            row = {"location_key": r.location_key, "pollen_status": "ok",
                   "date": f"{dt.get('year')}-{dt.get('month', 0):02d}-{dt.get('day', 0):02d}"}
            for kind, key in (("type", "pollenTypeInfo"), ("plant", "plantInfo")):
                for p in day.get(key, []):
                    c = p.get("code", "?").lower()
                    row[f"{kind}_{c}_index"] = p.get("indexInfo", {}).get("value")
                    row[f"{kind}_{c}_category"] = p.get("indexInfo", {}).get("category")
                    if kind == "type":
                        row[f"{kind}_{c}_in_season"] = p.get("inSeason")
            rows.append(row)
    return pd.DataFrame(rows)


# ---- weather ----------------------------------------------------------------
def _paged_get(g, api, url, params, list_key, tag):
    out, token = [], None
    while True:
        p = dict(params, **({"pageToken": token} if token else {}))
        ok, d = g.call(api, "GET", url, p, tag=tag)
        if not ok:
            return out, d.get("error")
        out += d.get(list_key, [])
        token = d.get("nextPageToken")
        if not token:
            return out, None


def step_weather(g, pts, tag):
    cur, hist, fc = [], [], []
    for r in tqdm(pts.itertuples(), total=len(pts), desc="weather"):
        base = {"location.latitude": r.lat, "location.longitude": r.lng, "unitsSystem": "METRIC"}
        ok, d = g.call("weather", "GET", URL["wx_current"], base, tag=tag)
        cur.append({"location_key": r.location_key, "weather_status": "ok" if ok else d.get("error"),
                    **(pd.json_normalize(d).iloc[0].to_dict() if ok and d else {})})
        hours, _ = _paged_get(g, "weather", URL["wx_history"],
                              {**base, "hours": WEATHER_HISTORY_HOURS, "pageSize": 24}, "historyHours", tag)
        for h in hours:
            hist.append({"location_key": r.location_key, **pd.json_normalize(h).iloc[0].to_dict()})
        days, _ = _paged_get(g, "weather", URL["wx_forecast"],
                             {**base, "days": WEATHER_FORECAST_DAYS, "pageSize": 10}, "forecastDays", tag)
        for dd in days:
            fc.append({"location_key": r.location_key, **pd.json_normalize(dd).iloc[0].to_dict()})
    return pd.DataFrame(cur), pd.DataFrame(hist), pd.DataFrame(fc)


# ---- solar ------------------------------------------------------------------
def step_solar(g, pts):
    rows = []
    for r in tqdm(pts.itertuples(), total=len(pts), desc="solar"):
        ok, d = g.call("solar", "GET", URL["solar"], {
            "location.latitude": r.lat, "location.longitude": r.lng, "requiredQuality": SOLAR_QUALITY})
        if not ok:
            rows.append({"location_key": r.location_key, "solar_status": d.get("error")})
            continue
        sp, c, img = d.get("solarPotential", {}), d.get("center", {}), d.get("imageryDate", {})
        rows.append({
            "location_key": r.location_key, "solar_status": "ok",
            "solar_building_distance_m": round(haversine((r.lat, r.lng), (c["latitude"], c["longitude"])), 1)
            if c else None,
            "solar_imagery_quality": d.get("imageryQuality"),
            "solar_imagery_date": f"{img.get('year')}-{img.get('month')}-{img.get('day')}" if img else None,
            "max_sunshine_hours_per_year": sp.get("maxSunshineHoursPerYear"),
            "max_array_panels_count": sp.get("maxArrayPanelsCount"),
            "max_array_area_m2": sp.get("maxArrayAreaMeters2"),
            "whole_roof_area_m2": sp.get("wholeRoofStats", {}).get("areaMeters2"),
            "carbon_offset_kg_per_mwh": sp.get("carbonOffsetFactorKgPerMwh"),
        })
    return pd.DataFrame(rows)


# =============================================================================
# Summary table (one row per location)
# =============================================================================
def summarize(out_dir, geo, tables):
    s = geo.copy()

    def add(df):
        nonlocal s
        if df is not None and len(df):
            s = s.merge(df, on="location_key", how="left")

    add(tables.get("elevation"))
    add(tables.get("distances"))
    img = tables.get("images")
    if img is not None and len(img):
        cols = [c for c in img.columns if c.endswith("_pct")]
        w = img.pivot(index="location_key", columns="altitude_m", values=cols)
        w.columns = [f"{c}_{a}m" for c, a in w.columns]
        add(w.reset_index())
    tr = tables.get("traffic")
    if tr is not None and len(tr):
        w = tr.pivot(index="location_key", columns="slot", values=["traffic_level_1to4", "congestion_ratio"])
        w.columns = [f"{v}_{slot}" for v, slot in w.columns]
        add(w.reset_index())
    aqh = tables.get("air_quality_hourly")
    if aqh is not None and len(aqh):
        num = aqh.drop(columns=["dateTime"]).apply(pd.to_numeric, errors="coerce")
        num["location_key"] = aqh["location_key"]
        num = num.dropna(axis=1, how="all")
        agg = num.groupby("location_key").agg(["mean", "max"])
        agg.columns = [f"aq30d_{c}_{f}" for c, f in agg.columns]
        agg["aq30d_hours"] = aqh.groupby("location_key").size()
        add(agg.reset_index())
    add(tables.get("solar"))
    return s


# =============================================================================
# Main
# =============================================================================
def estimate(n):
    per = {"geocode": 1, "elevation": 1 / 200, "static_maps": 3 * len(ALTITUDES_M) + 1,
           "routes": 2 * len(TRAFFIC_SLOTS), "air_quality": 1 + math.ceil(AQ_HISTORY_HOURS / 168),
           "pollen": 1, "weather": 3, "solar": 1}
    print(f"\n{n} unique locations. Upper-bound request counts:")
    for api, k in per.items():
        print(f"  {api:<12} {math.ceil(n * k):>8,}")
    print("Check prices / free monthly usage at https://mapsplatform.google.com/pricing/\n"
          "Routes calls use traffic + TRAFFIC_ON_POLYLINE (Pro/Enterprise SKU).")


def load_env_file():
    p = Path(__file__).with_name(".env")
    if p.exists():
        for line in p.read_text().splitlines():
            if "=" in line and not line.lstrip().startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def check_apis(key, session=None):
    """One request per API at a public location (SickKids, Toronto). No patient data, no cache."""
    s = session or requests.Session()
    lat, lng = 43.6573, -79.3879
    loc_q = {"location.latitude": lat, "location.longitude": lng}
    loc_b = {"location": {"latitude": lat, "longitude": lng}}
    ll = {"latLng": {"latitude": lat, "longitude": lng}}
    o2 = offset(lat, lng, 800, 0)
    tests = [
        ("Geocoding", "GET", URL["geocode"], {"components": "postal_code:M5G|country:CA"}, None),
        ("Maps Static", "GET", URL["static"], {"center": f"{lat},{lng}", "zoom": 15, "size": "100x100"}, None),
        ("Routes", "POST", URL["routes"], None, {
            "origin": {"location": ll},
            "destination": {"location": {"latLng": {"latitude": o2[0], "longitude": o2[1]}}},
            "travelMode": "DRIVE", "routingPreference": "TRAFFIC_AWARE_OPTIMAL",
            "extraComputations": ["TRAFFIC_ON_POLYLINE"]}),
        ("Air Quality", "POST", URL["aq_current"], None, loc_b),
        ("Pollen", "GET", URL["pollen"], {**loc_q, "days": 1}, None),
    ]
    print(f"\nChecking the key against {len(tests)} APIs (one request each):\n")
    n_ok = 0
    for name, method, url, params, body in tests:
        params, headers = dict(params or {}), {}
        if "maps.googleapis.com/maps/api" in url:
            params["key"] = key
        else:
            headers["X-Goog-Api-Key"] = key
        if name == "Routes":
            headers["X-Goog-FieldMask"] = "routes.duration,routes.staticDuration"
        try:
            r = s.request(method, url, params=params, json=body, headers=headers, timeout=30)
            if r.status_code < 400 and name == "Maps Static":
                ok = r.headers.get("Content-Type", "").startswith("image/")
                msg = "image received" if ok else r.text[:150]
            elif r.status_code < 400:
                d = r.json() if r.content else {}
                st = d.get("status") if isinstance(d, dict) else None
                ok = st in (None, "OK", "ZERO_RESULTS")
                msg = "OK" if ok else f"{st}: {d.get('error_message', '')}"
            else:
                ok, msg = False, f"HTTP {r.status_code} {_err_text(r)}"
        except requests.RequestException as e:
            ok, msg = False, f"network error: {type(e).__name__}"
        n_ok += ok
        print(f"  {'OK  ' if ok else 'FAIL'}  {name:<12} {msg[:160]}")
    print(f"\n{n_ok}/{len(tests)} APIs working.")
    if n_ok < len(tests):
        print("For FAIL lines: ask IT to enable that API on the key's Google Cloud project\n"
              "(and to allow it in the key's 'API restrictions' if the key is restricted).")
    return n_ok == len(tests)


def _vertical(rows):
    return pd.DataFrame(rows, columns=["Variable", "Value", "Unit", "Description"])


def run_sample(key, postal, out, session=None):
    """Fetch ONE location and write one sample workbook per requester:
    out/Angelika/sample_Angelika.xlsx (+ images) and out/Hilary/sample_Hilary.xlsx."""
    country, code, problem = normalize_postal(postal)
    if problem:
        raise SystemExit(f"--sample: {problem}")
    query = code[:3] if (country == "CA" and CA_LEVEL == "fsa") else code
    locs = pd.DataFrame([{"location_key": f"{country}_{query}", "country": country,
                          "query_code": query, "fsa_region": None, "n_patients": 0}])
    out.mkdir(parents=True, exist_ok=True)
    g = Google(key, out / "cache", session=session)
    today = date.today()
    geo = locs.merge(step_geocode(g, locs), on="location_key")
    if "lat" not in geo or geo["lat"].isna().all():
        raise SystemExit(f"Could not geocode {postal}: {geo.get('geocode_status', ['?'])[0]}")
    loc = geo.iloc[0]
    where = [("Postal code queried", query, "", f"{'FSA' if len(query) == 3 else 'postal code'}, {country}"),
             ("Google match", loc.get("formatted_address"), "", "Area Google matched"),
             ("Latitude", round(loc["lat"], 6), "deg", "Centre of the postal code area"),
             ("Longitude", round(loc["lng"], 6), "deg", "Centre of the postal code area"),
             ("Fetched on", today.isoformat(), "", "")]

    # ---------------- Angelika: images, green/blue, distances, traffic ----------------
    a_dir = out / "Angelika"
    imgs, dists = step_images(g, geo, a_dir)
    traffic = step_traffic(g, geo, today)
    rows = list(where)
    for r in imgs.itertuples():
        for kind, label in (("satellite", "Satellite"), ("roadmap", "Map")):
            rows.append((f"{label} {r.altitude_m:,} m: % green", getattr(r, f"{kind}_green_pct", None), "%",
                         f"Share of pixels in the green spectrum ({label.lower()} image)"))
            rows.append((f"{label} {r.altitude_m:,} m: % blue", getattr(r, f"{kind}_blue_pct", None), "%",
                         f"Share of pixels in the blue spectrum ({label.lower()} image)"))
    d = dists.iloc[0] if len(dists) else {}
    rows += [("Distance to nearest water", d.get("dist_water_m"), "m", "Lakes, rivers, sea on Google's map"),
             ("Distance to nearest major highway", d.get("dist_highway_m"), "m", "Google 'highway' road class"),
             ("Distance to nearest freeway", d.get("dist_freeway_m"), "m", "Controlled-access highways only")]
    for t in traffic.itertuples():
        rows.append((f"Typical traffic {t.slot.replace('_', ' ')}", t.traffic_level_1to4, "1-4",
                     "1 fast, 2 moderate, 3 slow, 4 heavy (traffic vs free-flow travel time)"))
    a_notes = _vertical([
        ("Images", "", "", f"images/{loc.location_key}/: map, satellite and 'analysis' PNG per altitude, "
                           "no markers. 'Altitude' = side of the ground square shown (1,000 m = 1 km x 1 km)."),
        ("% green / % blue", "", "", "Pixel colour (hue) classification of each image."),
        ("Distances", "", "", "Straight-line distance from the postal code centre, measured on Google's map."),
        ("Traffic", "", "", "Google's typical traffic for that weekday/time on two 800 m test routes through "
                            "the postal code centre; 4-level scale from the delay ratio."),
    ])
    with pd.ExcelWriter(a_dir / "sample_Angelika.xlsx") as xw:
        _vertical(rows).to_excel(xw, sheet_name="Summary", index=False)
        imgs.to_excel(xw, sheet_name="Images", index=False)
        traffic.to_excel(xw, sheet_name="Traffic detail", index=False)
        a_notes.to_excel(xw, sheet_name="Notes", index=False)

    # ---------------- Hilary: air quality, pollen (+ solar/weather status) ----------------
    h_dir = out / "Hilary"
    h_dir.mkdir(parents=True, exist_ok=True)
    aq_cur, aq_hist = step_airquality(g, geo, today.isoformat())
    pollen = step_pollen(g, geo, today.isoformat())
    rows = list(where)
    if len(aq_cur):
        c = aq_cur.iloc[0]
        rows += [("Air quality now: Universal AQI", c.get("aqi_uaqi"), "0-100", "Higher = better air"),
                 ("Air quality now: category", c.get("aqi_uaqi_category"), "", ""),
                 ("Air quality now: dominant pollutant", c.get("aqi_uaqi_dominant"), "", "")]
    if len(aq_hist):
        num = aq_hist.drop(columns=["location_key", "dateTime"]).apply(pd.to_numeric, errors="coerce")
        num = num.dropna(axis=1, how="all")
        span = f"{aq_hist['dateTime'].min()[:10]} to {aq_hist['dateTime'].max()[:10]}"
        for col in num.columns:
            rows.append((f"Last 30 days mean: {col}", round(num[col].mean(), 2), "", span))
            rows.append((f"Last 30 days max: {col}", round(num[col].max(), 2), "", span))
    if len(pollen) and "date" in pollen:
        p = pollen.iloc[0]
        for c in [c for c in pollen.columns if c.startswith("type_") and c.endswith("_index")]:
            rows.append((f"Pollen today: {c[5:-6]}", p.get(c), "0-5 index", f"UPI, date {p['date']}"))
    rows += [("Solar", "not collected", "", "Google Solar API = rooftop solar-panel potential of one building; "
                                           "sunlight/UV by date is available from NASA POWER"),
             ("Weather", "not collected", "", "Google Weather API = last 24 h + forecast only; historical "
                                             "weather by date is available from Open-Meteo / ERA5")]
    h_notes = _vertical([
        ("Time period", "", "", "Air quality = the 30 days before the fetch date; pollen = forecast from the "
                                "fetch date. Google has no air quality or pollen history before that, so these "
                                "are current conditions, not exposure before each patient's event."),
    ])
    with pd.ExcelWriter(h_dir / "sample_Hilary.xlsx") as xw:
        _vertical(rows).to_excel(xw, sheet_name="Summary", index=False)
        aq_hist.to_excel(xw, sheet_name="Air quality hourly 30d", index=False)
        aq_cur.to_excel(xw, sheet_name="Air quality now", index=False)
        pollen.to_excel(xw, sheet_name="Pollen 5-day", index=False)
        h_notes.to_excel(xw, sheet_name="Notes", index=False)

    log.info("requests sent: %s | cache hits: %s", dict(g.sent), dict(g.hits))
    print(f"\nSamples written:\n  {a_dir / 'sample_Angelika.xlsx'}  (+ {a_dir / 'images'})"
          f"\n  {h_dir / 'sample_Hilary.xlsx'}")


def main(argv=None, session=None):
    if argv is None:
        argv = sys.argv[1:]
    if "--sample" in argv:
        sp = argparse.ArgumentParser(description="One-location sample for Angelika and Hilary")
        sp.add_argument("--sample", nargs="?", const="M5G 1X8", default="M5G 1X8",
                        help="postal code / ZIP to sample (default M5G 1X8, SickKids)")
        sp.add_argument("--out", default="sample_output")
        sa = sp.parse_args(argv)
        logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
        load_env_file()
        key = os.environ.get("GOOGLE_MAPS_API_KEY")
        if not key:
            raise SystemExit("Set GOOGLE_MAPS_API_KEY (environment variable or .env next to the script).")
        run_sample(key, sa.sample, Path(sa.out), session)
        return 0
    if "--check-apis" in argv:
        load_env_file()
        key = os.environ.get("GOOGLE_MAPS_API_KEY")
        if not key:
            raise SystemExit("Set GOOGLE_MAPS_API_KEY (environment variable or .env next to the script).")
        return 0 if check_apis(key, session) else 1
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--check-apis", action="store_true",
                    help="send one request to each API at a public location and report which work")
    ap.add_argument("--input", required=True, help="ikdr_all_postal codes xlsx/csv (IKDR_REGID, ZIP)")
    ap.add_argument("--fsa-regions", help="FSA-merge-Info.xlsx (adds region to Canadian patients)")
    ap.add_argument("--out", default="output")
    ap.add_argument("--ca-level", choices=["fsa", "full"], default=CA_LEVEL)
    ap.add_argument("--steps", default=",".join(DEFAULT_STEPS),
                    help=f"comma list from {STEPS} (default: {DEFAULT_STEPS}; geocode always runs)")
    ap.add_argument("--limit", type=int, help="only the first N locations (for a test run)")
    ap.add_argument("--estimate", action="store_true", help="count locations and requests, call nothing")
    ap.add_argument("--new-snapshot", action="store_true",
                    help="start a new date snapshot for time-dependent APIs (traffic, AQ, pollen, weather)")
    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    patients, locs = build_locations(a.input, a.fsa_regions, a.ca_level)
    if a.limit:
        locs = locs.head(a.limit)
    log.info("%d patients, %d with a usable postal code, %d unique locations",
             len(patients), patients["location_key"].notna().sum(), len(locs))
    bad = patients["postal_problem"].value_counts()
    if len(bad):
        log.info("postal code problems: %s", bad.to_dict())
    if a.estimate:
        estimate(len(locs))
        return

    load_env_file()
    key = os.environ.get("GOOGLE_MAPS_API_KEY")
    if not key:
        raise SystemExit("Set GOOGLE_MAPS_API_KEY (environment variable or .env next to the script).")
    steps = [s.strip() for s in a.steps.split(",") if s.strip()]
    unknown = set(steps) - set(STEPS)
    if unknown:
        raise SystemExit(f"unknown steps: {unknown}")

    # Snapshot date: fixes traffic departure days and the cache for "last N days" APIs,
    # so a re-run on another day reuses what was already fetched.
    info_p = out / "run_info.json"
    info = json.loads(info_p.read_text()) if info_p.exists() else {}
    if a.new_snapshot or "snapshot_date" not in info:
        info["snapshot_date"] = date.today().isoformat()
    anchor = date.fromisoformat(info["snapshot_date"])
    tag = info["snapshot_date"]

    g = Google(key, out / "cache", session=session)
    tables = {}
    geo = step_geocode(g, locs)          # always (cached after the first run)
    geo = locs.merge(geo, on="location_key", how="left")
    pts = geo.dropna(subset=["lat", "lng"]) if "lat" in geo else geo.iloc[0:0]
    log.info("geocoded %d / %d locations", len(pts), len(geo))

    if "elevation" in steps:
        tables["elevation"] = step_elevation(g, pts)
    if "images" in steps:
        tables["images"], tables["distances"] = step_images(g, pts, out)
    if "traffic" in steps:
        tables["traffic"] = step_traffic(g, pts, anchor)
    if "airquality" in steps:
        tables["air_quality_current"], tables["air_quality_hourly"] = step_airquality(g, pts, tag)
    if "pollen" in steps:
        tables["pollen_daily"] = step_pollen(g, pts, tag)
    if "weather" in steps:
        tables["weather_current"], tables["weather_hourly_24h"], tables["weather_daily_forecast"] = \
            step_weather(g, pts, tag)
    if "solar" in steps:
        tables["solar"] = step_solar(g, pts)

    # ---- write ----
    summary = summarize(out, geo, tables)
    tables_dir = out / "tables"
    tables_dir.mkdir(exist_ok=True)
    for name, df in tables.items():
        if df is not None:
            df.to_csv(tables_dir / f"{name}.csv", index=False)
    summary.to_csv(out / "locations_summary.csv", index=False)
    # Patient-level file (contains IDs: keep it local, do not share outside the study team)
    patients.merge(summary.drop(columns=["country", "query_code", "fsa_region", "n_patients"], errors="ignore"),
                   on="location_key", how="left").to_csv(out / "patients_with_environment.csv", index=False)
    with pd.ExcelWriter(out / "google_env_results.xlsx") as xw:
        summary.to_excel(xw, sheet_name="locations_summary", index=False)
        for name in ("images", "traffic", "solar", "air_quality_current", "pollen_daily"):
            if name in tables and len(tables[name]) < 1_000_000:
                tables[name].to_excel(xw, sheet_name=name[:31], index=False)

    info.update({"last_run_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                 "steps": steps, "locations": len(locs), "ca_level": a.ca_level,
                 "altitudes_m": ALTITUDES_M, "span_factor": SPAN_FACTOR,
                 "requests_sent": dict(g.sent), "cache_hits": dict(g.hits)})
    info_p.write_text(json.dumps(info, indent=2))
    log.info("requests sent: %s | cache hits: %s", dict(g.sent), dict(g.hits))
    log.info("done -> %s", out.resolve())


if __name__ == "__main__":
    sys.exit(main())
