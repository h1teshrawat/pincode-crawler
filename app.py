import io
import os
import re
import time
import uuid

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Font

load_dotenv()

app = Flask(__name__)

TOMTOM_API_KEY = os.environ.get("TOMTOM_API_KEY", "")
TOMTOM_BASE = "https://api.tomtom.com"
REQUEST_TIMEOUT = 6
INSTAGRAM_RE = re.compile(
    r"(?:https?://)?(?:www\.)?instagram\.com/([A-Za-z0-9_.]+)", re.IGNORECASE
)
IGNORED_INSTAGRAM_PATHS = {"p", "reel", "reels", "tv", "explore", "accounts", "share"}

# In-memory cache of generated workbooks, keyed by a one-time download id.
# Fine for a single-instance hobby deployment; entries are pruned by age.
RESULTS_CACHE: dict[str, tuple[float, bytes, str]] = {}
CACHE_TTL_SECONDS = 30 * 60


def _prune_cache():
    cutoff = time.time() - CACHE_TTL_SECONDS
    for key in [k for k, (ts, _, _) in RESULTS_CACHE.items() if ts < cutoff]:
        RESULTS_CACHE.pop(key, None)


def geocode_pincode(pincode: str, country: str | None):
    params = {
        "key": TOMTOM_API_KEY,
        "limit": 5,
        "entityTypeSet": "PostalCodeArea",
    }
    if country:
        params["countrySet"] = country
    resp = requests.get(
        f"{TOMTOM_BASE}/search/2/geocode/{requests.utils.quote(pincode)}.json",
        params=params,
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    data = resp.json()
    results = data.get("results", [])
    if not results:
        # Retry without the PostalCodeArea restriction as a fallback.
        params.pop("entityTypeSet", None)
        resp = requests.get(
            f"{TOMTOM_BASE}/search/2/geocode/{requests.utils.quote(pincode)}.json",
            params=params,
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        results = resp.json().get("results", [])
    if not results:
        return None
    pos = results[0]["position"]
    return pos["lat"], pos["lon"], results[0].get("address", {}).get("freeformAddress", "")


def search_pois(query: str, lat: float, lon: float, radius: int, limit: int):
    params = {
        "key": TOMTOM_API_KEY,
        "lat": lat,
        "lon": lon,
        "radius": radius,
        "limit": min(limit, 100),
    }
    resp = requests.get(
        f"{TOMTOM_BASE}/search/2/poiSearch/{requests.utils.quote(query)}.json",
        params=params,
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    return resp.json().get("results", [])


def find_instagram_handle(website: str) -> str:
    if not website:
        return ""
    try:
        resp = requests.get(
            website,
            timeout=REQUEST_TIMEOUT,
            headers={"User-Agent": "Mozilla/5.0 (compatible; PincodeCrawler/1.0)"},
        )
        html = resp.text
    except requests.RequestException:
        return ""

    for match in INSTAGRAM_RE.finditer(html):
        handle = match.group(1).strip("/")
        if handle.lower() not in IGNORED_INSTAGRAM_PATHS and handle:
            return handle

    # Some sites only link Instagram from the homepage's <head> or footer via
    # a relative/JS-rendered widget; a quick BeautifulSoup pass over <a> tags
    # catches a few of those the raw regex on full HTML misses.
    try:
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            m = INSTAGRAM_RE.search(a["href"])
            if m:
                handle = m.group(1).strip("/")
                if handle.lower() not in IGNORED_INSTAGRAM_PATHS and handle:
                    return handle
    except Exception:
        pass

    return ""


def build_workbook(rows: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    headers = ["Name", "Category", "Phone", "Website", "Instagram", "Address", "Distance (m)"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([
            row["name"],
            row["category"],
            row["phone"],
            row["website"],
            f"@{row['instagram']}" if row["instagram"] else "",
            row["address"],
            row["distance"],
        ])

    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/search", methods=["POST"])
def search():
    if not TOMTOM_API_KEY:
        return jsonify(error="Server is missing TOMTOM_API_KEY. Set it in your environment."), 500

    data = request.get_json(silent=True) or request.form
    pincode = (data.get("pincode") or "").strip()
    query = (data.get("query") or "").strip()
    country = (data.get("country") or "").strip().upper() or None
    radius = int(data.get("radius", 5000))
    limit = int(data.get("limit", 50))
    fetch_instagram = str(data.get("fetch_instagram", "")).lower() in ("on", "true", "1")

    if not pincode or not query:
        return jsonify(error="Both pincode and business type are required."), 400

    geo = geocode_pincode(pincode, country)
    if not geo:
        return jsonify(error=f"Could not find location for pincode '{pincode}'."), 404
    lat, lon, area_name = geo

    pois = search_pois(query, lat, lon, radius, limit)

    rows = []
    for item in pois:
        poi = item.get("poi", {})
        address = item.get("address", {})
        website = poi.get("url", "")
        instagram = ""
        if fetch_instagram and website:
            instagram = find_instagram_handle(website)
            time.sleep(0.2)  # be polite to target sites
        rows.append({
            "name": poi.get("name", ""),
            "category": ", ".join(poi.get("categories", [])),
            "phone": poi.get("phone", ""),
            "website": website,
            "instagram": instagram,
            "address": address.get("freeformAddress", ""),
            "distance": round(item.get("dist", 0)) if item.get("dist") else "",
        })

    _prune_cache()
    safe_query = re.sub(r"[^A-Za-z0-9]+", "_", query).strip("_") or "results"
    filename = f"{safe_query}_{pincode}.xlsx"
    download_id = uuid.uuid4().hex
    RESULTS_CACHE[download_id] = (time.time(), build_workbook(rows), filename)

    return jsonify(area=area_name, rows=rows, download_id=download_id, count=len(rows))


@app.route("/download/<download_id>")
def download(download_id):
    entry = RESULTS_CACHE.get(download_id)
    if not entry:
        return "This results link has expired. Please run the search again.", 404
    _, workbook_bytes, filename = entry
    return send_file(
        io.BytesIO(workbook_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
