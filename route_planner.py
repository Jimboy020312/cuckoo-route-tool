"""
Cuckoo+ Service Specialist — address grouping / map

Run this AFTER main.py (the scraper) has produced cuckoo_export.xlsm (or
.xlsx if you haven't set up the WhatsApp macro template). It doesn't touch
anything the scraper wrote — it only ADDS to the same file, plus creates
one new standalone file:

  1. Geocodes each customer's install_address using OpenStreetMap's free
     Nominatim service (no API key, no account, no cost). Results are
     cached locally in geocode_cache.json, so re-running this script
     after adding a few new customers doesn't re-geocode addresses it's
     already resolved.
  2. Groups customers by natural proximity — any two addresses within
     CLUSTER_RADIUS_KM of each other end up in the same "Area", and that
     chains transitively (so an area can be bigger than one circle if
     addresses form a connected string). There's no target group size
     and no forced day-count — areas form organically based on how
     close things actually are, and how many houses you tackle in a
     day, and in what order, is entirely up to you.
  3. Writes one new "Area" column onto the existing "Cuckoo Export"
     sheet — added at the END (column K) specifically so nothing
     already there (the WhatsApp Message/Link formulas, the macro's
     hidden helper column) has to move or get renumbered.
  4. Adds a new "Route Plan" sheet: one block per area, listing its
     stops with a Google Maps link per address — for looking one up
     individually, not for a route in any particular order.
  5. Writes route_map.html — an interactive map (OpenStreetMap tiles,
     via the free Leaflet.js library, no API key) with every address
     plotted, colored by area, so you can actually SEE where things are
     and decide the day-by-day split and order yourself. Open it in any
     browser.

ADDRESSES THAT DON'T GEOCODE
-----------------------------
Some addresses (especially very detailed ones — specific block/unit
numbers) won't resolve on the first try. This script automatically
retries using just the postcode if the full address fails. If that
still fails, the customer is listed separately at the bottom of the
Route Plan sheet under "Could not place automatically", and left off
the map — you'll need to place those manually.

NOMINATIM USAGE NOTE
-----------------------------
Nominatim's usage policy asks for a maximum of 1 request/second and an
identifying User-Agent — both handled below (NOMINATIM_USER_AGENT,
REQUEST_DELAY_SECONDS). If you plan on running this regularly, consider
editing NOMINATIM_USER_AGENT to include a real contact (e.g. your
email) per their policy, though for a once-a-month run of a few dozen
addresses this is very low volume.
"""

import json
import math
import os
import re
import time
import urllib.parse
import urllib.request

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# CONFIG
# ============================================================

# Same filenames main.py uses — this script looks for the macro-enabled
# one first (so it doesn't clobber the macro), falling back to plain xlsx.
OUTPUT_FILE_XLSM = "cuckoo_export.xlsm"
OUTPUT_FILE_XLSX = "cuckoo_export.xlsx"

GEOCODE_CACHE_FILE = "geocode_cache.json"
MAP_HTML_FILE = "route_map.html"

# Two addresses within this distance of each other land in the same
# Area — and that chains, so an area can span further than one circle
# if addresses form a connected string. Raise this for fewer, bigger
# areas; lower it for more, smaller ones.
CLUSTER_RADIUS_KM = 2.0

# A geocoding match is rejected as "too broad to be useful" if its
# bounding box spans more than this many degrees in either direction
# (roughly 0.05° ≈ 5.5km at Malaysia's latitude). This is what stops a
# detailed address that fails to parse from silently falling back to a
# whole-city or whole-state match — without this, many different
# addresses in the same city could all collapse onto one identical
# point instead of being honestly treated as failures.
MAX_MATCH_SPAN_DEGREES = 0.05

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
# Nominatim's usage policy asks for an identifying User-Agent. Feel free
# to edit this to include a real contact if you run this often.
NOMINATIM_USER_AGENT = "cuckoo-plus-route-planner/1.0 (personal workflow tool)"
# Their policy caps usage at 1 request/second — this stays comfortably
# under that.
REQUEST_DELAY_SECONDS = 1.1

AREA_COLORS = [
    "#1f4e78", "#c00000", "#2e7d32", "#e65100", "#6a1b9a",
    "#00838f", "#ad1457", "#4e342e", "#546e7a", "#9e9d24",
]


# ============================================================
# Geocoding (OpenStreetMap Nominatim, free, cached)
# ============================================================

def load_geocode_cache():
    if not os.path.exists(GEOCODE_CACHE_FILE):
        return {}
    try:
        with open(GEOCODE_CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_geocode_cache(cache):
    with open(GEOCODE_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def _nominatim_query(params):
    """One rate-limited request to Nominatim. Returns the raw first
    result dict (with lat/lon/boundingbox/etc.) or None."""
    url = NOMINATIM_URL + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": NOMINATIM_USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            results = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"    [geocode] request failed: {type(e).__name__}: {e}")
        return None
    if not results:
        return None
    return results[0]


def _is_specific_enough(result):
    """
    Rejects matches that are too broad to be useful — e.g. when a
    detailed unit/block address doesn't parse and Nominatim quietly
    falls back to matching just the city or state name. Left
    unchecked, that gives every address in the same city the SAME
    coordinates (exactly the "all the same coordinates" bug), which
    silently produces a garbage grouping instead of an honest failure.

    Uses the result's boundingbox width/height as a proxy for how
    precise the match is — a specific address's bounding box is tiny;
    a city or state's is not.
    """
    bbox = result.get("boundingbox")
    if not bbox:
        return False
    south, north, west, east = (float(x) for x in bbox)
    return (north - south) <= MAX_MATCH_SPAN_DEGREES and (east - west) <= MAX_MATCH_SPAN_DEGREES


def _simplify_to_street_level(address):
    """
    Malaysian residential addresses often lead with a unit/block/lot
    code (e.g. "A-T07-U06, BLOK A ZONE 6B ...") that essentially never
    exists in OpenStreetMap's data — there's no realistic way to
    geocode down to a specific unit, and feeding Nominatim the whole
    noisy string often makes the full-address query fail entirely.

    This strips everything before the first recognizable street/area
    keyword (JALAN, PERSIARAN, LORONG, PRESINT, TAMAN, BANDAR, LEBUH,
    LINGKARAN — all common in Malaysian addresses), keeping the part
    that's actually likely to be mapped. Returns None if none of those
    keywords appear (nothing to simplify).
    """
    keywords = r"(JALAN|PERSIARAN|LORONG|PRESINT|TAMAN|BANDAR|LEBUH|LINGKARAN)"
    match = re.search(keywords, address, re.IGNORECASE)
    if not match:
        return None
    return address[match.start():]


def geocode_address(address, cache):
    """
    Returns (lat, lon) or None, using the cache first. On a cache miss,
    tries three queries in order, keeping the first that's actually
    usable:
      1. The full address, as free text.
      2. A simplified version with any unmapped unit/block prefix
         stripped off (see _simplify_to_street_level) — usually
         succeeds where #1 fails, since OSM rarely has unit-level data
         for Malaysian residential addresses, but the street/precinct
         name after it often IS mapped.
      3. A STRUCTURED query on just the postcode (a 5-digit number in
         the address) — the last resort. This is deliberately coarser
         (postcode-level, not street-level), so unlike #1 and #2 it is
         NOT rejected for being "too broad" — that check exists to
         catch a detailed address silently collapsing into a whole
         city/state match, not to second-guess a fallback that's
         supposed to be coarse.

    If none of the three give a usable result, this is treated as a
    genuine failure (None) rather than silently accepting a city/state-
    wide pin that would look "successful" but actually collapse many
    different addresses onto one identical point.
    """
    key = address.strip()
    if not key:
        return None
    if key in cache:
        return tuple(cache[key]) if cache[key] else None

    time.sleep(REQUEST_DELAY_SECONDS)
    result = _nominatim_query({
        "q": key,
        "format": "jsonv2",
        "countrycodes": "my",
        "limit": 1,
    })
    if result is not None and not _is_specific_enough(result):
        result = None

    if result is None:
        simplified = _simplify_to_street_level(key)
        if simplified:
            time.sleep(REQUEST_DELAY_SECONDS)
            result = _nominatim_query({
                "q": simplified,
                "format": "jsonv2",
                "countrycodes": "my",
                "limit": 1,
            })
            if result is not None and not _is_specific_enough(result):
                result = None

    if result is None:
        postcode_match = re.search(r"\b(\d{5})\b", key)
        if postcode_match:
            time.sleep(REQUEST_DELAY_SECONDS)
            # No _is_specific_enough check here on purpose — see docstring.
            result = _nominatim_query({
                "postalcode": postcode_match.group(1),
                "country": "Malaysia",
                "format": "jsonv2",
                "limit": 1,
            })

    coords = (float(result["lat"]), float(result["lon"])) if result else None
    if result:
        print(f"    [geocode] \"{key[:40]}...\" -> matched \"{result.get('display_name', '?')}\"")
    cache[key] = list(coords) if coords else None
    save_geocode_cache(cache)
    return coords


# ============================================================
# Distance + natural-proximity grouping (no fixed count/size)
# ============================================================

def haversine_km(lat1, lon1, lat2, lon2):
    """Great-circle distance between two lat/lon points, in kilometers."""
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def group_by_proximity(points, radius_km=CLUSTER_RADIUS_KM):
    """
    Union-find (disjoint set) over points: any two within radius_km of
    each other get linked into the same group, and that chains — if A
    is close to B and B is close to C, all three end up in one group
    even if A and C themselves are farther apart than radius_km. This
    is deliberately NOT k-means: there's no target number of groups or
    target group size, groups just fall out of the actual geography.

    Returns a list of group labels (small ints, not meaningfully
    ordered), one per input point.
    """
    n = len(points)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[ri] = rj

    for i in range(n):
        for j in range(i + 1, n):
            if haversine_km(points[i][0], points[i][1], points[j][0], points[j][1]) <= radius_km:
                union(i, j)

    roots = [find(i) for i in range(n)]
    # Relabel roots to small consecutive ints, ordered by group size
    # (largest first) purely so "Area 1" tends to be the biggest area
    # rather than an arbitrary one.
    from collections import Counter
    order = [root for root, _ in Counter(roots).most_common()]
    relabel = {root: idx for idx, root in enumerate(order)}
    return [relabel[r] for r in roots]


# ============================================================
# Google Maps link (plain URL — no API key needed)
# ============================================================

def gmaps_single_link(lat, lon):
    return f"https://www.google.com/maps/search/?api=1&query={lat},{lon}"


# ============================================================
# Visual map (Leaflet + OpenStreetMap tiles, no API key)
# ============================================================

def write_map_html(geocoded, path=MAP_HTML_FILE):
    """
    Writes a single self-contained HTML file with every geocoded
    address plotted as a colored, labeled marker (color = area, so
    areas are visually obvious at a glance).

    Leaflet's JS/CSS are embedded INLINE (read from leaflet.js and
    leaflet.css, which need to sit in the same folder as this script)
    rather than loaded from a CDN. A CDN-based version was tried first
    and produced a blank page — loading external scripts into a local
    file:// page is unreliable in practice (blocked by some firewalls,
    ad-blockers, or simply no internet at that exact moment), often
    with zero visible error. Bundling the library directly sidesteps
    all of that; the only thing that still needs an internet connection
    to display is the actual map tile imagery itself (there's no
    reasonable way to bundle that offline).
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    leaflet_js_path = os.path.join(script_dir, "leaflet.js")
    leaflet_css_path = os.path.join(script_dir, "leaflet.css")
    if not os.path.exists(leaflet_js_path) or not os.path.exists(leaflet_css_path):
        print(f"  !! leaflet.js / leaflet.css not found next to this script — "
              f"the map can't be built without them. Make sure both files "
              f"are in the same folder as route_planner.py.")
        return
    with open(leaflet_js_path, "r", encoding="utf-8") as f:
        leaflet_js = f.read()
    with open(leaflet_css_path, "r", encoding="utf-8") as f:
        leaflet_css = f.read()

    markers = []
    for r in geocoded:
        markers.append({
            "lat": r["lat"],
            "lon": r["lon"],
            "label": f"{r['contact']} ({r['sales_no']})",
            "address": r["address"],
            "phone": r["phone"],
            "color": AREA_COLORS[r["area"] % len(AREA_COLORS)],
            "area": r["area"] + 1,
        })

    if markers:
        center_lat = sum(m["lat"] for m in markers) / len(markers)
        center_lon = sum(m["lon"] for m in markers) / len(markers)
    else:
        center_lat, center_lon = 3.1390, 101.6869  # fallback: KL

    markers_json = json.dumps(markers, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Customer Address Map</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
{leaflet_css}
  html, body {{ margin: 0; padding: 0; height: 100%; font-family: Arial, sans-serif; }}
  #map {{ height: 100%; width: 100%; }}
  .popup-label {{ font-weight: bold; margin-bottom: 4px; }}
  .popup-area {{ color: #555; font-size: 0.85em; }}
</style>
</head>
<body>
<div id="map"></div>
<script>
{leaflet_js}
</script>
<script>
  const markers = {markers_json};
  const map = L.map('map').setView([{center_lat}, {center_lon}], 12);
  L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
    attribution: '&copy; OpenStreetMap contributors',
    maxZoom: 19,
  }}).addTo(map);

  const bounds = [];
  markers.forEach(m => {{
    const marker = L.circleMarker([m.lat, m.lon], {{
      radius: 9,
      color: '#ffffff',
      weight: 2,
      fillColor: m.color,
      fillOpacity: 0.9,
    }}).addTo(map);
    marker.bindPopup(
      '<div class="popup-label">' + m.label + '</div>' +
      '<div>' + m.address + '</div>' +
      '<div>' + m.phone + '</div>' +
      '<div class="popup-area">Area ' + m.area + '</div>'
    );
    bounds.push([m.lat, m.lon]);
  }});
  if (bounds.length > 0) {{
    map.fitBounds(bounds, {{ padding: [40, 40] }});
  }}
</script>
</body>
</html>
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


# ============================================================
# Main
# ============================================================

def find_input_file():
    if os.path.exists(OUTPUT_FILE_XLSM):
        return OUTPUT_FILE_XLSM
    if os.path.exists(OUTPUT_FILE_XLSX):
        return OUTPUT_FILE_XLSX
    return None


def run():
    target_file = find_input_file()
    if not target_file:
        print(f"Couldn't find {OUTPUT_FILE_XLSM} or {OUTPUT_FILE_XLSX} — "
              f"run main.py first to generate the export.")
        return

    is_xlsm = target_file.endswith(".xlsm")
    wb = openpyxl.load_workbook(target_file, keep_vba=is_xlsm)
    ws = wb.active  # "Cuckoo Export" sheet

    # Column A = sales_no, D = install_contact_person, E = install_mobile1,
    # F = install_address — matching main.py's DATA_COLUMNS order.
    rows = []
    for row_idx in range(2, ws.max_row + 1):
        sales_no = ws.cell(row=row_idx, column=1).value
        if not sales_no:
            continue
        rows.append({
            "row_idx": row_idx,
            "sales_no": sales_no,
            "contact": ws.cell(row=row_idx, column=4).value or "",
            "phone": ws.cell(row=row_idx, column=5).value or "",
            "address": ws.cell(row=row_idx, column=6).value or "",
        })

    if not rows:
        print("No customer rows found in the sheet — nothing to plan.")
        return

    print(f"Geocoding {len(rows)} address(es) (cached ones are instant, "
          f"new ones take ~1s each)...")
    cache = load_geocode_cache()
    geocoded, failed = [], []
    for i, r in enumerate(rows, start=1):
        coords = geocode_address(r["address"], cache)
        if coords:
            r["lat"], r["lon"] = coords
            geocoded.append(r)
        else:
            failed.append(r)
        print(f"  [{i}/{len(rows)}] {r['sales_no']}: "
              f"{'OK' if coords else 'could not place'}")

    if not geocoded:
        print("Nothing could be geocoded — check your internet connection "
              "and try again.")
        return

    points = [(r["lat"], r["lon"]) for r in geocoded]
    area_labels = group_by_proximity(points)
    for r, area in zip(geocoded, area_labels):
        r["area"] = area

    areas = {}
    for r in geocoded:
        areas.setdefault(r["area"], []).append(r)

    # --- Write Area onto the existing sheet, column K ---
    FONT = "Arial"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    ws.cell(row=1, column=11, value="Area").font = Font(
        name=FONT, size=10, bold=True, color="FFFFFF")
    ws.cell(row=1, column=11).fill = header_fill

    for r in geocoded:
        ws.cell(row=r["row_idx"], column=11,
                value=f"Area {r['area'] + 1}").font = Font(name=FONT, size=10)
    for r in failed:
        ws.cell(row=r["row_idx"], column=11, value="?").font = Font(name=FONT, size=10)

    ws.column_dimensions["K"].width = 10

    # --- Route Plan sheet ---
    if "Route Plan" in wb.sheetnames:
        del wb["Route Plan"]
    rp = wb.create_sheet("Route Plan")

    area_fill = PatternFill("solid", fgColor="1F4E78")
    subheader_fill = PatternFill("solid", fgColor="D9E1F2")
    r_idx = 1

    for area in sorted(areas.keys()):
        area_rows = areas[area]
        cell = rp.cell(row=r_idx, column=1,
                       value=f"Area {area + 1} — {len(area_rows)} address(es)")
        cell.font = Font(name=FONT, size=12, bold=True, color="FFFFFF")
        cell.fill = area_fill
        rp.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=4)
        r_idx += 1

        headers = ["Sales No", "Contact", "Phone", "Address (tap to open)"]
        for col_idx, h in enumerate(headers, start=1):
            c = rp.cell(row=r_idx, column=col_idx, value=h)
            c.font = Font(name=FONT, size=10, bold=True)
            c.fill = subheader_fill
        r_idx += 1

        for r in area_rows:
            rp.cell(row=r_idx, column=1, value=r["sales_no"]).font = Font(name=FONT, size=10)
            rp.cell(row=r_idx, column=2, value=r["contact"]).font = Font(name=FONT, size=10)
            rp.cell(row=r_idx, column=3, value=r["phone"]).font = Font(name=FONT, size=10)
            addr_cell = rp.cell(row=r_idx, column=4, value=r["address"])
            addr_cell.hyperlink = gmaps_single_link(r["lat"], r["lon"])
            addr_cell.font = Font(name=FONT, size=10, color="1155CC", underline="single")
            addr_cell.alignment = Alignment(wrap_text=True)
            r_idx += 1

        r_idx += 1  # blank row between areas

    if failed:
        cell = rp.cell(row=r_idx, column=1, value="Could not place automatically — assign manually")
        cell.font = Font(name=FONT, size=12, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="C00000")
        rp.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=4)
        r_idx += 1
        for r in failed:
            rp.cell(row=r_idx, column=1, value=r["sales_no"]).font = Font(name=FONT, size=10)
            rp.cell(row=r_idx, column=2, value=r["contact"]).font = Font(name=FONT, size=10)
            rp.cell(row=r_idx, column=3, value=r["phone"]).font = Font(name=FONT, size=10)
            addr_cell = rp.cell(row=r_idx, column=4, value=r["address"])
            addr_cell.alignment = Alignment(wrap_text=True)
            addr_cell.font = Font(name=FONT, size=10)
            r_idx += 1

    rp.column_dimensions["A"].width = 14
    rp.column_dimensions["B"].width = 24
    rp.column_dimensions["C"].width = 16
    rp.column_dimensions["D"].width = 50

    wb.save(target_file)
    write_map_html(geocoded)

    print(f"\nDone. {len(geocoded)} address(es) grouped into {len(areas)} area(s), "
          f"{len(failed)} address(es) need manual placement.")
    print(f"Written back to {target_file} — see the 'Route Plan' sheet and "
          f"the new Area column on 'Cuckoo Export'.")
    print(f"Open {MAP_HTML_FILE} in a browser to see everything on a map.")


if __name__ == "__main__":
    run()
