"""
process_venues.py — build venues_data.json from delivery_data.json + dispatch_data.json.

Pipeline position:
  process_cmmc.py      -> delivery_data.json
  process_dispatch.py  -> dispatch_data.json
  process_venues.py    -> venues_data.json           <-- this script
  publish.py           -> copy all 3 JSONs into git repo -> Cloudflare Pages

Output feeds the "Key Venues" Leaflet layer in index.html.

Ranking: top 30 venues by trailing-24-month rental revenue (TTM).
Grouping: exact Delivery_Address string match, with a light normalization pass
  (strip unit prefixes, punctuation, collapse whitespace, uppercase) as a
  safety net so "100 Front St W" and "100 FRONT STREET WEST, TORONTO" don't
  split. Contracts that hit a venue_overrides.json entry are grouped by
  override id (they already share a location key like "V:royal-york").

Lat/lng resolution for each top-30 venue (in order):
  1. venue_overrides.json (if override key in group)
  2. median of dispatch stops lat/lng for contracts in the group
  3. FSA centroid (fallback)

Popup contracts list covers the full data window (not just TTM) so the
timeline slider in the UI can scope it. Each contract includes route_id
(date + rn) if the contract appears in dispatch_data.json.

Usage:
    python process_venues.py <delivery_data.json> <dispatch_data.json> <venues_data.json>
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import statistics
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    from pypostalcode import PostalCodeDatabase
    _PC_DB = PostalCodeDatabase()
except Exception:
    _PC_DB = None


# ============================================================
# CONFIG
# ============================================================

TOP_N_VENUES_PER_REGION   = 15   # top venues   per business-unit region
TOP_N_CATERERS_PER_REGION = 15   # top caterers per business-unit region
TTM_DAYS = 730       # trailing ~24 months for the ranking window

# Region ordering for output. Matches COMPANY_BUCKETS in process_cmmc.py.
# "Unknown" catches contracts whose `co` didn't map to a known bucket — these
# are emitted last so the UI can still render them if they somehow slip through.
REGION_ORDER = ["GTA", "GTA Tents", "East", "West", "Unknown"]

# Customer type that flags a contract as a caterer relationship.
# These contracts are excluded from the venues ranking (so caterers don't
# compete with actual venues) and drive the separate caterers ranking.
CATERER_CUSTOMER_TYPE = "Caterer & Food Service"

# Warehouse / pickup-only addresses that show up as "venues" because
# Element-operated dropoff volume pushes them up the revenue list. These are
# internal locations — not true customer venues — and should never appear on
# the Key Venues layer. Matched as a substring (case-insensitive) against the
# NORMALIZED address key, which strips unit prefixes and canonicalizes street
# suffixes (ROAD -> RD, etc.), so entries below should use the normalized form.
# If the venue list ever looks contaminated by a depot again, add it here.
WAREHOUSE_ADDRESS_BLACKLIST = (
    "210 WICKSTEED",   # Meza / Leaside dispatch yard
    "510 CONSUMERS",   # Element warehouse
    "19 RANGEMORE",    # Element warehouse
)


def _is_warehouse_address(normalized: str) -> bool:
    """True if the normalized address matches any blacklisted warehouse."""
    if not normalized:
        return False
    for needle in WAREHOUSE_ADDRESS_BLACKLIST:
        if needle in normalized:
            return True
    return False

# Rounding precision for grouping caterer contracts by delivery location.
# 4 decimals ≈ 11 m — enough to collapse near-duplicates (same building,
# slightly different POR geocodes) without merging neighbors.
CATERER_LATLNG_ROUND = 4

# 15-color categorical palette for caterer bubbles. Drawn from Tableau 10 +
# a few extras. Keeps each caterer visually distinct when their bubbles
# overlap geographically.
CATERER_PALETTE = [
    "#0B4F6C", "#D94E1F", "#1B7A3A", "#7A3E9D", "#C9A227",
    "#2E86AB", "#A23B72", "#17A398", "#F18F01", "#5A5B9F",
    "#8B5A2B", "#6B7A7E", "#C44569", "#546E7A", "#B24A00",
]


# ============================================================
# Address normalization
# ============================================================

# Unit/suite/# prefixes to strip from the head of the address so
# "Unit 5 - 100 Front St W" and "100 Front St W" collapse to the same key.
_UNIT_HEAD_RE = re.compile(
    r"^\s*("
    r"(unit|ste|suite|apt|apartment|floor|fl|bldg|building|rm|room|#)\s*[a-z0-9\-]+"
    r"|[a-z0-9\-]+\s*-"
    r")\s*[-,]?\s*",
    re.IGNORECASE,
)

# Full-word street-type canonicalization. We keep abbreviated forms so
# "100 Front St" and "100 Front Street" collapse.
_STREET_CANON = {
    r"\bSTREET\b":     "ST",
    r"\bAVENUE\b":     "AVE",
    r"\bROAD\b":       "RD",
    r"\bBOULEVARD\b":  "BLVD",
    r"\bDRIVE\b":      "DR",
    r"\bCOURT\b":      "CT",
    r"\bCRESCENT\b":   "CR",
    r"\bPLACE\b":      "PL",
    r"\bHIGHWAY\b":    "HWY",
    r"\bLANE\b":       "LN",
    r"\bTRAIL\b":      "TRL",
    r"\bPARKWAY\b":    "PKWY",
    r"\bTERRACE\b":    "TER",
    r"\bWEST\b":       "W",
    r"\bEAST\b":       "E",
    r"\bNORTH\b":      "N",
    r"\bSOUTH\b":      "S",
}
_STREET_CANON_COMPILED = [(re.compile(k), v) for k, v in _STREET_CANON.items()]

# Non-alphanumeric (minus spaces) is dropped after canonicalization.
_PUNCT_RE = re.compile(r"[^A-Z0-9 ]+")
_MULTISPACE_RE = re.compile(r"\s+")


def normalize_address(raw: str) -> str:
    """Collapse address variants to a canonical key for grouping.

    Strips unit/suite prefixes, abbreviates street types, removes punctuation,
    uppercases. Empty/None returns ''."""
    if not raw:
        return ""
    s = raw.strip()
    if not s:
        return ""
    s = _UNIT_HEAD_RE.sub("", s)
    s = s.upper()
    for rx, repl in _STREET_CANON_COMPILED:
        s = rx.sub(repl, s)
    s = _PUNCT_RE.sub(" ", s)
    s = _MULTISPACE_RE.sub(" ", s).strip()
    return s


# ============================================================
# Venue override loading — shared via venue_utils.py
# ============================================================
from venue_utils import load_venue_overrides, get_venue_overrides_by_id


# ============================================================
# Dispatch indexing — contract → (date, rn) for route link
# ============================================================

def build_contract_route_index(dispatch: dict) -> dict:
    """Return { contract#: {date, rn} } for the most recent delivery stop
    per contract. Used to render a "View route" link in venue popups."""
    idx: dict = {}
    # Walk earliest → latest so later iterations overwrite with newer routes,
    # leaving the latest-dated route per contract in `idx` at the end.
    dates = sorted((dispatch.get("routes") or {}).keys())
    for d in dates:
        for rt in dispatch["routes"][d] or []:
            rn = rt.get("rn")
            for s in rt.get("stops") or []:
                con = s.get("con")
                if not con:
                    continue
                # Only link delivery-class stops so the popup's route
                # reference matches the delivery, not the pickup.
                if s.get("cls") != "D":
                    # ...unless there's no delivery entry yet; pickup/transfer
                    # is better than nothing as a route reference.
                    if con in idx:
                        continue
                idx[str(con)] = {"date": d, "rn": rn}
    return idx


def build_dispatch_coords_by_contract(dispatch: dict) -> dict:
    """Return { contract#: [ (lat, lng), ... ] } from all dispatch stops.
    Used to derive a venue's lat/lng when no override entry exists."""
    out: dict = defaultdict(list)
    for date_routes in (dispatch.get("routes") or {}).values():
        for rt in date_routes or []:
            for s in rt.get("stops") or []:
                con = s.get("con")
                ll = s.get("ll")
                if not con or not ll:
                    continue
                if not isinstance(ll, (list, tuple)) or len(ll) != 2:
                    continue
                lat, lng = ll[0], ll[1]
                if not lat or not lng:
                    continue
                out[str(con)].append((float(lat), float(lng)))
    return out


# ============================================================
# FSA / CITY:* centroid fallback
# ============================================================

# process_cmmc.py emits locKey "CITY:<NAME>" for rows that resolve to a city
# tag but no FSA (ambiguous-FSA splits like Sea-to-Sky, plus the address /
# customer-hint / store-hint rescue chain). The map UI resolves those via its
# own cityKeyToZone / zoneCoords / CITY_FALLBACK_COORDS tables, but this
# script had no equivalent — venue groups whose locKeys were all CITY:* fell
# out of the Key Venues layer as "unresolvable coords" (audit 2026-07-07
# §5.2: 10 top GTA-Tents venues dropped, largest ≈$634K TTM). This table
# mirrors the JS-side coordinates (index.html CITY_FALLBACK_COORDS + the
# zoneCoords CITY:* / Sea-to-Sky entries). Keys are the locKey suffix after
# "CITY:", uppercase. Data-entry placeholders (TBD, ADDRESS TBD, VARIOUS
# DROPS SEE LIST, ...) are deliberately absent so they keep getting dropped
# rather than rendering at a fake coordinate.
CITY_CENTROIDS: dict[str, tuple[float, float]] = {
    # --- Ontario ---
    "TORONTO":                        (43.6532, -79.3832),
    "OTTAWA":                         (45.4215, -75.6972),
    "MCMASTER UNIVERSITY":            (43.2609, -79.9192),
    "HAMILTON":                       (43.2557, -79.8711),
    "KINGSTON":                       (44.2312, -76.4860),
    "LONDON":                         (42.9849, -81.2453),
    "WINDSOR":                        (42.3149, -83.0364),
    "BARRIE":                         (44.3894, -79.6903),
    "GUELPH":                         (43.5448, -80.2482),
    "KITCHENER":                      (43.4516, -80.4925),
    "WATERLOO":                       (43.4643, -80.5204),
    "BLUE MOUNTAINS":                 (44.5004, -80.3920),
    "MUSKOKA":                        (44.9437, -79.3081),
    "JERSEYVILLE":                    (43.2000, -80.1170),
    "ORTON":                          (43.7390, -80.2060),
    "THORNBURY":                      (44.5610, -80.4520),
    "COLDWATER":                      (44.7000, -79.6170),
    "VINELAND":                       (43.1700, -79.3970),
    "PORT CARLING":                   (45.1060, -79.5850),
    "GRIMSBY":                        (43.1930, -79.5630),
    "MINETT":                         (45.0690, -79.5590),
    "ROSSEAU":                        (45.1870, -79.6360),
    # GTA sports/community venues seen in 2026 data
    "SRC JCC":                        (43.7787, -79.4385),
    "SRC JCC ARENA":                  (43.7787, -79.4385),
    "VAUGHAN GROVE SPORTS PARK":      (43.7790, -79.5430),
    "RICHMOND GREEN SPORTS CENTRE":   (43.9036, -79.4297),
    "LEBOVIC GOLF COURSE":            (43.9580, -79.4400),
    "CITY PLAYHOUSE + WESTMOUNT C.":  (43.8000, -79.5500),
    "THE SPORTS VILLAGE":             (43.7100, -79.6300),
    "VAUGHAN SPORTSPLEX":             (43.7800, -79.5800),
    "HOOPDOME":                       (43.7170, -79.4400),
    "TORONTO CITY HALL GROUNDS":      (43.6534, -79.3839),
    "BRONTE HERITAGE PARK":           (43.3950, -79.7100),
    "BRONTE HERITAGE WATERFRONT PAR": (43.3950, -79.7100),
    "WATERFRONT PARK BRONTE":         (43.3950, -79.7100),
    "CLARKSON / TRUSCOTT":            (43.5100, -79.6300),
    "CENTERVILLE CREEK":              (43.6532, -79.3832),
    # --- East (Halifax/Dartmouth NS, PEI) ---
    "HALIFAX":                        (44.6488, -63.5752),
    "EAST POINT":                     (46.4500, -61.9000),
    # --- West (Vancouver / Sea-to-Sky, BC) ---
    "VANCOUVER":                      (49.2827, -123.1207),
    "SQUAMISH":                       (49.7016, -123.1558),
    "SQUAMISH VALLEY":                (49.7016, -123.1558),
    "WHISTLER":                       (50.1163, -122.9574),
    "PEMBERTON":                      (50.3170, -122.8033),
    "LILLOOET":                       (50.6864, -121.9354),
    "DARCY":                          (50.3170, -122.8033),  # tiny community, roll up to Pemberton
    "D'ARCY":                         (50.3170, -122.8033),
    "MORRIS VALLEY":                  (49.2700, -121.9000),
    # --- Quebec ---
    "MONTREAL":                       (45.5019, -73.5674),
}


def city_centroid(loc_key: str) -> tuple[float, float] | None:
    """Return (lat, lng) for a "CITY:<NAME>" locKey, or None if unknown."""
    if not loc_key or not loc_key.startswith("CITY:"):
        return None
    return CITY_CENTROIDS.get(loc_key[5:].strip().upper())


def fsa_centroid(loc_key: str, delivery_city_ref: dict) -> tuple[float, float] | None:
    """Return (lat, lng) for an FSA from contractCityRef (preferred) or the
    bundled pypostalcode DB (fallback). CITY:* locKeys resolve via the static
    CITY_CENTROIDS table. Returns None if loc_key can't be resolved."""
    if not loc_key or loc_key.startswith("V:"):
        return None
    if loc_key.startswith("CITY:"):
        return city_centroid(loc_key)
    # First choice: contractCityRef already built by process_cmmc.py
    ref = delivery_city_ref.get(loc_key)
    if ref and ref.get("lat") and ref.get("lng"):
        return (float(ref["lat"]), float(ref["lng"]))
    # Fallback: bundled FSA DB
    if _PC_DB is not None:
        try:
            info = _PC_DB[loc_key]
            return (float(info.latitude), float(info.longitude))
        except Exception:
            pass
    return None


# ============================================================
# Core grouping + ranking
# ============================================================

def group_contracts(contract_info: dict, exclude_caterers: bool = True) -> dict:
    """Group contracts into venues by normalized delivery address (or
    override id for override-matched contracts).

    Caterer-customer contracts (customer_type == CATERER_CUSTOMER_TYPE) are
    excluded by default so they don't dominate the venues ranking — they
    feed the separate caterers view instead.

    Returns { venue_key: { "contracts": [con, ...], "address_raw": Counter,
              "name": Counter, "lk": Counter, "override_id": str|None } }
    """
    groups: dict = defaultdict(lambda: {
        "contracts": [],
        "address_raw": Counter(),   # raw address strings seen in this group
        "name": Counter(),          # key_venue/customer strings for display
        "lk": Counter(),            # FSA (or V:<id>) seen in this group
        "co": Counter(),            # business-unit bucket (GTA / East / ...)
        "override_id": None,
    })
    for con, info in contract_info.items():
        if exclude_caterers and info.get("ct") == CATERER_CUSTOMER_TYPE:
            continue
        lk = info.get("lk") or ""
        da = (info.get("da") or "").strip()

        if lk.startswith("V:"):
            venue_key = lk  # use override id as group key
            override_id = lk[2:]
        else:
            norm = normalize_address(da)
            if not norm:
                continue  # no address → can't group into a venue
            # Drop Element warehouses / pickup yards — they're not venues even
            # when revenue tied to them is large (internal transfer volume).
            if _is_warehouse_address(norm):
                continue
            venue_key = norm
            override_id = None

        g = groups[venue_key]
        g["contracts"].append(con)
        if da:
            g["address_raw"][da] += 1
        vn = info.get("vn") or info.get("cn")
        if vn:
            g["name"][vn] += 1
        if lk:
            g["lk"][lk] += 1
        co = info.get("co")
        if co:
            g["co"][co] += 1
        if override_id and g["override_id"] is None:
            g["override_id"] = override_id
    return groups


def group_caterers(contract_info: dict) -> dict:
    """Group caterer contracts by parent-customer name (ca), falling back
    to customer_name (cn). Only customer_type == CATERER_CUSTOMER_TYPE
    contracts are included.

    Returns { caterer_key: { "display_name": str, "contracts": [con, ...] } }
    """
    groups: dict = defaultdict(lambda: {
        "display_name": "",
        "contracts": [],
        "co": Counter(),  # business-unit bucket (for per-region ranking)
    })
    for con, info in contract_info.items():
        if info.get("ct") != CATERER_CUSTOMER_TYPE:
            continue
        key = (info.get("ca") or info.get("cn") or "").strip().upper()
        if not key:
            continue
        g = groups[key]
        if not g["display_name"]:
            g["display_name"] = info.get("ca") or info.get("cn") or key
        g["contracts"].append(con)
        co = info.get("co")
        if co:
            g["co"][co] += 1
    return groups


def ttm_revenue_for_contracts(contracts: list, contract_info: dict,
                              ttm_floor: dt.date, ttm_ceiling: dt.date) -> int:
    """Sum `tt` (rental) across contracts whose earliest date (sd) falls
    within the TTM window. Returns int dollars."""
    total = 0
    for c in contracts:
        info = contract_info.get(c) or {}
        sd = info.get("sd")
        if not sd:
            continue
        try:
            d = dt.date.fromisoformat(sd[:10])
        except ValueError:
            continue
        if ttm_floor <= d <= ttm_ceiling:
            total += int(info.get("tt") or 0)
    return total


# ============================================================
# Main
# ============================================================

def _dominant_region(co_counter: Counter) -> str:
    """Pick the region bucket with the most contracts in this group.
    Falls back to "Unknown" if the group has no region signal."""
    if not co_counter:
        return "Unknown"
    return co_counter.most_common(1)[0][0]


def _build_venues_section(contract_info: dict, city_ref: dict,
                          overrides: dict, coords_by_con: dict,
                          route_idx: dict,
                          ttm_floor: dt.date, ttm_ceiling: dt.date) -> tuple[list, dict]:
    """Top-N venues PER REGION, caterer-customer contracts excluded.

    Ranking is now regional: each business-unit bucket (GTA / GTA Tents /
    East / West) gets its own top-N list instead of one global list. This
    stops high-revenue urban venues from starving smaller regions off the
    map — West / East deserve to surface their own key venues.
    """
    groups = group_contracts(contract_info, exclude_caterers=True)
    print(f"[venues] Venue groups (caterers excluded): {len(groups):,}", file=sys.stderr)

    # Bucket groups by dominant region, computing TTM rev once per group.
    by_region: dict = defaultdict(list)
    for key, g in groups.items():
        ttm = ttm_revenue_for_contracts(g["contracts"], contract_info,
                                        ttm_floor, ttm_ceiling)
        if ttm <= 0:
            continue
        region = _dominant_region(g["co"])
        by_region[region].append((ttm, key, g))

    # Sort inside each region, take top-N-per-region, then flatten back out.
    # Global rank is assigned by descending TTM rev across ALL surviving
    # venues so the popup header still reads "#1 ..." meaningfully.
    top_per_region: list = []
    for region in REGION_ORDER + [r for r in by_region if r not in REGION_ORDER]:
        bucket = by_region.get(region) or []
        bucket.sort(key=lambda x: -x[0])
        top_per_region.extend((region, *tpl) for tpl in bucket[:TOP_N_VENUES_PER_REGION])
    # Global TTM sort for rank numbering
    top_per_region.sort(key=lambda x: -x[1])

    venues_out = []
    unresolved = []
    for rank, (region, ttm_rev, venue_key, g) in enumerate(top_per_region, start=1):
        # Display name: prefer override label, then key_venue name, then
        # most common raw address, then the group key.
        display_name = ""
        if overrides and g["override_id"]:
            ov = overrides.get(g["override_id"])
            if ov and ov.get("label"):
                display_name = ov["label"]
        if not display_name and g["name"]:
            display_name = g["name"].most_common(1)[0][0]
        if not display_name and g["address_raw"]:
            display_name = g["address_raw"].most_common(1)[0][0]
        if not display_name:
            display_name = venue_key

        address_raw = g["address_raw"].most_common(1)[0][0] if g["address_raw"] else ""

        lat, lng, coord_source = None, None, None
        if g["override_id"] and overrides.get(g["override_id"]):
            ov = overrides[g["override_id"]]
            lat, lng = float(ov["lat"]), float(ov["lng"])
            coord_source = "override"
        if lat is None:
            all_coords = []
            for c in g["contracts"]:
                all_coords.extend(coords_by_con.get(c, []))
            if all_coords:
                lat = statistics.median(c[0] for c in all_coords)
                lng = statistics.median(c[1] for c in all_coords)
                coord_source = "dispatch"
        if lat is None:
            # Walk the group's locKeys most-common-first so a group holding a
            # mix of resolvable and unresolvable keys still gets a coordinate.
            for lk_try, _n in g["lk"].most_common():
                centroid = fsa_centroid(lk_try, city_ref)
                if centroid:
                    lat, lng = centroid
                    coord_source = "city" if lk_try.startswith("CITY:") else "fsa"
                    break

        if lat is None:
            unresolved.append((rank, display_name, ttm_rev,
                               [k for k, _ in g["lk"].most_common(3)]))
            continue

        contracts_list = []
        for c in g["contracts"]:
            info = contract_info.get(c) or {}
            sd = info.get("sd")
            if not sd:
                continue
            row = {
                "con": c,
                "cust": info.get("cn") or "",
                "date": sd[:10],
                "tt": int(info.get("tt") or 0),
                "df": int(info.get("df") or 0),
            }
            rt = route_idx.get(c)
            if rt:
                row["route"] = rt
            contracts_list.append(row)
        contracts_list.sort(key=lambda x: x["date"], reverse=True)

        ttm_ct = sum(
            1 for c in g["contracts"]
            if (info := contract_info.get(c) or {}).get("sd")
            and ttm_floor <= dt.date.fromisoformat(info["sd"][:10]) <= ttm_ceiling
        )

        venues_out.append({
            "id": f"v{rank:03d}",
            "rank": rank,
            "region": region,
            "display_name": display_name[:80],
            "address": address_raw[:120],
            "lat": round(lat, 5),
            "lng": round(lng, 5),
            "coord_source": coord_source,
            "ttm_revenue": ttm_rev,
            "ttm_contract_count": ttm_ct,
            "total_contract_count": len(contracts_list),
            "override_id": g["override_id"],
            "contracts": contracts_list,
        })

    # Assign a per-region rank (1..N) so the UI can label regional position
    # without having to re-rank client-side.
    region_counters: dict = defaultdict(int)
    for v in sorted(venues_out, key=lambda x: -x["ttm_revenue"]):
        region_counters[v["region"]] += 1
        v["region_rank"] = region_counters[v["region"]]

    meta = {
        "groups_total": len(groups),
        "groups_with_ttm_revenue": sum(len(v) for v in by_region.values()),
        "venues_emitted": len(venues_out),
        "unresolved_coords": len(unresolved),
        "top_n_per_region": TOP_N_VENUES_PER_REGION,
        "regions": {r: len(v) for r, v in by_region.items()},
    }
    if unresolved:
        print(f"[venues] WARN: {len(unresolved)} top venues had no resolvable "
              f"lat/lng (dropped):", file=sys.stderr)
        for rank, name, rev, lks in unresolved:
            print(f"[venues]   #{rank:>2}  ${rev:>9,}  {name}  lk={lks}", file=sys.stderr)
    return venues_out, meta


def _build_caterers_section(contract_info: dict, coords_by_con: dict,
                            route_idx: dict,
                            ttm_floor: dt.date, ttm_ceiling: dt.date) -> tuple[list, dict]:
    """Top-N caterers, contracts grouped by precise dispatch lat/lng.

    Precise-only: contracts without dispatch coords are dropped from the
    bubble rendering but still counted for ranking. This matches Mark's
    "pin accuracy over coverage" decision — we only plot bubbles where we
    know exactly where the delivery happened.
    """
    caterer_groups = group_caterers(contract_info)
    print(f"[venues] Caterer groups: {len(caterer_groups):,}", file=sys.stderr)

    # Rank by TTM rental, bucketed by region (same approach as venues).
    by_region: dict = defaultdict(list)
    for key, g in caterer_groups.items():
        ttm = ttm_revenue_for_contracts(g["contracts"], contract_info,
                                        ttm_floor, ttm_ceiling)
        if ttm <= 0:
            continue
        region = _dominant_region(g["co"])
        by_region[region].append((ttm, key, g))

    top_per_region: list = []
    for region in REGION_ORDER + [r for r in by_region if r not in REGION_ORDER]:
        bucket = by_region.get(region) or []
        bucket.sort(key=lambda x: -x[0])
        top_per_region.extend((region, *tpl) for tpl in bucket[:TOP_N_CATERERS_PER_REGION])
    top_per_region.sort(key=lambda x: -x[1])

    caterers_out = []
    total_contracts_with_coords = 0
    total_contracts = 0

    for rank, (region, ttm_rev, key, g) in enumerate(top_per_region, start=1):
        # Group this caterer's contracts by rounded dispatch lat/lng.
        # Contracts without dispatch coords are tracked separately for the
        # "coverage" summary but not plotted.
        locations: dict = defaultdict(lambda: {"contracts": [], "lat": None, "lng": None})
        no_coord = []

        for c in g["contracts"]:
            info = contract_info.get(c) or {}
            sd = info.get("sd")
            if not sd:
                continue
            coords_list = coords_by_con.get(c) or []
            if not coords_list:
                no_coord.append(c)
                continue
            # Use the first dispatch stop coord (consistent across stops
            # for the same contract in the normal case).
            lat, lng = coords_list[0]
            bucket_key = (round(lat, CATERER_LATLNG_ROUND),
                          round(lng, CATERER_LATLNG_ROUND))
            row = {
                "con": c,
                "cust": info.get("cn") or "",
                "date": sd[:10],
                "tt": int(info.get("tt") or 0),
                "df": int(info.get("df") or 0),
            }
            da = (info.get("da") or "").strip()
            if da:
                row["addr"] = da[:100]
            rt = route_idx.get(c)
            if rt:
                row["route"] = rt
            bucket = locations[bucket_key]
            bucket["contracts"].append(row)
            # Keep the precise (un-rounded) lat/lng from the first contract
            # for rendering — bucket_key is only for grouping.
            if bucket["lat"] is None:
                bucket["lat"] = lat
                bucket["lng"] = lng

        # Emit one location entry per bucket, sorted by contract count desc.
        loc_list = []
        for (rlat, rlng), bucket in locations.items():
            loc_list.append({
                "lat": round(bucket["lat"], 5),
                "lng": round(bucket["lng"], 5),
                "contract_count": len(bucket["contracts"]),
                "contracts": sorted(bucket["contracts"],
                                    key=lambda x: x["date"], reverse=True),
            })
        loc_list.sort(key=lambda x: -x["contract_count"])

        total_contracts += len(g["contracts"])
        total_contracts_with_coords += sum(l["contract_count"] for l in loc_list)

        color = CATERER_PALETTE[(rank - 1) % len(CATERER_PALETTE)]

        ttm_ct = sum(
            1 for c in g["contracts"]
            if (info := contract_info.get(c) or {}).get("sd")
            and ttm_floor <= dt.date.fromisoformat(info["sd"][:10]) <= ttm_ceiling
        )

        caterers_out.append({
            "id": f"c{rank:03d}",
            "rank": rank,
            "region": region,
            "display_name": g["display_name"][:80],
            "color": color,
            "ttm_revenue": ttm_rev,
            "ttm_contract_count": ttm_ct,
            "total_contract_count": len(g["contracts"]),
            "plotted_contract_count": sum(l["contract_count"] for l in loc_list),
            "unique_locations": len(loc_list),
            "locations": loc_list,
        })

    # Per-region rank (1..N within each region) so UI can render "GTA #3".
    region_counters: dict = defaultdict(int)
    for c in sorted(caterers_out, key=lambda x: -x["ttm_revenue"]):
        region_counters[c["region"]] += 1
        c["region_rank"] = region_counters[c["region"]]

    meta = {
        "caterers_total": len(caterer_groups),
        "caterers_with_ttm_revenue": sum(len(v) for v in by_region.values()),
        "caterers_emitted": len(caterers_out),
        "top_n_per_region": TOP_N_CATERERS_PER_REGION,
        "regions": {r: len(v) for r, v in by_region.items()},
        "total_contracts": total_contracts,
        "total_contracts_with_dispatch_coords": total_contracts_with_coords,
        "coverage_pct": round(100 * total_contracts_with_coords / max(1, total_contracts), 1),
    }
    print(f"[venues] Caterer plot coverage: {total_contracts_with_coords:,}/{total_contracts:,} "
          f"contracts ({meta['coverage_pct']}%) have dispatch coords", file=sys.stderr)
    return caterers_out, meta


def build_venues(delivery: dict, dispatch: dict, overrides: dict) -> dict:
    contract_info = delivery.get("contractInfo") or {}
    city_ref = delivery.get("contractCityRef") or {}

    today = dt.date.today()
    ttm_floor = today - dt.timedelta(days=TTM_DAYS)
    ttm_ceiling = today

    print(f"[venues] TTM window: {ttm_floor}  →  {ttm_ceiling}", file=sys.stderr)
    print(f"[venues] Contracts in delivery_data.json: {len(contract_info):,}", file=sys.stderr)

    # Dispatch indexes (shared between venues + caterers sections)
    route_idx = build_contract_route_index(dispatch)
    coords_by_con = build_dispatch_coords_by_contract(dispatch)
    print(f"[venues] Dispatch route index: {len(route_idx):,} contracts", file=sys.stderr)
    print(f"[venues] Dispatch coord index: {len(coords_by_con):,} contracts", file=sys.stderr)

    venues_out, venues_meta = _build_venues_section(
        contract_info, city_ref, overrides, coords_by_con, route_idx,
        ttm_floor, ttm_ceiling)
    caterers_out, caterers_meta = _build_caterers_section(
        contract_info, coords_by_con, route_idx, ttm_floor, ttm_ceiling)

    return {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "ttm_window": {"start": ttm_floor.isoformat(), "end": ttm_ceiling.isoformat()},
        "top_n_venues_per_region":   TOP_N_VENUES_PER_REGION,
        "top_n_caterers_per_region": TOP_N_CATERERS_PER_REGION,
        "venues": venues_out,
        "caterers": caterers_out,
        "meta": {
            "venues": venues_meta,
            "caterers": caterers_meta,
        },
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("delivery_json", type=Path, help="Path to delivery_data.json")
    ap.add_argument("dispatch_json", type=Path, help="Path to dispatch_data.json")
    ap.add_argument("out", type=Path, help="Output venues_data.json path")
    args = ap.parse_args()

    if not args.delivery_json.exists():
        print(f"ERROR: {args.delivery_json} not found", file=sys.stderr)
        return 1
    if not args.dispatch_json.exists():
        print(f"ERROR: {args.dispatch_json} not found", file=sys.stderr)
        return 1

    delivery = json.loads(args.delivery_json.read_text(encoding="utf-8"))
    dispatch = json.loads(args.dispatch_json.read_text(encoding="utf-8"))
    load_venue_overrides()   # populate shared cache
    overrides = get_venue_overrides_by_id()
    print(f"[venues] Loaded {len(overrides)} venue overrides", file=sys.stderr)

    data = build_venues(delivery, dispatch, overrides)

    # Atomic write (same pattern as process_cmmc.py) — prevents partial-file
    # truncation if the process is killed mid-write.
    args.out.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = args.out.with_suffix(args.out.suffix + ".tmp")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    with open(tmp_path, "w", encoding="utf-8") as fh:
        fh.write(payload)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp_path, args.out)

    size_kb = args.out.stat().st_size / 1024
    print()
    print(f"Wrote {args.out} ({size_kb:.1f} KB)")
    print(f"  Venues emitted:   {len(data['venues'])}")
    print(f"  Caterers emitted: {len(data['caterers'])}")

    print()
    print("=" * 105)
    print(f"TOP {TOP_N_VENUES_PER_REGION} KEY VENUES PER REGION (caterer-customer contracts excluded)")
    print("=" * 105)
    print(f"{'Region':<10} {'R#':>3} {'Rank':<5} {'TTM Rev':>10} {'TTM Ct':>7} {'All Ct':>7}  Source   Venue")
    print("-" * 105)
    # Print grouped by region, per-region rank order
    for region in REGION_ORDER + sorted({v["region"] for v in data["venues"]} - set(REGION_ORDER)):
        region_venues = sorted([v for v in data["venues"] if v["region"] == region],
                               key=lambda x: x["region_rank"])
        for v in region_venues:
            print(f"{region:<10} {v['region_rank']:>3} #{v['rank']:<4} "
                  f"${v['ttm_revenue']:>9,} {v['ttm_contract_count']:>7} "
                  f"{v['total_contract_count']:>7}  "
                  f"{(v.get('coord_source') or '-'):<8} {v['display_name']}")

    print()
    print("=" * 105)
    print(f"TOP {TOP_N_CATERERS_PER_REGION} KEY CATERERS PER REGION (grouped by delivery location, precise coords only)")
    print("=" * 105)
    print(f"{'Region':<10} {'R#':>3} {'Rank':<5} {'TTM Rev':>10} {'TTM Ct':>7} {'All Ct':>7} {'Plot Ct':>7} {'Locs':>5}  Caterer")
    print("-" * 105)
    for region in REGION_ORDER + sorted({c["region"] for c in data["caterers"]} - set(REGION_ORDER)):
        region_caterers = sorted([c for c in data["caterers"] if c["region"] == region],
                                 key=lambda x: x["region_rank"])
        for c in region_caterers:
            print(f"{region:<10} {c['region_rank']:>3} #{c['rank']:<4} "
                  f"${c['ttm_revenue']:>9,} {c['ttm_contract_count']:>7} "
                  f"{c['total_contract_count']:>7} "
                  f"{c['plotted_contract_count']:>7} {c['unique_locations']:>5}  "
                  f"{c['display_name']}")

    cov = data["meta"]["caterers"]
    print()
    print(f"Caterer plot coverage: {cov['total_contracts_with_dispatch_coords']:,}/"
          f"{cov['total_contracts']:,} contracts ({cov['coverage_pct']}%)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
