"""
process_cmmc.py — build delivery_data.json from the daily CMMC_UniqueEvents.xlsx export.

Pipeline:
  CMMC_UniqueEvents.xlsx  ->  process_cmmc.py  ->  delivery_data.json  ->  publish.py (git push)  ->  Cloudflare Pages

Output matches the existing shape that the decoupled HTML shell expects:
  weekKeys, weekLabels, contractWeekly, contractCityRef, contractEvents, contractInfo
  (+ a new `meta` block with company buckets and unknown-company flags)

Usage:
    python process_cmmc.py <cmmc_xlsx_path> <output_delivery_data_json> [--ref <existing_delivery_data.json>]

The --ref flag points at the previous delivery_data.json so we can reuse
contractCityRef coordinates (no geocoding required). New FSAs that don't resolve
get reported in the summary so you can add them over time.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import calendar
import os

import openpyxl

# pypostalcode bundles the full Canadian FSA centroid dataset (no network
# required). If not installed, fall back to ref-only coords.
try:
    from pypostalcode import PostalCodeDatabase
    _PC_DB = PostalCodeDatabase()
except Exception:
    _PC_DB = None

from venue_utils import (
    load_venue_overrides,
    match_venue_override,
    get_venue_overrides_by_id,
)


# ============================================================
# CONFIG — flip any of these to change behavior
# ============================================================

# Date window on Reservation_Start_Date (inclusive).
# 2025-01-01 is the floor — pre-2025 FSA capture in POR is too spotty (57-68%
# mappable for 2019-2021, stepping up to ~89% in 2025) to be useful for the
# delivery map. Rebuild the window here if address-capture quality improves
# or the use case shifts from ops to long-horizon YoY analysis.
DATE_FLOOR = dt.date(2025, 1, 1)
DATE_CEILING_MODE = "EOY"   # "EOY" = end of current calendar year; "ROLLING" = today + N months
DATE_CEILING_MONTHS_AHEAD = 12   # only used when DATE_CEILING_MODE = "ROLLING"

# Column indices in CMMC_UniqueEvents.xlsx (0-based, confirmed 2026-04-16).
COL_CONTRACT          = 0
COL_RESERVATION_START = 3
COL_PICKUP_DATE       = 4   # Date WE pick up equipment from the customer site
COL_STORE             = 5
COL_NET_REVENUE       = 10
COL_DELIVERY_FEE      = 13
COL_DELIVERY_ADDRESS  = 27
COL_DELIVERY_POSTAL   = 28
COL_DELIVERY_FSA      = 29
COL_STATUS_FLAG       = 30
COL_DOCUMENT_TYPE     = 31
COL_CUSTOMER_NAME     = 32
COL_CUSTOMER_TYPE     = 34
COL_SALESMAN          = 35
COL_PARENT_CUSTOMER   = 36
COL_KEY_VENUE         = 37
COL_DELVR             = 38
COL_PICKUP            = 39

# Category quantity columns 41-66 in CMMC_UniqueEvents.xlsx.
# Each row stores the unit quantity for its category(ies). Per-row Net_Revenue
# spans multiple categories when populated together, so revenue-by-category
# isn't directly derivable — we only aggregate quantities for the drill-down.
# `null` (col 65) is an unlabeled overflow column that we skip.
CATEGORY_COLUMNS = [
    (41, "Glassware"),
    (42, "Flatware"),
    (43, "Dinnerware"),
    (44, "Other Tableware"),
    (45, "Chairs"),
    (46, "Linen"),
    (47, "Event Equipment"),
    (48, "Tent"),
    (49, "Bar & Beverage"),
    (50, "Other"),
    (51, "Tables"),
    (52, "Buffet & Display"),
    (53, "Platters & Trays"),
    (54, "Serving"),
    (55, "Disposables"),
    (56, "Decor & Accessories"),
    (57, "Drape"),
    (58, "Services"),
    (59, "Furniture"),
    (60, "Kitchen & Food Service"),
    (61, "Lighting"),
    (62, "Discount"),
    (63, "Delivery"),
    (64, "Floors"),
    (66, "Anchoring"),
]

# ------------------------------------------------------------
# Column binding by HEADER NAME (added 2026-06-26).
#
# The fixed indices above are layout "confirmed 2026-04-16" and are kept only
# as fallback defaults. On 2026-06-26 the CMMC export inserted a column, which
# shifted every field from ~index 3 onward by +1. Because the reader used fixed
# positions, Document_Type was read from Status_Flag, NOTHING matched "Contract",
# and all 141,864 rows were dropped — silently freezing the live map for days.
#
# We now resolve columns from the export's header row at load time and FAIL LOUD
# if an expected column is missing, so a layout change can never again silently
# zero out the data.
COL_NAMES = {
    "COL_CONTRACT":          "Contract_Number",
    "COL_RESERVATION_START": "Reservation_Start_Date",
    "COL_PICKUP_DATE":       "Pickup_Date",
    "COL_STORE":             "Store",
    "COL_NET_REVENUE":       "Net_Revenue",
    "COL_DELIVERY_FEE":      "Delivery_Fee",
    "COL_DELIVERY_ADDRESS":  "Delivery_Address",
    "COL_DELIVERY_POSTAL":   "Delivery_Postal_Code",
    "COL_DELIVERY_FSA":      "Delivery_FSA",
    "COL_STATUS_FLAG":       "Status_Flag",
    "COL_DOCUMENT_TYPE":     "Document_Type",
    "COL_CUSTOMER_NAME":     "Customer_Name",
    "COL_CUSTOMER_TYPE":     "Customer_Type",
    "COL_SALESMAN":          "Salesman_Name",
    "COL_PARENT_CUSTOMER":   "Parent_Customer_Name",
    "COL_KEY_VENUE":         "Key_Venue_Name",
    "COL_DELVR":             "Delvr",
    "COL_PICKUP":            "Pickup",
}

# Category quantity columns, resolved by export header name.
# Tuple = (output_label_used_in_JSON, export_header_name).
CATEGORY_COLUMN_NAMES = [
    ("Glassware", "Glassware"),
    ("Flatware", "Flatware"),
    ("Dinnerware", "Dinnerware"),
    ("Other Tableware", "Other Tableware"),
    ("Chairs", "Chairs"),
    ("Linen", "Linen"),
    ("Event Equipment", "Event Equipment"),
    ("Tent", "Tent"),
    ("Bar & Beverage", "Bar & Beverage"),
    ("Other", "Other"),
    ("Tables", "Tables"),
    ("Buffet & Display", "Buffet & Display"),
    ("Platters & Trays", "Platters & Trays"),
    ("Serving", "Serving"),
    ("Disposables", "Disposables"),
    ("Decor & Accessories", "Decor & Accessories"),
    ("Drape", "Drape"),
    ("Services", "SERVICES"),
    ("Furniture", "Furniture"),
    ("Kitchen & Food Service", "Kitchen & Food Service Equipment"),
    ("Lighting", "Lighting"),
    ("Discount", "Discount"),
    ("Delivery", "Delivery"),
    ("Floors", "Floors"),
    ("Anchoring", "Anchoring"),
]


def resolve_columns(header_row) -> None:
    """Bind COL_* globals and CATEGORY_COLUMNS to actual indices using the
    export's header row. Fails loudly if an expected column is missing."""
    norm = {}
    for idx, name in enumerate(header_row):
        if name is None:
            continue
        norm[str(name).strip()] = idx

    g = globals()
    missing = []
    for const, hname in COL_NAMES.items():
        if hname in norm:
            g[const] = norm[hname]
        else:
            missing.append(hname)

    cats = []
    for label, hname in CATEGORY_COLUMN_NAMES:
        if hname in norm:
            cats.append((norm[hname], label))
        else:
            missing.append(hname)

    if missing:
        raise SystemExit(
            "process_cmmc: CMMC export is missing expected column(s): "
            + ", ".join(sorted(set(missing)))
            + ".\n  The export layout changed. Update COL_NAMES / "
            "CATEGORY_COLUMN_NAMES in process_cmmc.py to match the new header."
        )
    g["CATEGORY_COLUMNS"] = cats


# Document types we keep.
ALLOWED_DOC_TYPES = {"Contract"}

# Status flag normalization.
#   Closed Contract    -> "closed"
#   Open Reservation   -> "open"
# Anything else is dropped and reported.
STATUS_MAP = {
    "Closed Contract":  "closed",
    "Open Reservation": "open",
}

# Company buckets. UI filter uses these.
# Any Store code NOT listed here is flagged in output `meta.unknown_companies`
# so Mark can classify it in the next iteration.
COMPANY_BUCKETS = {
    "GTA": {"CFR", "CMM", "ERG", "HIG"},
    "GTA Tents":   {"ATR", "RT"},
    "East":        {"MFL"},
    "West":        {"A&B", "LW"},
}
KNOWN_COMPANIES = set().union(*COMPANY_BUCKETS.values())

# Internal transfer filter: customer names containing any of these tokens
# as a standalone 3-letter word are treated as Element-to-Element transfers
# and excluded from the map.
INTERNAL_TOKENS = {"ATR", "CFR", "CMM", "ERG", "HIG", "RT", "MFL", "LW", "A&B"}
INTERNAL_TOKEN_RE = re.compile(
    r"\b(" + "|".join(re.escape(t) for t in INTERNAL_TOKENS) + r")\b",
    re.IGNORECASE,
)

# Prefix-based filter for Element-branded customer names. Catches inter-warehouse
# transfers (e.g. "ELEMENT EVENT SOLUTIONS - LEASIDE") that don't carry a store
# code token in the customer name. These rows are $0 / Discount-only placeholders
# that POR's address picker occasionally mis-geocodes (e.g. 7 Westwood Ct →
# L0S/Niagara instead of M4G/Leaside), so they pollute the revenue layer.
INTERNAL_CUSTOMER_PREFIXES = (
    "ELEMENT EVENT SOLUTIONS",
    "ELEMENT RENTAL",
)

# ============================================================
# Helpers
# ============================================================


def iso_week_key(d: dt.date) -> str:
    y, w, _ = d.isocalendar()
    return f"{y}-W{w:02d}"


def week_label(d: dt.date) -> str:
    """Human-readable week label matching existing 'Dec 30 – Jan 05' style."""
    # Find Monday of the ISO week
    monday = d - dt.timedelta(days=d.isoweekday() - 1)
    sunday = monday + dt.timedelta(days=6)
    return f"{monday.strftime('%b %d')} – {sunday.strftime('%b %d')}"


def build_week_index(floor: dt.date, ceiling: dt.date) -> tuple[list[str], dict[str, str]]:
    keys: list[str] = []
    labels: dict[str, str] = {}
    # Walk week-by-week from floor
    cur = floor - dt.timedelta(days=floor.isoweekday() - 1)  # Monday of floor's week
    end = ceiling
    seen: set[str] = set()
    while cur <= end:
        k = iso_week_key(cur)
        if k not in seen:
            seen.add(k)
            keys.append(k)
            labels[k] = week_label(cur)
        cur += dt.timedelta(days=7)
    return keys, labels


def to_date(v) -> dt.date | None:
    """Parse a cell value as a date. Drops sentinel years (<2000) that POR emits
    when a user enters a time without a date — Excel stores those as 1899-12-30
    or 1900-01-01, which would poison downstream min/max aggregations."""
    d = None
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        d = v.date()
    elif isinstance(v, dt.date):
        d = v
    elif isinstance(v, str):
        try:
            d = dt.date.fromisoformat(v[:10])
        except ValueError:
            return None
    else:
        return None
    if d and d.year < 2000:
        return None
    return d


def to_float(v) -> float:
    if v is None:
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def to_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    if isinstance(v, str):
        return v.strip().lower() in {"true", "1", "yes", "y"}
    return False


def normalize_fsa(postal: str | None, fsa_col: str | None) -> str | None:
    """Return a 3-char FSA (uppercase) from FSA column, falling back to postal."""
    for candidate in (fsa_col, postal):
        if not candidate:
            continue
        s = str(candidate).strip().upper().replace(" ", "")
        if len(s) >= 3 and s[0].isalpha() and s[1].isdigit() and s[2].isalpha():
            return s[:3]
    return None


# ----------------------------------------------------------------------------
# Ambiguous-FSA → city splitting
#
# Some rural FSAs span multiple distinct cities (e.g. V0N covers the entire
# Sea-to-Sky corridor: Whistler, Pemberton, D'Arcy, Lillooet — town centres
# 50+ km apart). For these we read the city from the delivery address and
# emit a "CITY:NAME" locKey so the heatmap can show each town as its own
# bubble. The JS side has matching cityKeyToZone + zoneCoords entries.
#
# Add new entries when an Ontario / cross-province FSA gets used for >1
# meaningful destination. Single-city FSAs should stay on the FSA path —
# they don't need this split overhead.
# ----------------------------------------------------------------------------
AMBIGUOUS_FSAS: set[str] = {"V0N"}  # Sea-to-Sky: Whistler / Pemberton / Lillooet / D'Arcy

# ----------------------------------------------------------------------------
# Customer-name → city hints (last-ditch rescue)
#
# Some recurring customers strongly imply a fixed delivery city even when the
# Delivery_Address column is blank (POR data-quality issue: large institutional
# customers often have addresses entered once at the customer level and not
# repeated on each contract). When venue_override matching fails AND there's
# no address to extract a city from, fall back to this dict.
#
# Format: substring match against UPPER(customer_name). First match wins.
# Substring is intentionally permissive — "ROYAL CANADIAN MOUNTED POLICE"
# matches "ROYAL CANADIAN MOUNTED POLICE PROCUREMENT", etc.
#
# Add cautiously — this is the lowest-precision rescue lane. Only use when the
# customer-to-city link is genuinely 1:1 in practice (a customer that delivers
# to multiple cities should NOT be in here). For multi-location customers,
# encode them as venue_overrides with address_match constraints instead.
# ----------------------------------------------------------------------------
CUSTOMER_CITY_HINTS: list[tuple[str, str]] = [
    # Substring of upper(customer_name)        →  CITY tag (matches a known
    #                                            cityKeyToZone or _CITY_FALLBACK
    #                                            entry on the JS side)
    ("ROYAL CANADIAN MOUNTED POLICE",          "HALIFAX"),  # 2025 Invictus context
    ("CANADIAN ASSOCIATION OF DEFENCE",        "OTTAWA"),   # CANSEC trade show
    ("WEB SUMMIT",                             "VANCOUVER"),  # 2025 Web Summit Vancouver
    ("TENNIS CANADA",                          "TORONTO"),  # safety net — main override should catch
]


def customer_city_hint(customer_name: str | None) -> str | None:
    """Return a city tag (UPPER) implied by the customer name, or None.

    Only used as a final rescue before a row would land in unmappedWeekly.
    Matched as a case-insensitive substring; first hit wins.
    """
    if not customer_name:
        return None
    cu = str(customer_name).upper()
    for needle, city in CUSTOMER_CITY_HINTS:
        if needle in cu:
            return city
    return None


# ----------------------------------------------------------------------------
# Store-code → warehouse-region city (deepest fallback)
#
# Every CMMC row has a Store column (alpha codes: CFR / CMM / MFL / A&B etc.)
# that ties it to a specific Element warehouse. When FSA, address, AND
# customer-name all fail to yield a city, the warehouse region is the most
# defensible last guess: a Halifax-store row almost certainly delivered in
# Atlantic Canada; a BC-store row in BC; a GTA-store row in greater Toronto.
#
# This is COARSER than the customer-hint lane — multiple cities can share
# a warehouse — but it's better than dropping the row into off-map. Mark
# can override per-customer or per-venue with a more specific hint (above)
# when accuracy matters.
#
# Codes mirror COMPANY_BUCKETS. East = MFL = Halifax/Dartmouth; West = A&B/LW
# = Vancouver area; everything else = Toronto. ERG/HIG could be split out to
# Stoney Creek / Etobicoke if a future audit shows it matters.
# ----------------------------------------------------------------------------
STORE_CITY_HINTS: dict[str, str] = {
    # GTA — all roll up to Toronto for off-map rescue (warehouse-anchored,
    # too coarse to distinguish individual sub-cities without an address).
    "CFR": "TORONTO",
    "CMM": "TORONTO",
    "ERG": "TORONTO",
    "HIG": "TORONTO",
    "ATR": "TORONTO",
    "RT":  "TORONTO",
    # East
    "MFL": "HALIFAX",
    # West
    "A&B": "VANCOUVER",
    "LW":  "VANCOUVER",
}


def store_city_hint(store_code: str | None) -> str | None:
    """Return a city tag (UPPER) implied by the warehouse store code, or None.

    Final rescue lane — runs after FSA, address-city, and customer-hint have
    all failed. Coarse but defensible: warehouse region is a strict superset
    of where deliveries from that warehouse end up.
    """
    if not store_code:
        return None
    return STORE_CITY_HINTS.get(str(store_code).strip().upper())

# Token-level junk patterns mirrored from the JS _cityFromAddr (index.html).
# Keep these in sync if either side gets tightened.
_PROV_RE        = re.compile(r"^[A-Z]{2}$", re.IGNORECASE)
_FSA_TOKEN_RE   = re.compile(r"^[A-Z]\d[A-Z]$", re.IGNORECASE)
_LDU_RE         = re.compile(r"^\d[A-Z]\d$", re.IGNORECASE)
_FULL_POSTAL_RE = re.compile(r"^[A-Z]\d[A-Z]\d[A-Z]\d$", re.IGNORECASE)
_US_ZIP_RE      = re.compile(r"^\d{5}(-\d{4})?$")


def _is_addr_junk_token(tok: str) -> bool:
    """True for province codes / postal fragments that should never be
    returned as a city tag."""
    if not tok:
        return True
    if _PROV_RE.match(tok):        return True
    if _FSA_TOKEN_RE.match(tok):   return True
    if _LDU_RE.match(tok):         return True
    if _FULL_POSTAL_RE.match(tok): return True
    if _US_ZIP_RE.match(tok):      return True
    return False


_STREET_SUFFIX_TOKENS = {
    # Standard suffixes
    "ST", "STREET", "STR",
    "AVE", "AVENUE", "AV",
    "BLVD", "BOULEVARD", "BV",
    "RD", "ROAD",
    "DR", "DRIVE",
    "CT", "COURT", "CRT",
    "PL", "PLACE",
    "LN", "LANE",
    "CR", "CRES", "CRESCENT",
    "HWY", "HIGHWAY",
    "PKWY", "PARKWAY",
    "TRAIL", "TR",
    "WAY", "WY",
    "CIR", "CIRCLE",
    "SQ", "SQUARE",
    "TER", "TERRACE",
    "GATE", "GT",
    "MEWS",
    "BAY",
    # French / Quebec equivalents (POR has bilingual addresses)
    "RUE", "BOUL", "CHEMIN", "CH", "ROUTE",
}


def city_from_address(addr: str | None) -> str | None:
    """Extract a city tag (UPPER) from a POR delivery address.

    Mirrors the JS _cityFromAddr in index.html: tokenize each comma-segment
    by whitespace, drop province / FSA / postal-LDU tokens, take the first
    segment whose remainder is non-numeric. Returns None when no segment
    yields a clean city — caller falls back to the FSA locKey.

    Single-segment guard: if the address has no commas (single segment), only
    return the tokens when they look like a city name — NOT a street. We
    detect street descriptions via:
      - presence of a street-suffix token (ST, AVE, BLVD, WAY, RD, etc.)
      - presence of an "&" intersection marker
      - leading numeric token (street number)
    Without this guard, "OLYMPIC WAY" or "13601 LESLIE STREET" would each
    return themselves as "city", which would then poison _CITY_FALLBACK
    lookups on the JS side. The 2026-05-04 audit caught $634K rescued under
    "OLYMPIC WAY" — exactly this false-positive pattern.
    """
    if not addr:
        return None
    parts = [p.strip() for p in str(addr).split(",") if p and p.strip()]
    if not parts:
        return None

    # Walk right-to-left (city is typically the penultimate segment) and pick
    # the first segment that yields non-numeric tokens.
    candidate: str | None = None
    for p in reversed(parts):
        tokens = [t for t in p.split() if not _is_addr_junk_token(t)]
        if not tokens:
            continue
        if tokens[0].isdigit():
            continue
        candidate = " ".join(tokens).upper()
        break

    if not candidate:
        return None

    # Single-segment guard: an address like "OLYMPIC WAY" or "WELLINGTON & JOHN"
    # is a street description without explicit city info. Reject by checking
    # for street-suffix tokens or an intersection ampersand.
    if len(parts) == 1:
        if "&" in candidate:
            return None
        candidate_tokens = candidate.split()
        if any(t in _STREET_SUFFIX_TOKENS for t in candidate_tokens):
            return None
        # Reject single-segment addresses with parenthesized venue annotations
        # (e.g. "MARKET STREET (HALIFAX CONVENTION CENTRE)") — the venue name
        # is more useful as a venue_override match than a city tag.
        if "(" in candidate or ")" in candidate:
            return None

    return candidate


def classify_delivery(delvr: bool, pickup: bool, fee: float) -> str:
    """Match the legacy classification: Delvr or Pickup True → 'D';
    both False + fee > 0 → 'D'; else 'CP'."""
    if delvr or pickup:
        return "D"
    if fee > 0:
        return "D"
    return "CP"


def company_bucket(store: str) -> str | None:
    for bucket, members in COMPANY_BUCKETS.items():
        if store in members:
            return bucket
    return None


def is_internal_transfer(customer_name: str) -> bool:
    if not customer_name:
        return False
    up = customer_name.strip().upper()
    if up.startswith(INTERNAL_CUSTOMER_PREFIXES):
        return True
    return bool(INTERNAL_TOKEN_RE.search(customer_name))


# ============================================================
# Main extract
# ============================================================


def process(xlsx_path: Path, ref_path: Path | None) -> dict:
    # Load previous contractCityRef so we can reuse coordinates.
    # Falls back gracefully to _ref_only.json (pristine snapshot) then empty dict
    # if the primary ref is missing/corrupt. A truncated delivery_data.json used
    # to hard-fail the whole pipeline (see 2026-04-22 health log) — now it just
    # forces a re-geocode of unknown locations, pipeline continues.
    ref_city_ref: dict = {}
    ref_candidates: list[Path] = []
    if ref_path and ref_path.exists():
        ref_candidates.append(ref_path)
    # Secondary fallback: _ref_only.json colocated with the script
    ref_only = Path(__file__).parent / "_ref_only.json"
    if ref_only.exists() and ref_only not in ref_candidates:
        ref_candidates.append(ref_only)

    for candidate in ref_candidates:
        try:
            with candidate.open(encoding="utf-8") as f:
                prev = json.load(f)
            ref_city_ref = prev.get("contractCityRef", {})
            print(f"[cmmc] Loaded {len(ref_city_ref)} reference locations from {candidate.name}", file=sys.stderr)
            break
        except (json.JSONDecodeError, OSError) as e:
            print(f"[cmmc] WARN: could not load ref from {candidate.name} ({type(e).__name__}: {e}); trying next", file=sys.stderr)
            continue
    else:
        if ref_candidates:
            print("[cmmc] WARN: all ref candidates failed to parse — proceeding with empty contractCityRef (geocoding will rebuild from scratch)", file=sys.stderr)

    today = dt.date.today()
    if DATE_CEILING_MODE == "EOY":
        # End of current calendar year
        ceiling = dt.date(today.year, 12, 31)
    else:
        # End-of-month in month (today + N months)
        ceiling_year = today.year + ((today.month - 1 + DATE_CEILING_MONTHS_AHEAD) // 12)
        ceiling_month = ((today.month - 1 + DATE_CEILING_MONTHS_AHEAD) % 12) + 1
        last_day = calendar.monthrange(ceiling_year, ceiling_month)[1]
        ceiling = dt.date(ceiling_year, ceiling_month, last_day)
    print(f"[cmmc] Date window: {DATE_FLOOR}  →  {ceiling}", file=sys.stderr)

    week_keys, week_labels = build_week_index(DATE_FLOOR, ceiling)

    # Aggregates
    contract_weekly: dict = {}                     # wk -> loc -> status -> co -> agg
    contract_events: dict = defaultdict(lambda: defaultdict(list))  # loc -> venue -> [row,...]
    contract_info: dict = {}                       # contract# -> info
    contract_revenue: dict = defaultdict(float)    # contract# -> rev sum (for tt)
    contract_delivery_fee: dict = defaultdict(float)  # contract# -> Delivery_Fee sum (for per-truck rate)
    # Per-contract category roll-up for the drill-down popup. Sums each
    # CATEGORY_COLUMNS quantity across every row of the contract.
    contract_cats: dict = defaultdict(lambda: defaultdict(float))
    # Key includes company bucket so the UI can filter by East / West / GTA
    # Rentals / GTA Tents without re-aggregating from contract_events each render.
    agg_by_week_loc_status_co: dict = defaultdict(lambda: {
        "r": 0.0, "f": 0.0, "d": 0, "p": 0, "e": 0,
        "venues": defaultdict(lambda: {"r": 0.0, "f": 0.0, "d": 0}),
    })

    # Revenue that survives every filter EXCEPT FSA resolution. POR data has a
    # lot of rows with a blank delivery postal — private estates, parks, custom
    # venues, customer pickups coded without a drop-off, etc. Those can't go on
    # a map but they're real revenue, so we track them separately by (week,
    # status, company) so the UI can surface an off-map stat that ties totals
    # back to POR's raw numbers.
    unmapped_by_week_status_co: dict = defaultdict(lambda: {
        "r": 0.0, "f": 0.0, "e": 0, "c": set(),
    })
    # Per-run sample of rows that landed in the off-map bucket — captures
    # contract/customer/address/revenue so Mark can review what's still
    # unresolved AFTER the FSA → city-tag rescue ran. Capped to keep the run
    # fast and the log readable; ranked by revenue at print-time so the highest-
    # impact rows surface first. Not written to JSON — print-only diagnostic.
    unmapped_samples: list = []

    company_counts: dict = defaultdict(int)        # store -> count (kept rows)
    unknown_companies: dict = defaultdict(int)     # store -> count (kept rows)

    # Diagnostics
    n_total = 0
    n_kept = 0
    dropped = defaultdict(int)

    print(f"[cmmc] Opening {xlsx_path} (streaming)...", file=sys.stderr)
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
      ws = wb["Unique_Events"]

      # Bind columns by header name (row 1), then stream data rows (row 2+).
      row_iter = ws.iter_rows(min_row=1, values_only=True)
      header_row = next(row_iter)
      resolve_columns(header_row)

      for row in row_iter:
        n_total += 1
        if n_total % 25000 == 0:
            print(f"[cmmc]   ...scanned {n_total:,} rows (kept {n_kept:,})", file=sys.stderr)

        # Document type filter
        doc_type = row[COL_DOCUMENT_TYPE]
        if doc_type not in ALLOWED_DOC_TYPES:
            dropped["doc_type"] += 1
            continue

        # Status
        status_flag = row[COL_STATUS_FLAG]
        status = STATUS_MAP.get(status_flag)
        if status is None:
            dropped[f"status:{status_flag}"] += 1
            continue

        # Date window
        rsd = to_date(row[COL_RESERVATION_START])
        if rsd is None:
            dropped["no_date"] += 1
            continue
        if rsd < DATE_FLOOR or rsd > ceiling:
            dropped["out_of_window"] += 1
            continue

        # Customer-name-based internal transfer filter
        customer_name = (row[COL_CUSTOMER_NAME] or "").strip()
        if is_internal_transfer(customer_name):
            dropped["internal_transfer"] += 1
            continue

        # Company tag
        store = (row[COL_STORE] or "").strip().upper()
        bucket = company_bucket(store)
        is_unknown_bucket = bucket is None
        if is_unknown_bucket:
            # Keep the row in the data but tag as "Other" so it shows up when
            # no company filter is applied. Mark will classify in next pass.
            bucket = "Other"

        # Contract number (required for any further accounting)
        contract = str(row[COL_CONTRACT] or "").strip()
        if not contract:
            dropped["no_contract"] += 1
            continue

        # Financials (needed here because unmapped tracking uses them)
        rent = to_float(row[COL_NET_REVENUE])           # see CONFIG note on revenue semantics
        fee = to_float(row[COL_DELIVERY_FEE])

        # Week key (needed for both the kept aggregate and the unmapped bucket)
        wk = iso_week_key(rsd)

        # Location key — drops here are real revenue we still want to account
        # for on an "off-map" tally so totals tie back to POR.
        fsa_loc = normalize_fsa(row[COL_DELIVERY_POSTAL], row[COL_DELIVERY_FSA])

        # Venue override check. If the row's venue/customer name matches a
        # known landmark, we split it out of its dense FSA bubble into its
        # own precise-coord bubble, keyed as "V:<id>" in contract_city_ref.
        # Peek key_venue + customer_name + address here (before the normal
        # venue-name computation later in the loop).
        _ovr_key_venue = (row[COL_KEY_VENUE] or "").strip()
        _ovr_customer  = (row[COL_CUSTOMER_NAME] or "").strip()
        _ovr_address   = (row[COL_DELIVERY_ADDRESS] or "").strip()
        override = match_venue_override(_ovr_key_venue, _ovr_customer, _ovr_address)
        if override is not None:
            loc = f"V:{override['id']}"
        elif fsa_loc and fsa_loc in AMBIGUOUS_FSAS:
            # Rural FSA spanning multiple cities — split by city tag so the
            # heatmap shows separate bubbles. If city extraction fails, fall
            # back to the raw FSA so the event still renders via fsaToZone.
            _city = city_from_address(_ovr_address)
            loc = f"CITY:{_city}" if _city else fsa_loc
        elif fsa_loc:
            loc = fsa_loc
        else:
            # No FSA at all (postal column blank, malformed, or non-Canadian
            # format that normalize_fsa rejected). Run a chained rescue:
            #   1. Address city extraction — most reliable when present.
            #   2. Customer-name → city hint — for institutional customers
            #      whose delivery city is fixed even when address is blank.
            #   3. Store-code → warehouse-region city — coarsest fallback.
            #      Better than dropping into off-map; defensible because the
            #      warehouse region is a strict superset of where deliveries
            #      from that warehouse end up.
            # If all three fail, fall into unmapped/off-map — totals still
            # tie back to POR.
            _city = city_from_address(_ovr_address)
            if not _city:
                _city = customer_city_hint(_ovr_customer)
            if not _city:
                _city = store_city_hint(store)
            loc = f"CITY:{_city}" if _city else None

        if loc is None:
            dropped["no_location"] += 1
            u = unmapped_by_week_status_co[(wk, status, bucket)]
            u["r"] += rent
            u["f"] += fee
            u["e"] += 1
            u["c"].add(contract)
            # Capture diagnostic sample so Mark can review what's still
            # unresolved. Cap at 200 — ranked + trimmed before printing.
            if len(unmapped_samples) < 200:
                unmapped_samples.append({
                    "wk": wk,
                    "contract": contract,
                    "customer": (_ovr_customer or "")[:60],
                    "address": (_ovr_address or "")[:80],
                    "venue": (_ovr_key_venue or "")[:60],
                    "rent": rent,
                    "fee": fee,
                    "company": bucket,
                })
            continue

        # Only count toward kept-row store tallies once the row is fully valid.
        if is_unknown_bucket:
            unknown_companies[store] += 1
        company_counts[store] += 1

        # Delivery classification
        delvr = to_bool(row[COL_DELVR])
        pickup = to_bool(row[COL_PICKUP])
        cls = classify_delivery(delvr, pickup, fee)

        # Venue
        venue = (row[COL_KEY_VENUE] or "").strip()
        if not venue:
            venue = customer_name or "UNKNOWN"

        # --- keep this row ---
        n_kept += 1

        # Per-contract category roll-up. Quantities only — revenue can't be
        # split cleanly because a single row often spans multiple categories.
        cat_tally = contract_cats[contract]
        for col_idx, cat_name in CATEGORY_COLUMNS:
            v = row[col_idx]
            if v is None:
                continue
            try:
                fv = float(v)
            except (TypeError, ValueError):
                continue
            if fv:
                cat_tally[cat_name] += fv

        a = agg_by_week_loc_status_co[(wk, loc, status, bucket)]
        a["r"] += rent
        a["f"] += fee
        a["e"] += 1
        if cls == "D":
            a["d"] += 1
        else:
            a["p"] += 1
        v = a["venues"][venue]
        v["r"] += rent
        v["f"] += fee
        if cls == "D":
            v["d"] += 1

        # Per-venue event detail (drill-down)
        contract_events[loc][venue].append([
            rsd.isoformat(),
            contract,
            round(rent, 2),
            round(fee, 2),
            cls,
            "C" if status == "closed" else "O",
        ])

        # Contract summary (last-write-wins for identifier fields is fine —
        # all rows of a contract share these)
        if contract not in contract_info:
            info = {}
            if customer_name:
                info["cn"] = customer_name
            if venue:
                info["vn"] = venue
            sp = (row[COL_SALESMAN] or "").strip()
            if sp:
                info["sp"] = sp
            ct = (row[COL_CUSTOMER_TYPE] or "").strip()
            if ct:
                info["ct"] = ct
            pc = (row[COL_PARENT_CUSTOMER] or "").strip()
            if pc:
                info["ca"] = pc
            info["co"] = bucket   # NEW: company bucket for UI filter
            info["st"] = store    # NEW: raw store code
            # Delivery address + location key — consumed by process_venues.py
            # to group contracts into venues without re-scanning CMMC. First-seen
            # per contract; POR typically writes one delivery address per contract.
            da = (row[COL_DELIVERY_ADDRESS] or "").strip()
            if da:
                info["da"] = da
            info["lk"] = loc   # FSA (e.g. 'M5V') or 'V:<override-id>'
            contract_info[contract] = info

        # Track contract rental span across all rows of the contract:
        #   sd = earliest Reservation_Start_Date (delivery)
        #   pd = latest Pickup_Date            (our on-site pickup)
        # Useful for tent contracts that span multi-day/week installs.
        _info = contract_info[contract]
        existing_sd = _info.get("sd")
        rsd_iso = rsd.isoformat()
        if existing_sd is None or rsd_iso < existing_sd:
            _info["sd"] = rsd_iso
        pud = to_date(row[COL_PICKUP_DATE])
        if pud is not None:
            pud_iso = pud.isoformat()
            existing_pd = _info.get("pd")
            if existing_pd is None or pud_iso > existing_pd:
                _info["pd"] = pud_iso

        # Delvr / Pickup flags — aggregated across all rows of the contract
        # (any row True -> contract-level True). Consumed by annotate_unassigned.py
        # to decide whether a contract is EXPECTED to have a delivery / pickup
        # dispatch entry. Without these, we can't distinguish customer-pickup
        # contracts (dv=False, pu=False) from dispatched-delivery contracts.
        if delvr:
            _info["dv"] = True
        if pickup:
            _info["pu"] = True

        contract_revenue[contract] += rent
        contract_delivery_fee[contract] += fee

    finally:
        wb.close()

    # Fold tt (total contract revenue) into contract_info
    for c, r in contract_revenue.items():
        contract_info[c]["tt"] = int(round(r))

    # Fold df (total delivery fee) into contract_info — consumed by the
    # drill-down to compute per-truck delivery rate (df / truck_count).
    for c, f in contract_delivery_fee.items():
        if c in contract_info and f:
            contract_info[c]["df"] = int(round(f))

    # Fold category quantities into contract_info. Sorted by qty desc so the UI
    # can just render top-N without re-sorting each click. Round to ints — all
    # POR category columns are unit counts, not partial units.
    # Category names are INTERNED: there are only ~25 distinct names, but they were
    # previously written out in full for every contract (305k+ times, ~2.8 MiB of
    # repeated strings). We now emit a top-level "catNames" lookup and store the
    # index instead: [[nameIdx, qty], ...]. This cut delivery_data.json from
    # 25.05 MiB to 22.03 MiB — Cloudflare Pages hard-rejects any file over 25 MiB,
    # and the old shape had just crossed it, breaking every deploy.
    # The UI (index.html) accepts BOTH shapes, so an older payload still renders.
    cat_names: list[str] = []
    cat_index: dict[str, int] = {}

    def _cat_id(name: str) -> int:
        i = cat_index.get(name)
        if i is None:
            i = len(cat_names)
            cat_index[name] = i
            cat_names.append(name)
        return i

    for c, cats in contract_cats.items():
        if c not in contract_info:
            continue
        # List form: [[nameIdx, qty], ...] keeps JSON small vs nested dicts.
        sorted_cats = sorted(
            ([_cat_id(name), int(round(q))] for name, q in cats.items() if q >= 1),
            key=lambda x: -x[1],
        )
        if sorted_cats:
            contract_info[c]["cats"] = sorted_cats

    # Collapse into the nested dict shape the UI expects:
    #   contract_weekly[wk][loc][status][co] = {r, f, d, p, e, v}
    # The extra `co` level lets the client filter by company bucket without
    # re-aggregating from contract_events each render.
    for (wk, loc, status, co), a in agg_by_week_loc_status_co.items():
        v_rows = []
        for vname, vd in a["venues"].items():
            v_rows.append([vname, round(vd["r"], 2), round(vd["f"], 2), vd["d"]])
        entry = {
            "r": int(round(a["r"])),
            "f": int(round(a["f"])),
            "d": a["d"],
            "p": a["p"],
            "e": a["e"],
            "v": v_rows,
        }
        contract_weekly.setdefault(wk, {}).setdefault(loc, {}).setdefault(status, {})[co] = entry

    # Location coordinates — reuse prior contractCityRef where possible.
    contract_city_ref: dict = {}
    unmapped_locations = []
    all_locs = set()
    for wk in contract_weekly:
        all_locs.update(contract_weekly[wk].keys())
    all_locs.update(contract_events.keys())

    # Build a lookup for venue overrides (by "V:<id>" key) so we can resolve
    # venue location keys introduced upstream in the main loop.
    _venue_by_key = {f"V:{v['id']}": v for v in load_venue_overrides()}

    pc_resolved = 0
    venue_resolved = 0
    for loc in all_locs:
        if loc in ref_city_ref:
            contract_city_ref[loc] = ref_city_ref[loc]
            continue
        # Venue-override key (precise lat/lng for named landmarks)
        if loc in _venue_by_key:
            v = _venue_by_key[loc]
            contract_city_ref[loc] = {
                "name": v.get("label", v["id"])[:40],
                "lat": round(float(v["lat"]), 4),
                "lng": round(float(v["lng"]), 4),
            }
            venue_resolved += 1
            continue
        # Try the bundled Canadian FSA dataset
        if _PC_DB is not None:
            try:
                info = _PC_DB[loc]
                contract_city_ref[loc] = {
                    "name": (info.city or loc).split(" (")[0][:40],
                    "lat": round(info.latitude, 4),
                    "lng": round(info.longitude, 4),
                }
                pc_resolved += 1
                continue
            except KeyError:
                pass
            except Exception:
                pass
        # CITY:* locKeys are deliberately resolved via the JS-side
        # cityKeyToZone + zoneCoords lookup (see ambiguous-FSA splitting
        # above). They have no FSA-shaped key for pypostalcode to find, so
        # don't flag them as unmapped — that warning is reserved for genuine
        # FSAs we couldn't geocode and need to add coverage for.
        if loc.startswith("CITY:"):
            continue
        unmapped_locations.append(loc)
    if pc_resolved:
        print(f"[cmmc] Resolved {pc_resolved} new FSAs via pypostalcode", file=sys.stderr)
    if venue_resolved:
        print(f"[cmmc] Resolved {venue_resolved} venue-override locations", file=sys.stderr)

    # Convert nested defaultdicts to plain dicts for JSON.
    ce_out: dict = {
        loc: {v: rows for v, rows in venues.items()}
        for loc, venues in contract_events.items()
    }

    # Collapse off-map revenue into a nested dict: wk → status → co → {r, f, e, c}
    # The UI sums across filters to display an "off-map" stat next to mapped totals.
    unmapped_weekly: dict = {}
    for (wk, status, co), u in unmapped_by_week_status_co.items():
        unmapped_weekly.setdefault(wk, {}).setdefault(status, {})[co] = {
            "r": int(round(u["r"])),
            "f": int(round(u["f"])),
            "e": u["e"],
            "c": len(u["c"]),
        }

    # Meta block — useful diagnostics in the payload itself so the UI can
    # surface counts if desired.
    meta = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "date_floor": DATE_FLOOR.isoformat(),
        "date_ceiling": ceiling.isoformat(),
        "rows_scanned": n_total,
        "rows_kept": n_kept,
        "company_buckets": {k: sorted(v) for k, v in COMPANY_BUCKETS.items()},
        "company_counts": dict(sorted(company_counts.items(), key=lambda x: -x[1])),
        "unknown_companies": dict(sorted(unknown_companies.items(), key=lambda x: -x[1])),
        "unmapped_locations": sorted(unmapped_locations),
        "dropped_counts": dict(sorted(dropped.items(), key=lambda x: -x[1])),
    }

    return {
        "catNames": cat_names,   # lookup for contractInfo[*].cats [[nameIdx, qty], ...]
        "weekKeys": week_keys,
        "weekLabels": week_labels,
        "contractWeekly": contract_weekly,
        "unmappedWeekly": unmapped_weekly,
        "contractCityRef": contract_city_ref,
        "contractEvents": ce_out,
        "contractInfo": contract_info,
        "meta": meta,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", type=Path, help="Path to CMMC_UniqueEvents.xlsx")
    ap.add_argument("out", type=Path, help="Output delivery_data.json path")
    ap.add_argument("--ref", type=Path, default=None,
                    help="Previous delivery_data.json — reused for location coordinates")
    args = ap.parse_args()

    if not args.xlsx.exists():
        print(f"ERROR: {args.xlsx} not found", file=sys.stderr)
        return 1

    data = process(args.xlsx, args.ref)

    args.out.parent.mkdir(parents=True, exist_ok=True)
    # Atomic write: serialize to .tmp, fsync, rename. Prevents partial-file
    # truncation if the process is killed mid-write (Task Scheduler timeout,
    # OneDrive sync contention, machine sleep, etc.). Fixed 2026-04-20 after
    # 3 consecutive days of truncated output stranded the live site.
    tmp_path = args.out.with_suffix(args.out.suffix + ".tmp")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    with open(tmp_path, "w", encoding="utf-8") as fh:
        fh.write(payload)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp_path, args.out)

    size_mb = args.out.stat().st_size / (1024 * 1024)
    m = data["meta"]
    print()
    print(f"Wrote {args.out} ({size_mb:.2f} MB)")
    print(f"  Rows scanned: {m['rows_scanned']:,}")
    print(f"  Rows kept:    {m['rows_kept']:,}")
    print(f"  Weeks:        {len(data['weekKeys'])}  ({data['weekKeys'][0]} -> {data['weekKeys'][-1]})")
    print(f"  Locations:    {len(data['contractCityRef'])} mapped, "
          f"{len(m['unmapped_locations'])} unmapped")
    print(f"  Contracts:    {len(data['contractInfo']):,}")
    print()
    print("Company counts (kept rows):")
    for k, v in m["company_counts"].items():
        tag = "" if k in KNOWN_COMPANIES else "  <-- UNKNOWN, please classify"
        print(f"  {k:<6} {v:>7,}{tag}")
    if m["unknown_companies"]:
        print()
        print("UNKNOWN COMPANIES (classify these in COMPANY_BUCKETS):")
        for k, v in m["unknown_companies"].items():
            print(f"  {k!r}: {v:,} rows")
    if m["unmapped_locations"]:
        print()
        print(f"UNMAPPED LOCATIONS ({len(m['unmapped_locations'])}) - "
              "need coords added to contractCityRef:")
        for loc in m["unmapped_locations"][:30]:
            print(f"  {loc}")
        if len(m["unmapped_locations"]) > 30:
            print(f"  ...and {len(m['unmapped_locations']) - 30} more")
    if m["dropped_counts"]:
        print()
        print("Drop reasons:")
        for reason, cnt in m["dropped_counts"].items():
            print(f"  {reason:<30} {cnt:>8,}")

    unmapped = data.get("unmappedWeekly") or {}
    if unmapped:
        off_map_by_co: dict = defaultdict(lambda: {"r": 0, "f": 0, "e": 0})
        for wk, by_status in unmapped.items():
            for status, by_co in by_status.items():
                for co, u in by_co.items():
                    b = off_map_by_co[co]
                    b["r"] += u["r"]
                    b["f"] += u["f"]
                    b["e"] += u["e"]
        total_r = sum(b["r"] for b in off_map_by_co.values())
        total_e = sum(b["e"] for b in off_map_by_co.values())
        print()
        print(f"Off-map revenue (no FSA): ${total_r:,.0f} across {total_e:,} rows")
        for co, b in sorted(off_map_by_co.items(), key=lambda x: -x[1]["r"]):
            print(f"  {co:<12} ${b['r']:>12,.0f}  ({b['e']:,} rows)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
