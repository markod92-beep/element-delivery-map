"""venue_utils.py — shared venue override loading and matching.

Consolidates the venue_overrides.json loader that was previously copy-pasted
into process_cmmc.py, process_dispatch.py, and process_venues.py.

The module-level cache is populated once per process on the first call to
load_venue_overrides() and shared across all importers in the same pipeline run.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

_VENUE_OVERRIDES: list = []
_VENUE_OVERRIDES_BY_ID: dict = {}


def load_venue_overrides(script_dir: Path | None = None) -> list:
    """Return the list of venue override entries with pre-uppercased match strings.

    Populates the module-level cache on first call; subsequent calls return the
    cached list. script_dir defaults to the directory containing this file.
    """
    global _VENUE_OVERRIDES, _VENUE_OVERRIDES_BY_ID
    if _VENUE_OVERRIDES:
        return _VENUE_OVERRIDES
    base = script_dir if script_dir is not None else Path(__file__).parent
    p = base / "venue_overrides.json"
    if not p.exists():
        return []
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"WARN: failed to load venue_overrides.json: {e}", file=sys.stderr)
        return []
    entries = raw.get("venues", [])
    # Pre-upper the match strings so we don't re-upper them per row.
    for e in entries:
        e["_names_u"] = [s.upper() for s in e.get("name_match", [])]
        e["_addrs_u"] = [s.upper() for s in e.get("address_match", [])]
    _VENUE_OVERRIDES = entries
    _VENUE_OVERRIDES_BY_ID = {e["id"]: e for e in entries if "id" in e}
    print(f"Loaded {len(entries)} venue overrides", file=sys.stderr)
    return entries


def get_venue_overrides_by_id() -> dict:
    """Return venue overrides indexed by id (for process_venues.py). Loads if not cached."""
    if not _VENUE_OVERRIDES_BY_ID:
        load_venue_overrides()
    return _VENUE_OVERRIDES_BY_ID


def match_venue_override(venue_name: str, customer_name: str, address: str) -> dict | None:
    """Match against Key_Venue_Name and Customer_Name (process_cmmc use-case).

    Matches name against venue_name OR customer_name (case-insensitive substring).
    If address_match is present on a candidate, address must also substring-match.
    Returns the matching override dict or None.
    """
    vn_u = (venue_name or "").upper()
    cn_u = (customer_name or "").upper()
    ad_u = (address or "").upper()
    for e in load_venue_overrides():
        name_hit = any(s in vn_u or s in cn_u for s in e["_names_u"])
        if not name_hit:
            continue
        if e["_addrs_u"]:
            if not any(s in ad_u for s in e["_addrs_u"]):
                continue
        return e
    return None


def match_venue_override_for_stop(cust: str, addr: str, city: str) -> dict | None:
    """Match a dispatch stop against venue overrides (process_dispatch use-case).

    Matches name against customer; address checks Delivery_Address1 + Delivery_City.
    Returns the matching override dict or None.
    """
    cu = (cust or "").upper()
    ad = ((addr or "") + " " + (city or "")).upper()
    for e in load_venue_overrides():
        name_hit = any(s in cu for s in e["_names_u"])
        if not name_hit:
            continue
        if e["_addrs_u"]:
            if not any(s in ad for s in e["_addrs_u"]):
                continue
        return e
    return None
