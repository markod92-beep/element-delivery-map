#!/usr/bin/env python3
"""annotate_unassigned.py — flag scheduled deliveries/pickups that weren't
dispatched to a truck and render them on the map under a synthetic
"No truck assigned" truck.

Runs AFTER process_dispatch.py (and after process_venues.py, so venue popups
only link real routes). Reads delivery_data.json + dispatch_data.json, cross-
checks scheduled D/P stops per contract against dispatch `contract_trucks`,
and injects a `truck: "No truck assigned"` synthetic route per date with the
missing stops. The synthetic route is marked `virtual: 1` so index.html skips
the polyline + depot markers — the pins render in isolation.

Missing criteria (per contract):
  - DELIVERY missing:  info.dv == True  AND  info.sd in [dispatch.window]
                       AND no 'D'-class entry in contract_trucks[contract]
  - PICKUP missing:    info.pu == True  AND  info.pd in [dispatch.window]
                       AND no 'P'-class entry in contract_trucks[contract]

Contracts flagged as customer-pickup-only in POR (dv=False, pu=False) are
excluded — they never had a dispatch obligation. Same-day contracts where
delivery and pickup both happen on the same day are naturally covered: if
either leg is missing, it shows up.

Geocoding for synthetic stops:
  1. contractCityRef[info.lk] — already resolved by process_cmmc.py for every
     contract's location key ("M4V", "V:ago", etc.)
  2. pypostalcode FSA centroid fallback (in case a new FSA slipped in)
  3. Skip the stop if neither resolves (logged; happens for <1% of contracts)

Usage:
    python annotate_unassigned.py <delivery_json> <dispatch_json>
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path


# Synthetic truck identity — index.html keys visual behavior (skip polyline,
# skip depot markers) off route.virtual, so the truck label is free to vary.
# We split synth routes per region so the BU filter in the UI excludes
# unassigned stops from other regions cleanly (a single all-region synth
# route would match every BU and leak cross-region pins).
SYNTH_TRUCK_PREFIX = "No truck assigned"
SYNTH_COLOR_BY_BUCKET = {
    "GTA":       "#9CA3AF",   # slate-gray
    "GTA Tents": "#A78BFA",   # violet
    "East":      "#60A5FA",   # blue
    "West":      "#34D399",   # emerald
    "Other":     "#D1D5DB",   # light gray (catch-all)
}
SYNTH_DEFAULT_COLOR = "#9CA3AF"
# Stable rn offset per bucket so click-drilldown indexes don't collide. Each
# (date, bucket) pair produces a unique rn = SYNTH_RN_BASE + offset + YYYYMMDD.
SYNTH_RN_BASE = 900_000
SYNTH_RN_OFFSET_BY_BUCKET = {
    "GTA":       0,
    "GTA Tents": 100_000_000,
    "East":      200_000_000,
    "West":      300_000_000,
    "Other":     400_000_000,
}

# Store-code -> region bucket. Mirrors COMPANY_BUCKETS in process_cmmc.py.
# Synth stops carry the contract's raw store code in info["st"] AND the
# pre-computed bucket in info["co"] — we prefer co (already classified, handles
# edge cases) and fall back to st-mapping only if co is missing.
STORE_TO_BUCKET = {
    "CFR": "GTA", "CMM": "GTA", "ERG": "GTA", "HIG": "GTA",
    "ATR": "GTA Tents", "RT": "GTA Tents",
    "MFL": "East",
    "A&B": "West", "LW": "West",
}


def _bucket_for(info: dict) -> str:
    """Resolve an unassigned contract's region bucket. Prefer pre-classified
    info["co"] (set by process_cmmc.py); fall back to mapping the raw store
    code. Returns "Other" if neither is recognized — keeps the stop on the
    map but in a clearly-bucketed catch-all so BU filtering still works."""
    co = (info.get("co") or "").strip()
    if co in SYNTH_COLOR_BY_BUCKET:
        return co
    st = (info.get("st") or "").strip()
    bucket = STORE_TO_BUCKET.get(st)
    return bucket or "Other"


try:
    from pypostalcode import PostalCodeDatabase
    _PC_DB = PostalCodeDatabase()
except Exception:
    _PC_DB = None


def _resolve_coord(info: dict, city_ref: dict) -> tuple[float, float, str] | None:
    """Return (lat, lng, source) for a contract's delivery/pickup location,
    or None if no usable coord can be resolved.

    Source tag:
      'city_ref'    -- coord came from contractCityRef (preferred — already
                       resolved by process_cmmc.py against venue overrides
                       and pypostalcode)
      'fsa_fallback' -- coord came from pypostalcode on-the-fly (rare: new FSA)
    """
    lk = info.get("lk")
    if lk:
        ref = city_ref.get(lk)
        if ref and ref.get("lat") and ref.get("lng"):
            return float(ref["lat"]), float(ref["lng"]), "city_ref"

    # Fallback: resolve FSA directly via pypostalcode. Only fires if city_ref
    # missed (shouldn't happen in steady state, but protects against schema
    # drift).
    if lk and not lk.startswith("V:") and _PC_DB is not None:
        try:
            pc = _PC_DB[lk]
            return float(pc.latitude), float(pc.longitude), "fsa_fallback"
        except Exception:
            pass
    return None


def _build_synth_stop(contract: str, info: dict, cls: str, lat: float, lng: float,
                      gs: str) -> dict:
    """Build a dispatch-shape stop dict for the synthetic 'No truck assigned'
    route. Mirrors the schema in process_dispatch.py so index.html can render
    it with zero special-case code (except the virtual-route filter for the
    polyline + depot markers)."""
    # 'tt' matches the dispatch trip-type strings so hover tooltips read
    # sensibly ("Delivery" / "Pick Up") rather than cls codes.
    tt = "Delivery" if cls == "D" else "Pick Up"
    stop = {
        "seq":   0,                          # no meaningful order within synth route
        "tt":    tt,
        "cls":   cls,
        "con":   contract,
        "cust":  info.get("cn") or info.get("vn") or "",
        "store": info.get("st") or "",
        # Net rental revenue — mirrors the dispatch stop schema. Rounded to 2dp
        # like real stops so tooltips don't show surprising precision.
        "rev":   round(float(info.get("tt") or 0), 2),
        "ll":    [round(lat, 6), round(lng, 6)],
        # On-time / promised arrival are unknown for undispatched stops — blank
        # strings match what real dispatch stops use when POR didn't populate.
        "ot":    "",
        "parr":  "",
        "aarr":  "",
        "addr":  info.get("da") or "",
    }
    if gs != "city_ref":
        stop["gs"] = gs
    return stop


def annotate(delivery_path: Path, dispatch_path: Path) -> dict:
    if not delivery_path.exists():
        raise SystemExit(f"delivery_data.json not found: {delivery_path}")
    if not dispatch_path.exists():
        raise SystemExit(f"dispatch_data.json not found: {dispatch_path}")

    with delivery_path.open(encoding="utf-8") as f:
        de = json.load(f)
    with dispatch_path.open(encoding="utf-8") as f:
        dd = json.load(f)

    contract_info = de.get("contractInfo") or {}
    city_ref = de.get("contractCityRef") or {}

    window = dd.get("window") or {}
    wstart = window.get("start")
    wend   = window.get("end")
    if not wstart or not wend:
        raise SystemExit("dispatch_data.json missing window.start/end")
    # Dates we actually have in dispatch output — used as the valid date
    # universe for synthetic routes. Stops on dates with zero real routes are
    # still allowed (the synthetic route will create the date entry).
    dates_in_window = set()
    cur = date.fromisoformat(wstart)
    end = date.fromisoformat(wend)
    while cur <= end:
        dates_in_window.add(cur.isoformat())
        cur += timedelta(days=1)

    contract_trucks = dd.get("contract_trucks") or {}

    # ---- Pass 1: find missing D/P stops per contract ----
    # Each missing stop -> (date, cls, contract)
    missing_by_date: dict[str, list[tuple[str, dict, str, str]]] = defaultdict(list)
    n_scanned = n_missing_d = n_missing_p = 0
    n_skip_no_coord = 0
    n_skip_cp_contract = 0
    n_skip_out_of_window = 0

    for contract, info in contract_info.items():
        n_scanned += 1
        dv = bool(info.get("dv"))
        pu = bool(info.get("pu"))

        if not dv and not pu:
            # Customer-pickup-only contract — never had a dispatch obligation,
            # so nothing to flag. (Legacy rows from pre-2026 contracts that
            # predate the dv/pu flags fall through here too; those won't be in
            # the dispatch window anyway.)
            n_skip_cp_contract += 1
            continue

        ct_entries = contract_trucks.get(contract) or []
        has_D = any(e[3] == "D" for e in ct_entries)
        has_P = any(e[3] == "P" for e in ct_entries)

        # Delivery leg
        if dv and not has_D:
            sd = info.get("sd")
            if sd and sd in dates_in_window:
                coord = _resolve_coord(info, city_ref)
                if coord is None:
                    n_skip_no_coord += 1
                else:
                    missing_by_date[sd].append((contract, info, "D", coord))
                    # coord tuple stored here; Pass 2 uses it directly (no second lookup)
            elif sd:
                n_skip_out_of_window += 1

        # Pickup leg
        if pu and not has_P:
            pd = info.get("pd")
            if pd and pd in dates_in_window:
                coord = _resolve_coord(info, city_ref)
                if coord is None:
                    n_skip_no_coord += 1
                else:
                    missing_by_date[pd].append((contract, info, "P", coord))
            elif pd:
                n_skip_out_of_window += 1

    # ---- Pass 2: build synthetic routes per (date, region bucket) and inject ----
    # Splitting by region (instead of one mega-route per date) so the UI's BU
    # filter cleanly excludes unassigned stops from other regions. Previously
    # a single all-region synth route was matched by every BU chip, leaking
    # East/West unassigned pins into a GTA-filtered view.
    routes = dd.get("routes") or {}
    n_stops_injected = 0
    n_routes_injected = 0
    dates_touched = []
    bucket_counts: dict = defaultdict(int)
    for d, rows in missing_by_date.items():
        if not rows:
            continue
        # Sort by (cls, customer name, contract) so the drill-down list reads
        # deterministically across refreshes.
        rows.sort(key=lambda r: (r[2], (r[1].get("cn") or "").upper(), r[0]))

        # Group by region bucket BEFORE building stops so each synth route
        # has homogeneous regional ownership.
        rows_by_bucket: dict = defaultdict(list)
        for contract, info, cls, coord in rows:
            bucket = _bucket_for(info)
            rows_by_bucket[bucket].append((contract, info, cls, coord))

        for bucket, bucket_rows in rows_by_bucket.items():
            stops = []
            for contract, info, cls, coord in bucket_rows:
                lat, lng, gs = coord  # already resolved in Pass 1 — no second lookup
                stops.append(_build_synth_stop(contract, info, cls, lat, lng, gs))

            if not stops:
                continue

            # Rebuild seq 1..N in sorted order within this region's synth route.
            for i, s in enumerate(stops):
                s["seq"] = i + 1

            # Deterministic rn per (date, bucket) so repeated annotate runs
            # don't churn ids; bucket offset prevents cross-bucket collisions.
            rn = SYNTH_RN_BASE + SYNTH_RN_OFFSET_BY_BUCKET.get(bucket, 0) + int(d.replace("-", ""))

            synth_route = {
                "rn":    rn,
                "truck": f"{SYNTH_TRUCK_PREFIX} ({bucket})",
                "tnum":  "",
                "drv":   "",
                "stps":  len(stops),
                "dist":  0,
                "time":  0,
                "load":  0,
                "hs":    "",         # no home store — skip depot fallback in index.html
                "co":    bucket,     # region bucket — read by index.html _routeCompanies
                "color": SYNTH_COLOR_BY_BUCKET.get(bucket, SYNTH_DEFAULT_COLOR),
                "virtual": 1,        # flag — index.html skips polyline + depot markers
                "stops": stops,
            }
            # Append rather than sort — keeps real routes at their existing order.
            routes.setdefault(d, []).append(synth_route)
            n_stops_injected += len(stops)
            n_routes_injected += 1
            bucket_counts[bucket] += len(stops)

            # Count for summary
            n_missing_d += sum(1 for s in stops if s["cls"] == "D")
            n_missing_p += sum(1 for s in stops if s["cls"] == "P")

        dates_touched.append((d, sum(len(b) for b in rows_by_bucket.values())))

    # Update the dates list in case we added a date that had no real routes
    existing_dates = set(dd.get("dates") or [])
    for d in missing_by_date:
        existing_dates.add(d)
    dd["dates"] = sorted(existing_dates)
    dd["routes"] = routes

    # Stamp annotate run in meta (doesn't exist yet on dispatch payload, so
    # add a small sub-block). Helpful for debugging "is this a refreshed file
    # or stale?".
    dd["annotate_unassigned"] = {
        "injected_stops":  n_stops_injected,
        "injected_routes": n_routes_injected,
        "injected_D":      n_missing_d,
        "injected_P":      n_missing_p,
        "dates_touched":   len(dates_touched),
        "by_bucket":       dict(bucket_counts),
        "contracts_scanned": n_scanned,
        "skip_cp_contract":  n_skip_cp_contract,
        "skip_out_of_window": n_skip_out_of_window,
        "skip_no_coord":   n_skip_no_coord,
    }

    print(f"[annotate] Scanned {n_scanned:,} contracts", file=sys.stderr)
    print(f"[annotate] Injected {n_stops_injected:,} synth stops "
          f"(D={n_missing_d:,} / P={n_missing_p:,}) across "
          f"{n_routes_injected:,} per-bucket routes / {len(dates_touched)} dates",
          file=sys.stderr)
    if bucket_counts:
        bucket_summary = ", ".join(
            f"{b}={n:,}" for b, n in sorted(bucket_counts.items(), key=lambda x: -x[1])
        )
        print(f"[annotate] Stops by bucket: {bucket_summary}", file=sys.stderr)
    print(f"[annotate] Skipped: cp-contract={n_skip_cp_contract:,} "
          f"out-of-window={n_skip_out_of_window:,} no-coord={n_skip_no_coord:,}",
          file=sys.stderr)
    if dates_touched:
        # Surface the top 5 worst days so Mark can triage
        dates_touched.sort(key=lambda x: -x[1])
        head = ", ".join(f"{d}({n})" for d, n in dates_touched[:5])
        print(f"[annotate] Top dates by missing-stop count: {head}", file=sys.stderr)

    # Atomic write — same pattern as process_dispatch.py
    tmp = dispatch_path.with_suffix(dispatch_path.suffix + ".tmp")
    data_str = json.dumps(dd, separators=(",", ":"))
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(data_str)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, dispatch_path)
    size_mb = dispatch_path.stat().st_size / 1024 / 1024
    print(f"[annotate] Wrote {dispatch_path} ({size_mb:.1f} MB)", file=sys.stderr)
    return dd


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("delivery_json", type=Path, help="Path to delivery_data.json")
    ap.add_argument("dispatch_json", type=Path, help="Path to dispatch_data.json (modified in place)")
    args = ap.parse_args()
    annotate(args.delivery_json, args.dispatch_json)
    return 0


if __name__ == "__main__":
    sys.exit(main())
