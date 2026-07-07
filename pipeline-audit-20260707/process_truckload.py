"""
process_truckload.py — estimate how full each truck is per leg (delivery / pickup).

Model (deliberately rough — "truck-space cube units", not real volume):
  1. Per contract, cube = Σ (category quantity × cube factor).  The category
     quantities come straight out of CMMC_UniqueEvents.xlsx (same columns
     process_cmmc.py already reads — we import its column-resolution logic so a
     layout change can't silently diverge the two).  Factors are tuned so a
     genuinely full standard truck ≈ ~1000 cube units: bulky items (chairs,
     tables, tents, equipment, bar, buffet) drive the cube; tableware / linen
     pack tight and contribute very little per piece.
  2. Join ~the dispatch window of routes: for each (date, truck) sum the cube of
     the contracts on its DELIVERY stops. Truck capacity = the p85 of *loaded*
     per-truck-day delivery cubes (near-empty days excluded), calibrated PER BU
     (GTA / GTA Tents / East / West) — a tent truck's "full" is nothing like a
     tableware truck's. GTA Tents is handled specially (thin/peaky data).
  3. Emit contract_load.json:
        { "generated_at": …,
          "capacity_by_bu": {"GTA": X, "GTA Tents": …, "East": …, "West": …},
          "cube_by_con": { "<con>": <cube>, … },
          "meta": { join rate, factors, capacity calibration + fill distribution } }

The tracker (and dashboard) then compute, per truck/day:
    delivery_cube = Σ cube_by_con[con]  over the truck's D stops
    pickup_cube   = Σ cube_by_con[con]  over the truck's P stops
    fill%         = leg_cube / capacity_by_bu[truck's BU]

Usage:
    python process_truckload.py <cmmc_xlsx> <dispatch_data.json> <out_contract_load.json>

Reports the category cube-share table, the dispatch↔CMMC join rate, the per-BU
capacities, and the resulting per-BU fill distribution to stdout.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

# Reuse process_cmmc's header-name → column resolution so this script reads the
# exact same category columns the map does. If the export layout shifts, both
# break the same way (loudly) instead of silently diverging.
import process_cmmc as pc


# ============================================================
# CUBE FACTORS — relative "truck-space units" per item.
# EDIT THESE to retune. Keyed by the category labels process_cmmc emits
# (see pc.CATEGORY_COLUMN_NAMES — left element of each tuple).
#
# Recalibrated 2026-06-27: the original factors over-weighted high-count compact
# items (linen/glass/flatware), so a handful of mega-count orders produced
# outlier loads that inflated the percentile capacity and made typical trucks
# read ~7–17% full. New scheme — bulky items drive cube; tableware/linen pack
# tight — targets a full standard truck ≈ ~1000 cube units.
# ============================================================
CUBE_FACTORS: dict[str, float] = {
    # --- bulky: these should dominate a truck's cube ---
    "Tent":                   120.0,
    "Tables":                 12.0,
    "Event Equipment":        8.0,
    "Buffet & Display":       4.0,
    "Bar & Beverage":         3.0,
    "Chairs":                 2.0,
    "Decor & Accessories":    2.0,   # spec "Decor"
    "Other":                  1.5,
    # --- compact / dense-packing: contribute little per piece ---
    "Platters & Trays":       0.15,
    "Serving":                0.15,
    "Dinnerware":             0.06,
    "Other Tableware":        0.06,
    "Linen":                  0.05,
    "Glassware":              0.05,
    "Disposables":            0.05,
    "Flatware":               0.02,
    # --- not named in the spec; sized to the same scheme ---
    "Furniture":              3.0,   # bulky lounge pieces
    "Kitchen & Food Service": 4.0,   # ovens / hot boxes / prep
    "Floors":                 2.0,   # subfloor / dance floor panels
    "Anchoring":              1.0,   # ballast / stakes for tents
    "Drape":                  0.05,  # fabric, packs tight like linen
    "Lighting":               0.30,
    # --- non-physical / accounting lines — never add cube ---
    "Services":               0.0,
    "Discount":               0.0,
    "Delivery":               0.0,
}
# Any category not listed above falls back to this (shouldn't happen — every
# pc category is covered — but keeps the math defined if a new one appears).
DEFAULT_FACTOR = 1.0

# ============================================================
# STORE CODE → business unit. Mirrors COMPANY_BUCKETS in process_cmmc.py and
# HS_TO_CO in tracker.html. We classify a truck-day's BU by the majority store
# code across its stops (more reliable than the route `hs`, which is blank on
# ~10% of routes). The tracker uses this same map so both sides agree.
# ============================================================
STORE_TO_BU: dict[str, str] = {
    "CFR": "GTA", "CMM": "GTA", "ERG": "GTA", "HIG": "GTA",
    "ATR": "GTA Tents", "RT": "GTA Tents",
    "MFL": "East",
    "A&B": "West", "LW": "West",
}
BU_ORDER = ["GTA", "GTA Tents", "East", "West"]

# ============================================================
# CAPACITY CALIBRATION — all tunable in one place.
# ============================================================
# Percentile of LOADED truck-days that defines "full" (≈100%) for each BU.
# The per-truck-day delivery-cube distribution is heavily right-skewed (a van of
# chairs vs a truck of tables+equipment both count as "one delivery" but differ
# 10×), so a high percentile can't make the MEDIAN read 55–85% without turning a
# third of days red. p80 is the chosen compromise: a well-loaded (upper-quartile)
# day reads ~80%, busy days approach/exceed 100%, and ~20% run over. Lower this
# (e.g. 75) for fuller readings at the cost of more >100% days.
CAPACITY_PERCENTILE = 80
# A truck-day counts as "loaded" (i.e. it really carried a delivery, not a token
# drop) if its delivery cube ≥ this fraction of the BU's median positive-day
# cube. Near-empty days below that are excluded from the capacity percentile so
# they don't drag the "full" line down.
LOADED_FRACTION = 0.40
# Bucket "trucks" that aren't a single physical vehicle. Like the synthetic
# "No truck assigned" routes, one bucket-day aggregates many vehicles' loads
# (audit 2026-07-07 found an "Unknown" day at 16,586 cube ≈ 9× capacity), so
# including them inflates the capacity percentile (~2.5% on GTA) and pollutes
# the fill distribution. Excluded from CAPACITY CALIBRATION only — their
# contracts still ship in cube_by_con and still count toward the join rate.
CAPACITY_EXCLUDE_TRUCKS = {"Unknown", "Large Runs", "Small Runs"}
# A BU needs at least this many LOADED real-truck delivery days to calibrate from
# its own data; otherwise it falls back (peer-BU average, or a manual default for
# tents whose cube scale is unlike any other BU).
MIN_LOADED_DAYS = 20
# Fallback "full" cube for GTA Tents when there's too little assigned-tent-truck
# data (its deliveries are almost all dispatched as "unassigned" in this window,
# and the per-tent factor of 120 puts it on a totally different scale from the
# tableware BUs, so a peer average would badly mis-call it). Low-confidence —
# flagged as such in the output meta.
TENTS_DEFAULT_CAPACITY = 2000.0


def percentile(values: list[float], pct: float) -> float:
    """Linear-interpolation percentile (numpy-free). `values` need not be sorted."""
    if not values:
        return 0.0
    s = sorted(values)
    if len(s) == 1:
        return s[0]
    rank = (pct / 100.0) * (len(s) - 1)
    lo = int(rank)
    hi = min(lo + 1, len(s) - 1)
    frac = rank - lo
    return s[lo] + (s[hi] - s[lo]) * frac


def build_cube_by_con(xlsx_path: Path):
    """Stream the CMMC export, summing cube per contract number.

    Returns a dict:
      cube_by_con      : {con: total_cube}
      top_cats_by_con  : {con: [[category, cube_contribution], …]}  (top 3)
      cats_by_con      : {con: {category: raw_qty, …}}  (ALL non-zero physical
                         categories — drives the tracker popup's "Items" row)
      cat_cube_total   : {category: Σ cube}        (diagnostic: cube share)
      cat_qty_total    : {category: Σ raw qty}      (diagnostic)
      tent_contract_cubes : [total_cube, …] for contracts carrying any Tent qty
    """
    print(f"[truckload] Opening {xlsx_path} (streaming)…", file=sys.stderr)
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    cube_by_con: dict[str, float] = defaultdict(float)
    cat_cube_by_con: dict[str, dict] = defaultdict(lambda: defaultdict(float))
    cat_qty_by_con: dict[str, dict] = defaultdict(lambda: defaultdict(float))
    cat_cube_total: dict[str, float] = defaultdict(float)
    cat_qty_total: dict[str, float] = defaultdict(float)
    contract_has_tent: set[str] = set()
    try:
        ws = wb["Unique_Events"]
        row_iter = ws.iter_rows(min_row=1, values_only=True)
        header = next(row_iter)
        # Binds pc.COL_CONTRACT, pc.COL_DOCUMENT_TYPE, pc.CATEGORY_COLUMNS, … on
        # the process_cmmc module globals.
        pc.resolve_columns(header)
        cat_cols = pc.CATEGORY_COLUMNS                # [(idx, label), …]
        col_contract = pc.COL_CONTRACT
        col_doc = pc.COL_DOCUMENT_TYPE

        n = 0
        for row in row_iter:
            n += 1
            # Keep only Contract documents — quotes/leads carry the same con but
            # aren't dispatched cube. (Matches process_cmmc's ALLOWED_DOC_TYPES.)
            if row[col_doc] not in pc.ALLOWED_DOC_TYPES:
                continue
            con = str(row[col_contract] or "").strip()
            if not con:
                continue
            for col_idx, label in cat_cols:
                v = row[col_idx]
                if v is None:
                    continue
                try:
                    qty = float(v)
                except (TypeError, ValueError):
                    continue
                if not qty:
                    continue
                cat_qty_total[label] += qty
                if label == "Tent":
                    contract_has_tent.add(con)
                factor = CUBE_FACTORS.get(label, DEFAULT_FACTOR)
                if factor <= 0:
                    continue
                # raw item count per contract/category (physical items only —
                # zero-factor accounting lines Services/Discount/Delivery are
                # skipped above) → feeds the "Items" row in the tracker popup.
                cat_qty_by_con[con][label] += qty
                contrib = qty * factor
                cube_by_con[con] += contrib
                cat_cube_by_con[con][label] += contrib
                cat_cube_total[label] += contrib
        print(f"[truckload]   scanned {n:,} rows, {len(cube_by_con):,} contracts with cube",
              file=sys.stderr)
    finally:
        wb.close()

    # Round cube; build top-categories list per contract.
    cube_out = {c: round(v, 1) for c, v in cube_by_con.items() if v > 0}
    top_cats: dict[str, list] = {}
    for c, cats in cat_cube_by_con.items():
        ranked = sorted(([name, round(q, 1)] for name, q in cats.items() if q > 0),
                        key=lambda x: -x[1])
        if ranked:
            top_cats[c] = ranked[:3]
    # Full per-contract raw item quantities (ALL non-zero physical categories),
    # sorted by quantity desc so the popup reads "164 Chairs · 53 Tables · …"
    # without re-sorting. Quantities rounded to whole items.
    cats_qty: dict[str, dict] = {}
    for c, cats in cat_qty_by_con.items():
        ranked = sorted(((name, int(round(q))) for name, q in cats.items()
                         if round(q) >= 1),
                        key=lambda x: -x[1])
        if ranked:
            cats_qty[c] = {name: q for name, q in ranked}
    tent_contract_cubes = [cube_by_con[c] for c in contract_has_tent
                           if cube_by_con.get(c, 0) > 0]
    return {
        "cube_by_con": cube_out,
        "top_cats_by_con": top_cats,
        "cats_by_con": cats_qty,
        "cat_cube_total": dict(cat_cube_total),
        "cat_qty_total": dict(cat_qty_total),
        "tent_contract_cubes": tent_contract_cubes,
    }


def truck_bu(stops: list) -> str | None:
    """Classify a truck-day's BU by the majority store code across its stops."""
    votes = Counter()
    for s in stops:
        bu = STORE_TO_BU.get(str(s.get("store") or "").strip().upper())
        if bu:
            votes[bu] += 1
    if not votes:
        return None
    return votes.most_common(1)[0][0]


def _dist(fills: list[float]) -> dict:
    """median / p75 / p90 / %>100 for a list of fill percentages."""
    if not fills:
        return {"n": 0}
    return {
        "n": len(fills),
        "median": round(percentile(fills, 50), 1),
        "p75": round(percentile(fills, 75), 1),
        "p90": round(percentile(fills, 90), 1),
        "pct_over_100": round(100 * sum(1 for f in fills if f > 100) / len(fills), 1),
    }


def calibrate(dispatch: dict, cube_by_con: dict):
    """Join routes → per-(date,truck) delivery cube → per-BU capacity.

    Capacity = p{CAPACITY_PERCENTILE} of LOADED real-truck-day delivery cubes.
    BUs without enough assigned-truck data fall back (peer-BU average, or a
    manual default for tents). Returns (capacity_by_bu, diag, dispatch_cons).
    """
    routes = dispatch.get("routes", {})
    pos_by_bu: dict[str, list] = defaultdict(list)   # bu -> positive d_cube days

    dispatch_cons: set[str] = set()
    matched_cons: set[str] = set()

    for date, day_routes in routes.items():
        # Collapse to per-truck (a truck can appear as >1 route object/day).
        by_truck: dict[str, dict] = {}
        for r in day_routes:
            tkey = str(r.get("tnum") or r.get("truck") or "").replace("#", "").strip()
            if not tkey:
                continue
            # Skip synthetic "No truck assigned (…)" buckets: they aren't a real
            # vehicle — annotate_unassigned bundles every unassigned stop for a
            # BU/day into one pseudo-truck, whose huge aggregate cube would
            # otherwise dominate the capacity percentile. (The tracker hides
            # their fill icons for the same reason.)
            truck_name = str(r.get("truck") or "").strip()
            if truck_name.lower().startswith("no truck assigned"):
                continue
            # Bucket trucks (Unknown / Large Runs / Small Runs) aggregate many
            # vehicles — count their contracts for the join, but keep their
            # stops out of the per-truck-day capacity distribution.
            is_bucket = truck_name in CAPACITY_EXCLUDE_TRUCKS
            slot = None if is_bucket else by_truck.setdefault(tkey, {"stops": [], "d_cube": 0.0})
            for s in r.get("stops", []):
                if slot is not None:
                    slot["stops"].append(s)
                con = str(s.get("con") or "").strip()
                if not con:
                    continue
                dispatch_cons.add(con)
                cube = cube_by_con.get(con)
                if cube is None:
                    continue
                matched_cons.add(con)
                if slot is not None and s.get("cls") == "D":
                    slot["d_cube"] += cube

        for tkey, slot in by_truck.items():
            if slot["d_cube"] <= 0:
                continue
            bu = truck_bu(slot["stops"])
            if bu is not None:
                pos_by_bu[bu].append(slot["d_cube"])

    def loaded_of(pos: list) -> list:
        if not pos:
            return []
        med = percentile(pos, 50)
        floor = LOADED_FRACTION * med
        return [x for x in pos if x >= floor]

    # First pass: calibrate every BU that has enough real loaded truck-days.
    capacity_by_bu: dict[str, float] = {}
    loaded_by_bu: dict[str, list] = {}
    sample_sizes: dict[str, int] = {}
    cap_source: dict[str, str] = {}
    for bu in BU_ORDER:
        pos = pos_by_bu.get(bu, [])
        sample_sizes[bu] = len(pos)
        loaded = loaded_of(pos)
        loaded_by_bu[bu] = loaded
        if len(loaded) >= MIN_LOADED_DAYS:
            capacity_by_bu[bu] = round(percentile(loaded, CAPACITY_PERCENTILE), 1)
            cap_source[bu] = f"p{CAPACITY_PERCENTILE} of {len(loaded)} loaded truck-days"

    # Second pass: fill in the data-poor BUs.
    peer_caps = [capacity_by_bu[b] for b in capacity_by_bu]
    peer_avg = round(sum(peer_caps) / len(peer_caps), 1) if peer_caps else TENTS_DEFAULT_CAPACITY
    for bu in BU_ORDER:
        if bu in capacity_by_bu:
            continue
        nl = len(loaded_by_bu.get(bu, []))
        if bu == "GTA Tents":
            capacity_by_bu[bu] = TENTS_DEFAULT_CAPACITY
            cap_source[bu] = (f"MANUAL default {TENTS_DEFAULT_CAPACITY:.0f} — only {nl} "
                              f"loaded assigned-tent-truck days (deliveries dispatched "
                              f"as unassigned); low confidence")
        else:
            capacity_by_bu[bu] = peer_avg
            cap_source[bu] = (f"FALLBACK peer-BU average {peer_avg:.0f} — only {nl} "
                              f"loaded real-truck days; low confidence")

    # --- fill distribution per BU over its loaded real-truck-days ---
    fill_dist: dict[str, dict] = {}
    for bu in BU_ORDER:
        cap = capacity_by_bu.get(bu)
        loaded = loaded_by_bu.get(bu, [])
        if cap and loaded:
            fill_dist[bu] = _dist([100 * x / cap for x in loaded])
        else:
            fill_dist[bu] = {"n": len(loaded)}

    join_rate = (len(matched_cons) / len(dispatch_cons)) if dispatch_cons else 0.0
    diag = {
        "dispatch_distinct_con": len(dispatch_cons),
        "matched_con": len(matched_cons),
        "join_rate": round(join_rate, 4),
        "capacity_percentile": CAPACITY_PERCENTILE,
        "loaded_fraction": LOADED_FRACTION,
        "min_loaded_days": MIN_LOADED_DAYS,
        "truck_day_samples_by_bu": sample_sizes,
        "loaded_day_counts_by_bu": {bu: len(loaded_by_bu.get(bu, [])) for bu in BU_ORDER},
        "capacity_source_by_bu": cap_source,
        "fill_distribution_by_bu": fill_dist,
    }
    return capacity_by_bu, diag, dispatch_cons


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmmc_xlsx", type=Path, help="CMMC_UniqueEvents.xlsx")
    ap.add_argument("dispatch_json", type=Path, help="dispatch_data.json")
    ap.add_argument("out", type=Path, help="Output contract_load.json")
    args = ap.parse_args()

    if not args.cmmc_xlsx.exists():
        print(f"ERROR: {args.cmmc_xlsx} not found", file=sys.stderr)
        return 1
    if not args.dispatch_json.exists():
        print(f"ERROR: {args.dispatch_json} not found", file=sys.stderr)
        return 1

    built = build_cube_by_con(args.cmmc_xlsx)
    cube_by_con = built["cube_by_con"]
    top_cats = built["top_cats_by_con"]
    cats_by_con = built["cats_by_con"]

    with args.dispatch_json.open(encoding="utf-8") as f:
        dispatch = json.load(f)

    capacity_by_bu, diag, dispatch_cons = calibrate(dispatch, cube_by_con)

    # Ship only the cube for contracts the dispatch window actually references —
    # the tracker never asks about the other ~130k. Keeps the browser payload
    # small (~0.3 MB vs ~9 MB for the full book).
    cube_ship = {c: cube_by_con[c] for c in dispatch_cons if c in cube_by_con}
    top_cats_ship = {c: top_cats[c] for c in dispatch_cons if c in top_cats}
    cats_ship = {c: cats_by_con[c] for c in dispatch_cons if c in cats_by_con}

    # Category cube-share diagnostic (uses the current factors).
    total_cube = sum(built["cat_cube_total"].values()) or 1.0
    cat_share = sorted(
        ([cat, cube, 100 * cube / total_cube, built["cat_qty_total"].get(cat, 0.0)]
         for cat, cube in built["cat_cube_total"].items()),
        key=lambda x: -x[1],
    )

    payload = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "capacity_by_bu": capacity_by_bu,
        "cube_by_con": cube_ship,
        "top_cats_by_con": top_cats_ship,
        "cats_by_con": cats_ship,
        "meta": {
            "cube_factors": CUBE_FACTORS,
            "store_to_bu": STORE_TO_BU,
            "cat_cube_share": [
                {"category": c, "cube": round(cube, 0), "share_pct": round(sh, 1),
                 "qty": round(q, 0)}
                for c, cube, sh, q in cat_share
            ],
            **diag,
        },
    }

    # Atomic write (serialize → fsync → rename) — same guard process_cmmc uses.
    args.out.parent.mkdir(parents=True, exist_ok=True)
    tmp = args.out.with_suffix(args.out.suffix + ".tmp")
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, args.out)

    size_mb = args.out.stat().st_size / (1024 * 1024)
    print()
    print(f"Wrote {args.out} ({size_mb:.2f} MB)")
    print(f"  Contracts with cube (all CMMC): {len(cube_by_con):,}")
    print(f"  Contracts shipped (dispatch-referenced): {len(cube_ship):,}")
    print(f"  Join rate (dispatch con found in CMMC): {diag['join_rate']*100:.1f}% "
          f"({diag['matched_con']:,}/{diag['dispatch_distinct_con']:,})")

    print()
    print("CATEGORY CUBE SHARE (current factors):")
    print(f"  {'Category':<24}{'Cube':>14}{'Share':>9}{'TotalQty':>14}")
    for c, cube, sh, q in cat_share:
        if cube <= 0:
            continue
        print(f"  {c:<24}{cube:>14,.0f}{sh:>8.1f}%{q:>14,.0f}")

    print()
    print(f"PER-BU CAPACITY (p{CAPACITY_PERCENTILE} of loaded real-truck-day delivery cube; "
          f"loaded = ≥{LOADED_FRACTION:.0%} of BU median):")
    for bu in BU_ORDER:
        cap = capacity_by_bu.get(bu)
        n = diag["truck_day_samples_by_bu"].get(bu, 0)
        nl = diag["loaded_day_counts_by_bu"].get(bu, 0)
        cap_s = f"{cap:,.1f}" if cap is not None else "—"
        print(f"    {bu:<10} cap={cap_s:>12}   ({n:,} positive / {nl:,} loaded real truck-days)")
        print(f"               source: {diag['capacity_source_by_bu'].get(bu,'')}")

    print()
    print("FILL DISTRIBUTION over loaded truck-days (target: median ~55–85%, small tail >100%):")
    print(f"  {'BU':<10}{'n':>6}{'median':>9}{'p75':>8}{'p90':>8}{'%>100':>8}")
    for bu in BU_ORDER:
        d = diag["fill_distribution_by_bu"].get(bu, {})
        if d.get("n"):
            print(f"  {bu:<10}{d['n']:>6}{d.get('median','—'):>8}%{d.get('p75','—'):>7}%"
                  f"{d.get('p90','—'):>7}%{d.get('pct_over_100','—'):>7}%")
        else:
            print(f"  {bu:<10}{'(no loaded truck-days)':>40}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
