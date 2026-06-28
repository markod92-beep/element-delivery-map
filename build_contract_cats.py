#!/usr/bin/env python3
"""Add `cats_by_con` (per-contract non-zero item-category quantities) to
contract_load.json so the tracker's quote popup can show an item summary
(e.g. "164 Chairs · 53 Tables · 15 Event Equipment").

The authoritative source for per-category quantities is CMMC_UniqueEvents.xlsx,
read by the truck-fill model that produces contract_load.json. That workbook is
NOT kept in this repo, so this script derives the quantities from the cube
breakdown the model already emitted: top_cats_by_con stores, per contract, the
top categories as [category, cube] where cube = quantity x cube_factor. Dividing
by meta.cube_factors recovers the original integer quantity.

Limitation vs. a full rebuild from the xlsx: top_cats_by_con only keeps each
contract's top ~3 categories by cube, so cats_by_con covers those. To capture
every non-zero category, re-run the real contract_load builder against
CMMC_UniqueEvents.xlsx and emit cats_by_con there directly.

Usage:  python build_contract_cats.py        # rewrites contract_load.json in place
"""
import json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, "contract_load.json")


def main():
    with open(PATH, "r", encoding="utf-8") as f:
        d = json.load(f)

    factors = (d.get("meta") or {}).get("cube_factors") or {}
    top = d.get("top_cats_by_con") or {}
    if not top or not factors:
        sys.exit("contract_load.json is missing top_cats_by_con or meta.cube_factors")

    cats_by_con = {}
    for con, cats in top.items():
        rec = {}
        for name, cube in cats:
            fac = factors.get(name, 0)
            if fac <= 0:
                continue
            qty = round(cube / fac)
            if qty > 0:
                rec[name] = qty
        if rec:
            # store sorted by quantity desc (the popup re-sorts, this is just tidy)
            cats_by_con[con] = dict(sorted(rec.items(), key=lambda kv: -kv[1]))

    d["cats_by_con"] = cats_by_con
    with open(PATH, "w", encoding="utf-8") as f:
        json.dump(d, f, separators=(",", ":"), ensure_ascii=False)

    print(f"cats_by_con written for {len(cats_by_con)} contracts "
          f"(of {len(top)} with cube breakdowns).")


if __name__ == "__main__":
    main()
