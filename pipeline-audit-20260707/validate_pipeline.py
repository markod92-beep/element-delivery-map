#!/usr/bin/env python3
"""validate_pipeline.py — pre-deploy validator for daily_refresh.ps1.

Called by daily_refresh.ps1 instead of inline python -c strings, so paths
with special characters (apostrophes, spaces, etc.) are handled safely by
subprocess argument passing rather than PowerShell string interpolation.

Usage:
    python validate_pipeline.py delivery  <path>
    python validate_pipeline.py dispatch  <path>
    python validate_pipeline.py venues    <path>

Exit 0 on success; exit 1 with a message on failure.
"""
import json
import re
import sys
from pathlib import Path

# Internal-leg customer patterns that must never appear on a shipped stop —
# mirrors is_internal_stop() in process_dispatch.py. Audit guard 2026-07-07:
# catches filter under-reach (transfers/internal legs leaking back in) the
# day it regresses instead of on the live map.
_INTERNAL_STORE_RE = re.compile(r"\bSTR\s*#?\s*0\d")
_INTERNAL_MARKERS = ("SCHEDULED STOP", "STOP AT STORE", "ELEMENT EVENT SOLUTIONS")
_EXCLUDED_TRUCKS = {"Transfers", "Product for Internal use"}


def _stop_is_internal(cust: str) -> bool:
    c = (cust or "").upper()
    if not c:
        return False
    if any(m in c for m in _INTERNAL_MARKERS):
        return True
    return bool(_INTERNAL_STORE_RE.search(c))


def main() -> int:
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <delivery|dispatch|venues> <path>", file=sys.stderr)
        return 1

    kind = sys.argv[1]
    path = Path(sys.argv[2])

    if not path.exists():
        print(f"ERROR: file not found: {path}", file=sys.stderr)
        return 1

    try:
        raw = path.read_bytes().rstrip(b"\x00").rstrip()
        j = json.loads(raw.decode("utf-8"))
    except Exception as e:
        print(f"ERROR: JSON parse failed for {path}: {type(e).__name__}: {e}", file=sys.stderr)
        return 1

    if kind == "delivery":
        m = j.get("meta") or {}
        gen_at = m.get("generated_at")
        if not gen_at:
            print("ERROR: missing meta.generated_at", file=sys.stderr)
            return 1
        rk = m.get("rows_kept", 0)
        if not rk:
            print("ERROR: rows_kept=0 — file likely truncated or empty", file=sys.stderr)
            return 1
        print(f"ok  generated_at={gen_at}  rows_kept={rk:,}")

    elif kind == "dispatch":
        dates = j.get("dates") or []
        if not dates:
            print("ERROR: no dates in dispatch output", file=sys.stderr)
            return 1
        window = j.get("window") or {}

        # Filter-regression gates (audit 2026-07-07). Real routes only —
        # synthetic "No truck assigned" routes (virtual=1) are skipped so this
        # validator works both before AND after annotate_unassigned.py runs.
        n_routes = n_stops = 0
        n_cls_t = n_internal = n_empty = 0
        bad_trucks = set()
        for day_routes in (j.get("routes") or {}).values():
            for r in day_routes or []:
                if r.get("virtual"):
                    continue
                n_routes += 1
                stops = r.get("stops") or []
                if not stops:
                    n_empty += 1
                if (r.get("truck") or "") in _EXCLUDED_TRUCKS:
                    bad_trucks.add(r.get("truck"))
                for s in stops:
                    n_stops += 1
                    if s.get("cls") == "T":
                        n_cls_t += 1
                    if _stop_is_internal(s.get("cust")):
                        n_internal += 1
        errs = []
        if n_cls_t:
            errs.append(f"{n_cls_t} Store-Transfer (cls=T) stops leaked into output")
        if n_internal:
            errs.append(f"{n_internal} internal-leg stops leaked (SCHEDULED STOP / STOP AT STORE / Element store)")
        if n_empty:
            errs.append(f"{n_empty} routes shipped with zero stops (should be dropped)")
        if bad_trucks:
            errs.append(f"excluded truck bucket(s) present: {sorted(bad_trucks)}")
        if errs:
            for e in errs:
                print(f"ERROR: {e}", file=sys.stderr)
            return 1
        print(f"ok  dates={len(dates):,}  window={window}  routes={n_routes:,}  stops={n_stops:,}  "
              f"(0 transfer/internal leaks, 0 empty routes)")

    elif kind == "venues":
        venues = j.get("venues") or []
        if not venues:
            print("ERROR: no venues emitted", file=sys.stderr)
            return 1
        gen_at = j.get("generated_at", "")
        print(f"ok  venues={len(venues):,}  generated_at={gen_at}")

    else:
        print(f"ERROR: unknown kind '{kind}' — expected delivery|dispatch|venues", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
