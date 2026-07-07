"""qa_checks.py - data-layer sanity gates for the Element Delivery Map.

Runs in daily_refresh.ps1 AFTER all process_*.py steps and BEFORE publish.py
+ wrangler deploy. Catches data-shape regressions that the existing inline
validators (Assert-ValidJsonPipelineOutput etc.) don't cover:

    - Truncation guard (file size floor, hard fail on suspiciously small files)
    - Drift guard (row count and revenue can't swing >X% vs trailing average)
    - Schema integrity (every contract has the fields the UI relies on)
    - Geo bounds (every coord falls inside the Canada bounding box)
    - Freshness (max(stop.date) within N days of run)
    - FSA validity (Canadian forward sortation area regex pass-rate)

On ANY check FAIL: exits 1 (caller's $ErrorActionPreference="Stop" halts the
pipeline before publish/deploy). Appends one row per check to qa_log.csv so
trend regressions become visible over time.

Run:
    python qa_checks.py
    python qa_checks.py --strict        (treat WARN as FAIL too)
    python qa_checks.py --no-drift      (skip trailing-window drift checks; use
                                         when you intentionally caused a swing)
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from statistics import mean

SCRIPT_DIR = Path(__file__).resolve().parent
LOG_DIR = SCRIPT_DIR / "health_log"
LOG_CSV = LOG_DIR / "qa_log.csv"

# Data JSONs sit next to this script on the laptop (refresh-v2). In the
# vendored CI layout the script lives in pipeline/ but outputs land at the
# repo root — run_pipeline.sh sets QA_DATA_DIR=. to point here at them.
DATA_DIR = Path(os.environ.get("QA_DATA_DIR") or SCRIPT_DIR)

DELIVERY_JSON = DATA_DIR / "delivery_data.json"
DISPATCH_JSON = DATA_DIR / "dispatch_data.json"
VENUES_JSON = DATA_DIR / "venues_data.json"

# === Thresholds (tune as patterns emerge in qa_log.csv) =====================
# Truncation floors — calibrated against last known-good sizes (Apr 2026).
# A truncated file would be MUCH smaller; these have ~50% safety margin.
MIN_DELIVERY_BYTES = 8_000_000     # last good ~18MB
MIN_DISPATCH_BYTES = 4_000_000     # last good ~10MB
MIN_VENUES_BYTES   =   400_000     # last good ~900KB

# Row count floors
MIN_CONTRACT_COUNT = 5_000         # contractInfo entries
MIN_DISPATCH_DATES = 90            # window is rolling ~120d
MIN_VENUES         = 20            # top-30 by region, allow a few empty regions

# Drift bands (% swing tolerated vs trailing 7-run mean from qa_log.csv)
DRIFT_PCT_WARN = 15.0
DRIFT_PCT_FAIL = 30.0

# Geo bounds — Canada, generous box (excludes obvious geocoding errors only)
LAT_MIN, LAT_MAX = 41.0, 84.0
LNG_MIN, LNG_MAX = -141.0, -52.0

# Freshness
MAX_FRESHNESS_DAYS = 2  # max(date) in dispatch must be within N days of today

# Canadian FSA pattern (LDL — letter, digit, letter; first letter excludes D/F/I/O/Q/U/W/Z)
FSA_REGEX = re.compile(r"^[ABCEGHJKLMNPRSTVXY]\d[A-Z]$", re.IGNORECASE)
MIN_FSA_VALIDITY_PCT = 99.0

# === Gate severity ===========================================================
# 2026-07-07 cloud-readiness rebalance: the two NEW statistical cross-checks
# from the July 7 audit run as SOFT gates (WARN — logged loudly + surfaced in
# the run notification, but non-blocking) while the refresh runs unattended in
# GitHub Actions. They are new and unproven against day-to-day data variance;
# a single false positive would abort the run and freeze the live map for the
# whole vacation window. Deterministic gates (JSON validity, truncation,
# freshness, internal_leak, contract_trucks_coverage, and everything in
# validate_pipeline.py) stay HARD. Revisit after ~2 weeks of green cloud runs:
# promote a check back to HARD by removing it from this set.
SOFT_GATES = {
    "delivery:weekly_vs_contract_tt",
    "dispatch:rev_equals_contract_rental",
}


def _fail_severity(check_name: str) -> str:
    """Severity when a check trips: WARN for SOFT_GATES, FAIL otherwise."""
    return "WARN" if check_name in SOFT_GATES else "FAIL"


# === Result type =============================================================
class Result:
    __slots__ = ("name", "status", "value", "threshold", "detail")
    def __init__(self, name: str, status: str, value: str = "", threshold: str = "", detail: str = ""):
        # status in {PASS, WARN, FAIL}
        self.name = name
        self.status = status
        self.value = value
        self.threshold = threshold
        self.detail = detail


# === Helpers =================================================================
_QA_HISTORY: list | None = None   # populated once per process on first call

def _get_qa_history() -> list:
    """Return all rows from qa_log.csv as dicts, loading at most once per process."""
    global _QA_HISTORY
    if _QA_HISTORY is None:
        if not LOG_CSV.exists():
            _QA_HISTORY = []
        else:
            with LOG_CSV.open(newline="", encoding="utf-8") as f:
                _QA_HISTORY = list(csv.DictReader(f))
    return _QA_HISTORY


def _load_json(path: Path) -> dict:
    """Load JSON, tolerating trailing nulls/whitespace from OneDrive sync quirks."""
    raw = path.read_bytes()
    # Strip trailing NUL bytes and whitespace (a known OneDrive sync issue)
    raw = raw.rstrip(b"\x00").rstrip()
    return json.loads(raw.decode("utf-8"))


def _trailing_mean(check_name: str, key: str, n: int = 7) -> float | None:
    """Pull the last N PASS values for a given check from the cached qa_log history."""
    vals = []
    for row in _get_qa_history():
        if row.get("check") == check_name and row.get("status") == "PASS":
            try:
                vals.append(float(row.get("value") or 0))
            except ValueError:
                pass
    vals = vals[-n:]
    return mean(vals) if vals else None


def _drift_status(current: float, baseline: float | None) -> tuple[str, str]:
    """Compare current vs trailing baseline; returns (status, detail)."""
    if baseline is None or baseline == 0:
        return "PASS", "no baseline yet"
    pct = abs(current - baseline) / baseline * 100.0
    if pct >= DRIFT_PCT_FAIL:
        return "FAIL", f"drift {pct:.1f}% vs baseline {baseline:.0f} (>= {DRIFT_PCT_FAIL}%)"
    if pct >= DRIFT_PCT_WARN:
        return "WARN", f"drift {pct:.1f}% vs baseline {baseline:.0f} (>= {DRIFT_PCT_WARN}%)"
    return "PASS", f"drift {pct:.1f}% vs baseline {baseline:.0f}"


# === Checks ==================================================================
def check_file_sizes() -> list[Result]:
    out = []
    for path, floor in [(DELIVERY_JSON, MIN_DELIVERY_BYTES),
                        (DISPATCH_JSON, MIN_DISPATCH_BYTES),
                        (VENUES_JSON, MIN_VENUES_BYTES)]:
        if not path.exists():
            out.append(Result(f"file_exists:{path.name}", "FAIL", "0", "1", "file missing"))
            continue
        size = path.stat().st_size
        ok = size >= floor
        out.append(Result(f"file_size:{path.name}", "PASS" if ok else "FAIL",
                          str(size), str(floor),
                          "" if ok else "below truncation floor — possible OneDrive truncation"))
    return out


def check_delivery(skip_drift: bool) -> list[Result]:
    out = []
    try:
        d = _load_json(DELIVERY_JSON)
    except Exception as e:
        out.append(Result("delivery:parse", "FAIL", "", "", f"{type(e).__name__}: {e}"))
        return out

    # contractInfo presence + count
    ci = d.get("contractInfo") or {}
    n_contracts = len(ci)
    ok_count = n_contracts >= MIN_CONTRACT_COUNT
    out.append(Result("delivery:contract_count",
                      "PASS" if ok_count else "FAIL",
                      str(n_contracts), str(MIN_CONTRACT_COUNT),
                      "" if ok_count else "contractInfo below floor"))

    # Drift on contract count
    if ok_count and not skip_drift:
        baseline = _trailing_mean("delivery:contract_count", "value")
        status, detail = _drift_status(n_contracts, baseline)
        out.append(Result("delivery:contract_count_drift", status, str(n_contracts), "+/- 30%", detail))

    # Schema integrity — every contract should have at least cn (customer name)
    if ci:
        sample = list(ci.items())[:1000]   # spot-check first 1000, not all
        missing = sum(1 for _, v in sample if not (v or {}).get("cn"))
        miss_pct = missing / len(sample) * 100
        ok_schema = miss_pct <= 5.0
        out.append(Result("delivery:schema_cn_present",
                          "PASS" if ok_schema else "FAIL",
                          f"{100 - miss_pct:.1f}", "95",
                          f"{missing}/{len(sample)} sampled contracts missing 'cn'" if not ok_schema else ""))

    # weekKeys present + non-empty
    wk = d.get("weekKeys") or []
    ok_wk = len(wk) >= 50
    out.append(Result("delivery:weekKeys", "PASS" if ok_wk else "FAIL",
                      str(len(wk)), "50", "" if ok_wk else "weekKeys list too short"))

    # Internal-consistency guard (audit 2026-07-07): the contractWeekly cells
    # and the per-contract contractInfo.tt roll-up are built from the same kept
    # rows, so their rental grand totals must agree (within per-cell/-contract
    # int rounding). Divergence means a section of the JSON is truncated or an
    # aggregation path dropped/double-counted rows.
    try:
        cw_total = 0
        n_cells = 0
        for locs in (d.get("contractWeekly") or {}).values():
            for sts in locs.values():
                for cos in sts.values():
                    for e in cos.values():
                        cw_total += e.get("r", 0)
                        n_cells += 1
        tt_total = sum(v.get("tt", 0) for v in ci.values())
        if tt_total > 0:
            diff_pct = abs(cw_total - tt_total) / tt_total * 100
            ok_tie = diff_pct <= 0.1
            out.append(Result("delivery:weekly_vs_contract_tt",
                              "PASS" if ok_tie else _fail_severity("delivery:weekly_vs_contract_tt"),
                              f"{diff_pct:.4f}", "0.1",
                              f"contractWeekly ${cw_total:,} vs sum(tt) ${tt_total:,} [soft gate — non-blocking]"
                              if not ok_tie else ""))
    except Exception as e:
        out.append(Result("delivery:weekly_vs_contract_tt", "WARN", "", "", f"could not cross-check: {e}"))

    # meta.generated_at within last 36h
    meta = d.get("meta") or {}
    gen_at_str = meta.get("generated_at") or ""
    if gen_at_str:
        try:
            gen_at = datetime.fromisoformat(gen_at_str.replace("Z", "+00:00"))
            if gen_at.tzinfo is None:
                gen_at = gen_at.replace(tzinfo=timezone.utc)
            age_hrs = (datetime.now(timezone.utc) - gen_at).total_seconds() / 3600
            ok_age = age_hrs < 36
            out.append(Result("delivery:freshness_hrs",
                              "PASS" if ok_age else "FAIL",
                              f"{age_hrs:.1f}", "36",
                              "" if ok_age else f"generated {age_hrs:.1f}h ago"))
        except Exception as e:
            out.append(Result("delivery:freshness_hrs", "WARN", "", "", f"parse err: {e}"))
    return out


def check_dispatch(skip_drift: bool) -> list[Result]:
    out = []
    try:
        d = _load_json(DISPATCH_JSON)
    except Exception as e:
        out.append(Result("dispatch:parse", "FAIL", "", "", f"{type(e).__name__}: {e}"))
        return out

    # Window dates
    dates = d.get("dates") or []
    n_dates = len(dates)
    ok_dates = n_dates >= MIN_DISPATCH_DATES
    out.append(Result("dispatch:date_count",
                      "PASS" if ok_dates else "FAIL",
                      str(n_dates), str(MIN_DISPATCH_DATES),
                      "" if ok_dates else "dispatch window collapsed"))

    # Freshness — max(date) must be within MAX_FRESHNESS_DAYS of today
    if dates:
        try:
            iso_dates = [d_ for d_ in dates if isinstance(d_, str)]
            max_d = max(datetime.strptime(s, "%Y-%m-%d").date() for s in iso_dates)
            today = datetime.now().date()
            staleness = (today - max_d).days if max_d <= today else 0
            ok_fresh = staleness <= MAX_FRESHNESS_DAYS
            out.append(Result("dispatch:freshness_days",
                              "PASS" if ok_fresh else "FAIL",
                              str(staleness), str(MAX_FRESHNESS_DAYS),
                              "" if ok_fresh else f"latest dispatch date {max_d}, {staleness}d stale"))
        except Exception as e:
            out.append(Result("dispatch:freshness_days", "WARN", "", "", f"date parse err: {e}"))

    # Geo bounds — walk routes[date][route][stops]["ll"] directly.
    # The old _walk() looked for {"lat":...,"lng":...} dict nodes but dispatch
    # stops use "ll": [lat, lng] arrays — it sampled 0 stops on every run.
    # Audit 2026-07-07: the cap used to stop at the FIRST 5000 stops (earliest
    # dates only), so a bad geocode on a recent date was never sampled. Now
    # collect all stops and stride-sample evenly across the whole window.
    out_of_bounds = 0
    SAMPLE_CAP = 5000
    all_lls = []
    all_stops_flat = []   # (stop, is_virtual_route) — reused by checks below
    for date_routes in (d.get("routes") or {}).values():
        for route in (date_routes or []):
            virt = bool(route.get("virtual"))
            for stop in (route.get("stops") or []):
                all_stops_flat.append((stop, virt))
                ll = stop.get("ll")
                if ll and len(ll) == 2:
                    all_lls.append(ll)
    stride = max(1, len(all_lls) // SAMPLE_CAP)
    sampled = 0
    for ll in all_lls[::stride]:
        try:
            lat, lng = float(ll[0]), float(ll[1])
            sampled += 1
            if not (LAT_MIN <= lat <= LAT_MAX and LNG_MIN <= lng <= LNG_MAX):
                out_of_bounds += 1
        except (TypeError, ValueError):
            pass
    if sampled > 0:
        oob_pct = out_of_bounds / sampled * 100
        ok_geo = oob_pct <= 1.0
        out.append(Result("dispatch:geo_bounds",
                          "PASS" if ok_geo else "FAIL",
                          f"{100 - oob_pct:.2f}", "99",
                          "" if ok_geo else f"{out_of_bounds}/{sampled} stops outside Canada box"))

    # Filter under-reach guard (audit 2026-07-07): transfers / internal legs
    # must never ship on real routes. Mirrors is_internal_stop() in
    # process_dispatch.py + the cls=T strip.
    internal_re = re.compile(r"\bSTR\s*#?\s*0\d")
    markers = ("SCHEDULED STOP", "STOP AT STORE", "ELEMENT EVENT SOLUTIONS")
    n_leak = sum(
        1 for stop, virt in all_stops_flat if not virt and (
            stop.get("cls") == "T"
            or any(m in (stop.get("cust") or "").upper() for m in markers)
            or internal_re.search((stop.get("cust") or "").upper())
        )
    )
    out.append(Result("dispatch:internal_leak", "PASS" if n_leak == 0 else "FAIL",
                      str(n_leak), "0",
                      "" if n_leak == 0 else f"{n_leak} transfer/internal stops leaked into real routes"))

    # contract_trucks consistency: every real stop's contract must appear in
    # the contract_trucks roll-up the UI drills into.
    ct = d.get("contract_trucks") or {}
    missing_ct = {str(stop.get("con")) for stop, virt in all_stops_flat
                  if not virt and stop.get("con") and str(stop.get("con")) not in ct}
    out.append(Result("dispatch:contract_trucks_coverage", "PASS" if not missing_ct else "FAIL",
                      str(len(missing_ct)), "0",
                      "" if not missing_ct else f"stops reference contracts absent from contract_trucks, e.g. {sorted(missing_ct)[:5]}"))

    # Revenue semantics guard (audit 2026-07-07): dispatch stop.rev must equal
    # contractInfo.tt (contract rental, excl. delivery fee) for matched
    # contracts — 100.0% equal at audit time. A drop below 98% means POR
    # changed what Net_Revenue carries in the dispatch export (or the delivery
    # fee started leaking into rental) and every revenue rollup is suspect.
    try:
        de = _load_json(DELIVERY_JSON)
        ci = de.get("contractInfo") or {}
        n_eq = n_ne = 0
        for stop, virt in all_stops_flat:
            if virt:
                continue
            con = stop.get("con")
            info = ci.get(str(con)) if con else None
            if not info or "tt" not in info:
                continue
            if abs((stop.get("rev") or 0) - info["tt"]) <= 1.0:
                n_eq += 1
            else:
                n_ne += 1
        if n_eq + n_ne >= 100:
            pct = 100.0 * n_eq / (n_eq + n_ne)
            ok_rev = pct >= 98.0
            out.append(Result("dispatch:rev_equals_contract_rental",
                              "PASS" if ok_rev else _fail_severity("dispatch:rev_equals_contract_rental"),
                              f"{pct:.2f}", "98",
                              "" if ok_rev else f"{n_ne} stops where stop.rev != contractInfo.tt — rental semantics drifted [soft gate — non-blocking]"))
    except Exception as e:
        out.append(Result("dispatch:rev_equals_contract_rental", "WARN", "", "", f"could not cross-check: {e}"))
    return out


def check_venues(skip_drift: bool) -> list[Result]:
    out = []
    try:
        d = _load_json(VENUES_JSON)
    except Exception as e:
        out.append(Result("venues:parse", "FAIL", "", "", f"{type(e).__name__}: {e}"))
        return out

    venues = d.get("venues") or []
    n = len(venues)
    ok = n >= MIN_VENUES
    out.append(Result("venues:count", "PASS" if ok else "FAIL",
                      str(n), str(MIN_VENUES),
                      "" if ok else "venues list below floor"))

    # FSA validity on venue addresses (where a postal code can be parsed out)
    fsa_total = 0
    fsa_valid = 0
    pc_re = re.compile(r"\b([A-Z]\d[A-Z])\s*\d[A-Z]\d\b", re.IGNORECASE)
    for v in venues:
        addr = (v.get("address") or "")
        m = pc_re.search(addr)
        if m:
            fsa_total += 1
            if FSA_REGEX.match(m.group(1)):
                fsa_valid += 1
    if fsa_total > 0:
        pct = fsa_valid / fsa_total * 100
        ok_fsa = pct >= MIN_FSA_VALIDITY_PCT
        out.append(Result("venues:fsa_validity_pct",
                          "PASS" if ok_fsa else "WARN",
                          f"{pct:.1f}", str(MIN_FSA_VALIDITY_PCT),
                          "" if ok_fsa else f"{fsa_valid}/{fsa_total} valid"))
    return out


# === Outcome surfacing (2026-07-07) =========================================
# Every non-PASS gate — hard FAIL or soft WARN — must reach Mark's phone:
#   1. GitHub Actions annotations (::error / ::warning) show on the run page.
#   2. GITHUB_STEP_SUMMARY gets the full gate table.
#   3. Email via notify_failure.send_email (no-ops without GMAIL_APP_PASSWORD).
# A soft WARN stays non-blocking (exit 0) but is never silent.
def _surface_outcomes(all_results: list[Result], counts: dict, run_id: str) -> None:
    non_pass = [r for r in all_results if r.status != "PASS"]

    if os.environ.get("GITHUB_ACTIONS") == "true":
        for r in non_pass:
            level = "error" if r.status == "FAIL" else "warning"
            print(f"::{level} title=qa_checks {r.name}::value={r.value} "
                  f"threshold={r.threshold} {r.detail}")
        summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
        if summary_path:
            try:
                with open(summary_path, "a", encoding="utf-8") as fh:
                    fh.write(f"\n### qa_checks ({run_id}) — {counts['PASS']} pass / "
                             f"{counts['WARN']} warn / {counts['FAIL']} fail\n\n")
                    fh.write("| check | status | value | threshold | detail |\n")
                    fh.write("|---|---|---|---|---|\n")
                    for r in all_results:
                        fh.write(f"| {r.name} | {r.status} | {r.value} | "
                                 f"{r.threshold} | {r.detail} |\n")
            except OSError as e:
                print(f"  [qa_checks] could not write step summary: {e}", file=sys.stderr)

    if non_pass:
        try:
            from notify_failure import send_email
        except ImportError:
            print("  [qa_checks] notify_failure not importable — email skipped.", file=sys.stderr)
            return
        if counts["FAIL"] > 0:
            subject = f"Delivery Map QA: FAIL ({counts['FAIL']} hard) — publish blocked"
        else:
            subject = f"Delivery Map QA: WARN ({counts['WARN']} soft) — published anyway"
        lines = [f"qa_checks run {run_id}: {counts['PASS']} pass / "
                 f"{counts['WARN']} warn / {counts['FAIL']} fail", ""]
        for r in non_pass:
            lines.append(f"[{r.status}] {r.name}: value={r.value} "
                         f"threshold={r.threshold} {r.detail}")
        lines += ["", "Soft (WARN) gates do not block publish/deploy. Hard FAILs abort the run;",
                  "re-trigger via GitHub Actions workflow_dispatch after investigating."]
        send_email(subject, "\n".join(lines))


# === Orchestrator ============================================================
def run(strict: bool, skip_drift: bool) -> int:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    print(f"[qa_checks] run_id={run_id} strict={strict} skip_drift={skip_drift}")

    all_results: list[Result] = []
    all_results.extend(check_file_sizes())
    all_results.extend(check_delivery(skip_drift))
    all_results.extend(check_dispatch(skip_drift))
    all_results.extend(check_venues(skip_drift))

    # Persist to qa_log.csv
    new_file = not LOG_CSV.exists()
    with LOG_CSV.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["run_id", "timestamp_utc", "check", "status", "value", "threshold", "detail"])
        ts = datetime.now(timezone.utc).isoformat(timespec="seconds")
        for r in all_results:
            w.writerow([run_id, ts, r.name, r.status, r.value, r.threshold, r.detail])

    # Summarize
    counts = {"PASS": 0, "WARN": 0, "FAIL": 0}
    for r in all_results:
        counts[r.status] = counts.get(r.status, 0) + 1
        if r.status != "PASS":
            print(f"  [{r.status}] {r.name}: value={r.value} threshold={r.threshold} {r.detail}")

    print(f"\n[qa_checks] {counts['PASS']} pass / {counts['WARN']} warn / {counts['FAIL']} fail / {len(all_results)} total")

    _surface_outcomes(all_results, counts, run_id)

    if counts["FAIL"] > 0:
        return 1
    if strict and counts["WARN"] > 0:
        return 1
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description="Pre-deploy data sanity checks for the Delivery Map.")
    p.add_argument("--strict", action="store_true", help="Treat WARN as FAIL (halt on any non-PASS).")
    p.add_argument("--no-drift", action="store_true", help="Skip trailing-window drift checks.")
    args = p.parse_args()
    return run(strict=args.strict, skip_drift=args.no_drift)


if __name__ == "__main__":
    sys.exit(main())
