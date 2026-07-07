"""smoke_test.py - end-to-end UI smoke test for the Element Delivery Map.

Runs AFTER `wrangler pages deploy` in daily_refresh.ps1. Loads the live site
in headless Chromium and verifies that the three recurring UI failure modes
have not regressed:

    1. Contract info panel populates when a contract is opened
    2. Timeline filter chips render (date range, day-of-week, contract status)
    3. Layer toggles still function and add/remove map features

Plus baseline checks: page loads with no JS console errors, all three JSON
data files are reachable, all expected layer buttons exist, and a screenshot
is saved per run for visual regression review.

On ANY failure:
    - Writes refresh-v2\\LAST_RUN_FAILED.flag with one-line summary
    - Emails markod92@gmail.com (if GMAIL_APP_PASSWORD env var is set)
    - Exits with code 1 (causes daily_refresh.ps1 to halt)

On success: exits 0 silently, removes any prior LAST_RUN_FAILED.flag.

Always:
    - Appends one row per check to refresh-v2\\health_log\\smoke_log.csv
    - Saves screenshot to refresh-v2\\health_log\\screenshots\\YYYY-MM-DD.png

Install (one-time):
    pip install playwright requests --break-system-packages
    python -m playwright install chromium

Run:
    python smoke_test.py
    python smoke_test.py --url https://staging.example.com   (override target)
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import smtplib
import ssl
import sys
import traceback
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

try:
    import requests
except ImportError:
    print("ERROR: requests not installed. Run: pip install requests --break-system-packages", file=sys.stderr)
    sys.exit(2)

try:
    from playwright.sync_api import sync_playwright, Error as PlaywrightError, TimeoutError as PlaywrightTimeout
except ImportError:
    print("ERROR: playwright not installed. Run:", file=sys.stderr)
    print("  pip install playwright --break-system-packages", file=sys.stderr)
    print("  python -m playwright install chromium", file=sys.stderr)
    sys.exit(2)


# === Config ==================================================================
DEFAULT_URL = "https://map.elementdeliverymap.ca"
SCRIPT_DIR = Path(__file__).resolve().parent
LOG_DIR = SCRIPT_DIR / "health_log"
SCREENSHOT_DIR = LOG_DIR / "screenshots"
LOG_CSV = LOG_DIR / "smoke_log.csv"
FAIL_FLAG = SCRIPT_DIR / "LAST_RUN_FAILED.flag"

PAGE_LOAD_TIMEOUT_MS = 60_000   # 60s for full page + JSON fetches
DATA_READY_TIMEOUT_MS = 45_000  # contractInfo populated + map painted
ACTION_SETTLE_MS = 800          # short wait after toggling a layer

# Pages-Function password gate (refresh-v2/functions/_middleware.js).
# The site is protected by an HMAC-signed cookie set by POST /__auth/login.
# We authenticate once at startup and ride the resulting `map_auth` cookie on
# every subsequent request (both `requests` and Playwright).
MAP_PASSWORD = os.environ.get("MAP_PASSWORD")
AUTH_LOGIN_PATH = "/__auth/login"
AUTH_COOKIE_NAME = "map_auth"


def _authenticate(url: str) -> tuple["requests.Session", str | None]:
    """POST the password to /__auth/login, return (session, cookie_value).

    Returns (Session(), None) when MAP_PASSWORD is unset — caller should warn
    and proceed (network checks will then fail loudly with 401, which is
    the correct behavior — better than silently passing).
    """
    s = requests.Session()
    if not MAP_PASSWORD:
        return s, None
    r = s.post(
        f"{url.rstrip('/')}{AUTH_LOGIN_PATH}",
        data={"password": MAP_PASSWORD},
        allow_redirects=False,
        timeout=30,
    )
    if r.status_code == 401:
        raise RuntimeError("Login POST returned 401 — MAP_PASSWORD is wrong.")
    if r.status_code != 303:
        raise RuntimeError(f"Login POST returned {r.status_code} (expected 303). Check the gate is up.")
    cookie_val = s.cookies.get(AUTH_COOKIE_NAME)
    if not cookie_val:
        raise RuntimeError(f"Login succeeded but '{AUTH_COOKIE_NAME}' cookie was not set on the response.")
    return s, cookie_val

# Layer buttons exposed on every load
EXPECTED_LAYER_BUTTONS = ["btn-rev", "btn-dispatch"]

# Filter chips that must render after delivery_data.json loads
EXPECTED_FILTER_IDS = [
    "tl-preset-chips",   # quick-jump preset row
    "tl-from-input",     # date range start
    "tl-to-input",       # date range end
    "cf-all", "cf-closed", "cf-open",                                      # contract status
    "dow-mon", "dow-tue", "dow-wed", "dow-thu", "dow-fri", "dow-sat", "dow-sun",  # day-of-week
    "co-all", "co-gta", "co-gta-tents", "co-east", "co-west",              # company
]


# === Logging =================================================================
def _ensure_dirs() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)


def _log_check(run_id: str, check: str, status: str, detail: str = "") -> None:
    """Append one CSV row per check. Header is written on first creation."""
    _ensure_dirs()
    new_file = not LOG_CSV.exists()
    with LOG_CSV.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["run_id", "timestamp_utc", "check", "status", "detail"])
        w.writerow([run_id, datetime.now(timezone.utc).isoformat(timespec="seconds"), check, status, detail])


# === Notification ============================================================
def _write_flag(summary: str) -> None:
    FAIL_FLAG.write_text(
        f"Smoke test FAILED at {datetime.now().isoformat(timespec='seconds')}\n\n{summary}\n",
        encoding="utf-8",
    )


def _clear_flag() -> None:
    if FAIL_FLAG.exists():
        FAIL_FLAG.unlink()


def _send_email(subject: str, body: str) -> None:
    """Email via Gmail SMTP. Silently no-ops if GMAIL_APP_PASSWORD not set."""
    pwd = os.environ.get("GMAIL_APP_PASSWORD")
    if not pwd:
        return
    sender = os.environ.get("GMAIL_SENDER", "markod92@gmail.com")
    recipient = os.environ.get("QA_NOTIFY_EMAIL", "markod92@gmail.com")
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
            s.login(sender, pwd)
            s.send_message(msg)
    except Exception as e:
        # Don't let email failure mask the actual smoke failure
        print(f"  [warn] email send failed: {e}", file=sys.stderr)


# === Helpers =================================================================
def _pick_test_contracts(url: str, session: "requests.Session", n: int = 3) -> list[str]:
    """Pull top-N contract numbers by Total Value from live delivery_data.json.

    Falls back to first-N keys if 'tt' (total value) field is absent on most
    entries. Self-healing: contract list refreshes automatically each run.
    """
    r = session.get(
        f"{url.rstrip('/')}/delivery_data.json?cb={int(datetime.now().timestamp())}",
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    ci = data.get("contractInfo") or {}
    if not ci:
        raise RuntimeError("delivery_data.json has empty contractInfo — cannot pick test contracts")
    # Sort by total value desc; stable on contracts that have it
    ranked = sorted(ci.items(), key=lambda kv: float((kv[1] or {}).get("tt") or 0), reverse=True)
    return [k for k, _ in ranked[:n]]


# === Checks ==================================================================
class CheckResult:
    __slots__ = ("name", "passed", "detail")
    def __init__(self, name: str, passed: bool, detail: str = ""):
        self.name = name
        self.passed = passed
        self.detail = detail


def check_data_files_reachable(url: str, session: "requests.Session") -> list[CheckResult]:
    """Each JSON the page fetches must return 200 + valid JSON + expected key."""
    results = []
    targets = [
        ("delivery_data.json", "contractInfo"),
        ("dispatch_data.json", "dates"),
        ("venues_data.json", "venues"),
    ]
    for fname, expected_key in targets:
        try:
            r = session.get(
                f"{url.rstrip('/')}/{fname}?cb={int(datetime.now().timestamp())}",
                timeout=30,
            )
            if r.status_code != 200:
                results.append(CheckResult(f"data_file:{fname}", False, f"HTTP {r.status_code}"))
                continue
            j = r.json()
            if expected_key not in j:
                results.append(CheckResult(f"data_file:{fname}", False, f"missing key '{expected_key}'"))
                continue
            size_kb = len(r.content) // 1024
            results.append(CheckResult(f"data_file:{fname}", True, f"{size_kb}KB ok"))
        except Exception as e:
            results.append(CheckResult(f"data_file:{fname}", False, f"{type(e).__name__}: {e}"))
    return results


def check_page_loads_clean(page, url: str) -> list[CheckResult]:
    """Page must reach 'load' and produce no JS console errors during render."""
    console_errors: list[str] = []
    page.on("console", lambda msg: console_errors.append(msg.text) if msg.type == "error" else None)
    page.on("pageerror", lambda exc: console_errors.append(f"pageerror: {exc}"))

    results = []
    try:
        page.goto(url, wait_until="load", timeout=PAGE_LOAD_TIMEOUT_MS)
        results.append(CheckResult("page_load", True, "load event fired"))
    except PlaywrightTimeout as e:
        results.append(CheckResult("page_load", False, f"timeout: {e}"))
        return results

    # Wait for delivery_data.json to finish populating. We can't see top-level
    # `const contractInfo` from Playwright evaluate (those declarations aren't
    # on `window`), so instead we wait for a DOM signal that proves the render
    # pass after the fetch completed: layer toggle buttons only get injected
    # once data has loaded.
    try:
        page.wait_for_function(
            "() => document.querySelectorAll('.toggle-btn').length >= 6",
            timeout=DATA_READY_TIMEOUT_MS,
        )
        btn_count = page.evaluate("() => document.querySelectorAll('.toggle-btn').length")
        results.append(CheckResult("data_loaded", True, f"{btn_count} toggle buttons rendered"))
    except PlaywrightTimeout:
        results.append(CheckResult("data_loaded", False, "layer toggle buttons never rendered within 45s"))

    # Filter to genuinely-bad errors. The page logs a benign warning when
    # dispatch_data.json is empty during back-fill — don't fail on that.
    real_errors = [e for e in console_errors if "dispatch_data.json unavailable" not in e]
    if real_errors:
        results.append(CheckResult("no_console_errors", False, f"{len(real_errors)} JS errors: {real_errors[0][:120]}"))
    else:
        results.append(CheckResult("no_console_errors", True, "clean"))
    return results


def check_layer_buttons(page) -> list[CheckResult]:
    """Every expected layer toggle button must be present in the DOM."""
    results = []
    for btn_id in EXPECTED_LAYER_BUTTONS:
        exists = page.evaluate(f"() => !!document.getElementById('{btn_id}')")
        results.append(CheckResult(f"layer_btn:{btn_id}", bool(exists), "" if exists else "missing from DOM"))
    return results


def check_filter_chips(page) -> list[CheckResult]:
    """Timeline + contract + DOW + company filter chips must all render."""
    results = []
    for fid in EXPECTED_FILTER_IDS:
        exists = page.evaluate(f"() => !!document.getElementById('{fid}')")
        results.append(CheckResult(f"filter:{fid}", bool(exists), "" if exists else "missing from DOM"))
    # Preset chip row must be non-empty (catches "filters didn't render" failure)
    preset_count = page.evaluate(
        "() => { const el=document.getElementById('tl-preset-chips'); return el ? el.children.length : 0; }"
    )
    results.append(CheckResult("filter:tl-preset-chips_populated", preset_count > 0, f"{preset_count} chips"))
    return results


def check_layer_toggles_function(page) -> list[CheckResult]:
    """Toggle each layer on/off via window.toggleLayer; assert button state flips.

    We don't hardcode class names — different layers use different active-class
    patterns (`active`, `active-new`, `active-rev`, etc.) and those names drift
    as the CSS evolves. Instead we capture the button's full className before
    and after the toggle and assert it changed. Then restore state so later
    checks see a clean page.
    """
    results = []
    layer_keys = ["rev", "dispatch"]
    for key in layer_keys:
        try:
            result = page.evaluate(f"""() => {{
                const btn = document.getElementById('btn-{key}');
                if (!btn) return {{ ok: false, reason: 'button missing' }};
                const before = btn.className;
                window.toggleLayer('{key}');
                const after = btn.className;
                window.toggleLayer('{key}');  /* restore */
                return {{ ok: before !== after, before, after }};
            }}""")
            page.wait_for_timeout(ACTION_SETTLE_MS)
            if result.get("ok"):
                results.append(CheckResult(
                    f"layer_toggle:{key}", True,
                    f"className changed ('{result.get('before','')}' -> '{result.get('after','')}')",
                ))
            else:
                results.append(CheckResult(
                    f"layer_toggle:{key}", False,
                    f"className unchanged after toggle: '{result.get('before','')}' (reason: {result.get('reason','no diff')})",
                ))
        except PlaywrightError as e:
            results.append(CheckResult(f"layer_toggle:{key}", False, f"JS error: {e}"))
    return results


def check_contract_detail_populates(page, contract_nums: list[str]) -> list[CheckResult]:
    """Inject a test detail container, call window.showContractDetail for each
    sample contract, assert the container fills with a contract-detail-card.

    This directly exercises the failure path Mark sees most: contract info
    not rendering after a refresh changes the data shape.
    """
    results = []
    for cn in contract_nums:
        slot_id = f"smoke-cd-{cn}"
        try:
            html = page.evaluate(f"""(cn) => {{
                /* Make a host element if not already there */
                let host = document.getElementById('{slot_id}');
                if (!host) {{
                    host = document.createElement('div');
                    host.id = '{slot_id}';
                    host.className = 'cd-inline';
                    document.body.appendChild(host);
                }}
                window.showContractDetail(String(cn), '{slot_id}');
                return host.innerHTML || '';
            }}""", cn)
            ok = ("contract-detail-card" in html) and ("Contract #" in html)
            # 'No detail available' indicates contractInfo lookup miss — partial fail
            no_detail = "No detail available" in html
            if ok and not no_detail:
                results.append(CheckResult(f"contract_detail:{cn}", True, f"{len(html)} bytes rendered"))
            elif no_detail:
                results.append(CheckResult(f"contract_detail:{cn}", False, "contractInfo lookup miss"))
            else:
                results.append(CheckResult(f"contract_detail:{cn}", False, f"unexpected HTML: {html[:120]}"))
        except PlaywrightError as e:
            results.append(CheckResult(f"contract_detail:{cn}", False, f"JS error: {e}"))
    return results


def check_map_initialized(page) -> list[CheckResult]:
    """Leaflet map must have bootstrapped — detected purely via DOM signals.

    We intentionally do NOT reference a `map` global: the bare identifier `map`
    in the page's evaluate context resolves to the `<div id="map">` DOM element
    via browser auto-named-access, so `map.getZoom()` throws even when Leaflet
    initialized correctly. Instead we check the signals Leaflet writes to the
    DOM once it paints: the container gets `.leaflet-container`, panes are
    created, and tile images load.
    """
    try:
        info = page.evaluate("""() => {
            const el = document.querySelector('#map');
            if (!el) return { ok: false, reason: 'no #map element' };
            const isContainer = el.classList.contains('leaflet-container');
            const panes = document.querySelectorAll('#map .leaflet-pane').length;
            const tiles = document.querySelectorAll('#map .leaflet-tile-loaded').length;
            return {
                ok: isContainer && panes > 0 && tiles > 0,
                isContainer, panes, tiles,
            };
        }""")
        if info.get("ok"):
            return [CheckResult(
                "map_initialized", True,
                f"leaflet-container={info['isContainer']} panes={info['panes']} tiles={info['tiles']}",
            )]
        return [CheckResult("map_initialized", False, json.dumps(info))]
    except PlaywrightError as e:
        return [CheckResult("map_initialized", False, f"JS error: {e}")]


def take_screenshot(page, run_id: str) -> Path:
    _ensure_dirs()
    out = SCREENSHOT_DIR / f"{run_id}.png"
    page.screenshot(path=str(out), full_page=False)
    return out


# === Orchestrator ============================================================
def run(url: str) -> int:
    run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    print(f"[smoke_test] run_id={run_id} target={url}")

    # Auth: log in to the password gate, get a session + cookie usable for
    # both `requests` calls and the Playwright browser context.
    auth_cookie: str | None = None
    if not MAP_PASSWORD:
        print("  [warn] MAP_PASSWORD env var not set — site is behind a password gate.")
        print("         Set it via [Environment]::SetEnvironmentVariable('MAP_PASSWORD','<pwd>','User')")
        session = requests.Session()
    else:
        try:
            print("  [auth] Logging in to /__auth/login...")
            session, auth_cookie = _authenticate(url)
            print(f"  [auth] OK — got {AUTH_COOKIE_NAME} cookie ({len(auth_cookie)} chars)")
        except Exception as e:
            print(f"  [auth] FAILED: {type(e).__name__}: {e}")
            session = requests.Session()

    all_results: list[CheckResult] = []
    all_results.append(CheckResult(
        "auth_login",
        bool(auth_cookie),
        f"{AUTH_COOKIE_NAME} cookie acquired" if auth_cookie else "no cookie — downstream checks will likely 401",
    ))

    # Phase 1: pre-browser network checks (fast, independent of Playwright)
    print("  [1/5] Checking JSON data files reachable...")
    all_results.extend(check_data_files_reachable(url, session))

    # Phase 2: pick test contracts dynamically
    try:
        print("  [2/5] Picking test contracts from delivery_data.json...")
        test_contracts = _pick_test_contracts(url, session, n=3)
        print(f"        -> {test_contracts}")
        all_results.append(CheckResult("test_contracts_picked", True, f"{test_contracts}"))
    except Exception as e:
        all_results.append(CheckResult("test_contracts_picked", False, f"{type(e).__name__}: {e}"))
        test_contracts = []

    # Phase 3-5: browser-driven checks
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1440, "height": 900})
            # Inject the auth cookie so every browser request rides past the
            # password gate — root HTML, JSON fetches, sub-resources alike.
            if auth_cookie:
                from urllib.parse import urlparse
                domain = urlparse(url).hostname or ""
                context.add_cookies([{
                    "name": AUTH_COOKIE_NAME,
                    "value": auth_cookie,
                    "domain": domain,
                    "path": "/",
                    "secure": True,
                    "httpOnly": True,
                    "sameSite": "Lax",
                }])
            page = context.new_page()

            print("  [3/5] Loading page + waiting for data...")
            all_results.extend(check_page_loads_clean(page, url))

            print("  [4/5] Checking layer buttons + filter chips + map...")
            all_results.extend(check_layer_buttons(page))
            all_results.extend(check_filter_chips(page))
            all_results.extend(check_map_initialized(page))
            all_results.extend(check_layer_toggles_function(page))

            if test_contracts:
                print(f"  [5/5] Verifying contract detail renders for {len(test_contracts)} contracts...")
                all_results.extend(check_contract_detail_populates(page, test_contracts))

            try:
                shot = take_screenshot(page, run_id)
                all_results.append(CheckResult("screenshot", True, str(shot.name)))
            except Exception as e:
                all_results.append(CheckResult("screenshot", False, f"{type(e).__name__}: {e}"))

            browser.close()
    except Exception:
        all_results.append(CheckResult("playwright_runtime", False, traceback.format_exc().splitlines()[-1]))

    # === Persist + summarize ===
    for r in all_results:
        _log_check(run_id, r.name, "PASS" if r.passed else "FAIL", r.detail)

    failures = [r for r in all_results if not r.passed]
    total = len(all_results)
    passed = total - len(failures)
    print(f"\n[smoke_test] {passed}/{total} checks passed")

    if failures:
        summary_lines = [f"{r.name}: {r.detail}" for r in failures]
        summary = "\n".join(summary_lines)
        print("FAILURES:")
        for line in summary_lines:
            print(f"  - {line}")
        _write_flag(summary)
        _send_email(
            subject=f"[Delivery Map] Smoke test FAILED — {len(failures)}/{total} checks ({run_id})",
            body=f"Live site: {url}\nRun: {run_id}\n\nFailing checks:\n{summary}\n\nFull log: {LOG_CSV}\nScreenshot: {SCREENSHOT_DIR / (run_id + '.png')}",
        )
        return 1

    _clear_flag()
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description="UI smoke test for the live Delivery Map.")
    p.add_argument("--url", default=DEFAULT_URL, help=f"Site root (default: {DEFAULT_URL})")
    args = p.parse_args()
    return run(args.url)


if __name__ == "__main__":
    sys.exit(main())
