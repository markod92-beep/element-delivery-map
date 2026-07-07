"""
publish.py — deploy a fresh delivery_data.json to Cloudflare Pages via GitHub.

Cloudflare Pages is wired to auto-build on every push to the connected GitHub
repo. So "publishing" is just: copy the JSON into the repo, commit, push. The
HTML shell (index.html) only needs to be pushed once — after that, every refresh
is a JSON-only commit.

Typical usage (daily refresh):
    python publish.py \\
        --data-json "C:/.../Delivery Map Refresh/delivery_data.json" \\
        --repo-dir  "C:/.../delivery-map-repo"

First-time setup (push both the shell AND the data):
    python publish.py \\
        --data-json  "C:/.../delivery_data.json" \\
        --index-html "C:/.../index.html" \\
        --repo-dir   "C:/.../delivery-map-repo" \\
        --message    "initial deploy: decoupled data"

Prereqs:
  - You have a local git clone of the Cloudflare-Pages-connected repo
  - `git` is on PATH and you can push to that remote without a password prompt
    (SSH key, stored HTTPS credential, or `gh auth login`)
"""
import argparse
import datetime as dt
import shutil
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str], cwd: Path) -> subprocess.CompletedProcess:
    """Run a command, streaming output, raise on nonzero exit."""
    print(f"  $ {' '.join(cmd)}")
    return subprocess.run(cmd, cwd=cwd, check=True, text=True)


def ensure_git_repo(repo_dir: Path) -> None:
    if not (repo_dir / ".git").exists():
        raise SystemExit(
            f"ERROR: {repo_dir} is not a git repository.\n"
            "  Run `git clone <your-repo-url> <dir>` first, or point --repo-dir "
            "at the existing local clone of the Cloudflare-Pages-connected repo."
        )


def copy_if_different(src: Path, dst: Path) -> bool:
    """Copy src to dst only if content differs. Returns True if a copy happened."""
    if not src.exists():
        raise SystemExit(f"ERROR: source file missing: {src}")
    # CI layout: the pipeline writes outputs directly into the repo working
    # tree, so src IS dst. Nothing to copy, but the file may still differ from
    # HEAD — report it as changed and let the `git diff --cached` guard below
    # decide whether a commit is actually needed.
    if dst.exists() and src.resolve() == dst.resolve():
        print(f"  (in place) {dst.name}")
        return True
    if dst.exists() and src.read_bytes() == dst.read_bytes():
        print(f"  (unchanged) {dst.name}")
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    print(f"  copied {src} -> {dst}")
    return True


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Publish delivery_data.json (and optionally index.html) to the "
        "Cloudflare-Pages-connected GitHub repo via git push."
    )
    ap.add_argument("--data-json", required=True, type=Path,
                    help="Path to the freshly built delivery_data.json")
    ap.add_argument("--repo-dir", required=True, type=Path,
                    help="Path to the local git clone of the Pages repo")
    ap.add_argument("--index-html", type=Path, default=None,
                    help="Optional: also publish a new index.html (rarely needed — "
                         "only after editing the shell)")
    ap.add_argument("--dispatch-json", type=Path, default=None,
                    help="Optional: also publish dispatch_data.json for the Routes layer")
    ap.add_argument("--venues-json", type=Path, default=None,
                    help="Optional: also publish venues_data.json for the Key Venues layer")
    ap.add_argument("--driver-hours-json", type=Path, default=None,
                    help="Optional: also publish driver_hours.json for the Daily Truck Tracker")
    ap.add_argument("--contract-load-json", type=Path, default=None,
                    help="Optional: also publish contract_load.json (truck-fill cubes + per-BU capacity)")
    ap.add_argument("--branch", default="main",
                    help="Branch to push (default: main)")
    ap.add_argument("--message", default=None,
                    help="Commit message (default: 'data: refresh YYYY-MM-DD')")
    ap.add_argument("--dry-run", action="store_true",
                    help="Copy files and stage them but skip commit/push")
    args = ap.parse_args()

    ensure_git_repo(args.repo_dir)

    # Stage file(s)
    print("Staging files...")
    changed = False
    changed |= copy_if_different(args.data_json, args.repo_dir / "delivery_data.json")
    if args.index_html:
        changed |= copy_if_different(args.index_html, args.repo_dir / "index.html")
    if args.dispatch_json:
        changed |= copy_if_different(args.dispatch_json, args.repo_dir / "dispatch_data.json")
    if args.venues_json:
        changed |= copy_if_different(args.venues_json, args.repo_dir / "venues_data.json")
    if args.driver_hours_json:
        changed |= copy_if_different(args.driver_hours_json, args.repo_dir / "driver_hours.json")
    if args.contract_load_json:
        changed |= copy_if_different(args.contract_load_json, args.repo_dir / "contract_load.json")

    if not changed:
        print("Nothing to publish — files are identical to the current commit. Exiting.")
        return 0

    # Commit + push
    today = dt.date.today().isoformat()
    message = args.message or f"data: refresh {today}"

    print("Committing...")
    run(["git", "add", "delivery_data.json"], cwd=args.repo_dir)
    if args.index_html:
        run(["git", "add", "index.html"], cwd=args.repo_dir)
    if args.dispatch_json:
        run(["git", "add", "dispatch_data.json"], cwd=args.repo_dir)
    if args.venues_json:
        run(["git", "add", "venues_data.json"], cwd=args.repo_dir)
    if args.driver_hours_json:
        run(["git", "add", "driver_hours.json"], cwd=args.repo_dir)
    if args.contract_load_json:
        run(["git", "add", "contract_load.json"], cwd=args.repo_dir)

    # If nothing is staged (e.g., file content same but mtime differed), exit clean
    staged = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        cwd=args.repo_dir,
    ).returncode
    if staged == 0:
        print("Nothing staged — skipping commit.")
        return 0

    run(["git", "commit", "-m", message], cwd=args.repo_dir)

    if args.dry_run:
        print("DRY RUN — skipping push. Run without --dry-run to deploy.")
        return 0

    print(f"Pushing to origin/{args.branch}...")
    run(["git", "push", "origin", args.branch], cwd=args.repo_dir)

    print()
    print("DONE. Cloudflare Pages will auto-build from the new commit.")
    print("Typical deploy time: 30-90 seconds. Check the Pages dashboard for status.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
