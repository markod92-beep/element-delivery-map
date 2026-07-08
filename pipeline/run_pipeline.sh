#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/.."   # repo root
P=pipeline
SRC=_src

# CI layout wiring: outputs land at the repo root, scripts live in pipeline/.
export QA_DATA_DIR=.                        # qa_checks.py reads the JSONs from here
export OSRM_CACHE_PATH=$P/osrm_cache.json   # persisted between runs via actions/cache

git config user.name  "delivery-map-bot"
git config user.email "delivery-map-bot@users.noreply.github.com"

echo "[1/9] process_cmmc";        python $P/process_cmmc.py "$SRC/CMMC_UniqueEvents.xlsx" delivery_data.json --ref delivery_data.json
                                  python $P/validate_pipeline.py delivery delivery_data.json
echo "[2/9] process_dispatch";    python $P/process_dispatch.py "$SRC/CMMC_Dispatch_Routes.xlsx" dispatch_data.json --fetch-geometry
                                  python $P/validate_pipeline.py dispatch dispatch_data.json
echo "[3/9] process_venues";      python $P/process_venues.py delivery_data.json dispatch_data.json venues_data.json
                                  python $P/validate_pipeline.py venues venues_data.json
echo "[4/9] annotate_unassigned"; python $P/annotate_unassigned.py delivery_data.json dispatch_data.json
echo "[5/9] qa_checks";           python $P/qa_checks.py

echo "[6/9] process_payworks (non-fatal)"
PUB_DH=""
if [ -f "$SRC/Element Event Solutions - Timesheet.xlsx" ]; then
  if python $P/process_payworks.py "$SRC/Element Event Solutions - Timesheet.xlsx" driver_hours.json; then
    PUB_DH="--driver-hours-json driver_hours.json"
  else echo "  WARN process_payworks failed - keeping last good"; fi
else echo "  WARN Payworks timesheet missing - skipping"; fi

echo "[7/9] process_truckload (non-fatal)"
PUB_CL=""
if python $P/process_truckload.py "$SRC/CMMC_UniqueEvents.xlsx" dispatch_data.json contract_load.json; then
  PUB_CL="--contract-load-json contract_load.json"
else echo "  WARN process_truckload failed - keeping last good"; fi

echo "[8/9] publish to GitHub (NO --index-html)"
python $P/publish.py --repo-dir . \
  --data-json delivery_data.json --dispatch-json dispatch_data.json --venues-json venues_data.json \
  $PUB_DH $PUB_CL

echo "[9/9] deploy to Cloudflare + smoke test"
# Cloudflare Pages uploads every file in the deploy dir and rejects any file >25 MiB.
# `.assetsignore` is NOT honored by `wrangler pages deploy`, so keep the large non-web
# inputs out of the tree at deploy time (the laptop's repo never contains them):
# remove the source xlsx (not needed post-processing) and stash the OSRM cache,
# restoring it afterward so the workflow's actions/cache save step still finds it.
rm -rf "$SRC"
CACHE_STASH=""
if [ -f "$OSRM_CACHE_PATH" ]; then CACHE_STASH="$(mktemp)"; mv "$OSRM_CACHE_PATH" "$CACHE_STASH"; fi
npx wrangler pages deploy . --project-name=element-delivery-map --branch=main --commit-dirty=true
if [ -n "$CACHE_STASH" ]; then mv "$CACHE_STASH" "$OSRM_CACHE_PATH"; fi
sleep 20
python $P/smoke_test.py
echo "--- done ---"
