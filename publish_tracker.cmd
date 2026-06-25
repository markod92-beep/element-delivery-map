@echo off
setlocal
REM ===========================================================
REM  Commit the Daily Truck Tracker changes, then push + deploy.
REM  Self-contained in the repo (no Downloads dependency).
REM ===========================================================
set "REPO=C:\Dev\element-delivery-map"
cd /d "%REPO%"
echo --- Staging tracker + data + generator ---
git add tracker.html index.html driver_hours.json process_payworks.py deploy.cmd publish_tracker.cmd
echo.
echo --- Committing ---
git commit -m "feat: Daily Truck Tracker page + map button (auto-populated from dispatch data)"
echo.
echo --- Push + Cloudflare deploy ---
call "%REPO%\deploy.cmd"
