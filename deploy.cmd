@echo off
setlocal
REM ===========================================================
REM  Element Delivery Map - push + Cloudflare Pages deploy
REM  Self-contained; lives in the repo (no Downloads dependency)
REM  Log is written next to this script: deploy_log.txt
REM ===========================================================
set "REPO=C:\Dev\element-delivery-map"
set "LOG=%REPO%\deploy_log.txt"
REM Account ID is an identifier (not a secret); the API token is read
REM from your CLOUDFLARE_API_TOKEN environment variable (kept private).
set "CLOUDFLARE_ACCOUNT_ID=fdcc6b2ba811829869600a30621be271"

echo ===== PUSH + DEPLOY %DATE% %TIME% =====> "%LOG%"
REM clear a stale lock left by any killed/stuck git process (safe)
if exist "%REPO%\.git\index.lock" del /f "%REPO%\.git\index.lock"
echo --- local state before --->> "%LOG%"
git -C "%REPO%" status -sb>> "%LOG%" 2>&1
echo --- pushing origin main (authorize the GitHub prompt if it pops up) --->> "%LOG%"
git -C "%REPO%" push origin main>> "%LOG%" 2>&1
echo push_exit=%ERRORLEVEL%>> "%LOG%"
echo --- deploying to Cloudflare Pages (wrangler) --->> "%LOG%"
pushd "%REPO%"
call wrangler pages deploy . --project-name=element-delivery-map --branch=main --commit-dirty=true>> "%LOG%" 2>&1
echo deploy_exit=%ERRORLEVEL%>> "%LOG%"
popd
echo --- final state --->> "%LOG%"
git -C "%REPO%" status -sb>> "%LOG%" 2>&1
git -C "%REPO%" log --oneline -1 origin/main>> "%LOG%" 2>&1
echo ===== DONE =====>> "%LOG%"

echo.
type "%LOG%"
echo.
echo Finished. Full log saved to: %LOG%
pause
