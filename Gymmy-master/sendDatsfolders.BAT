@echo off
setlocal enabledelayedexpansion

:: --- CONFIGURATION (PASTE YOUR TOKEN HERE) ---
set "GITHUB_TOKEN=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"
set "REPO_URL=github.com/netanelr6/GymmyNetanelAndTal2026.git"
:: ---------------------------------------------

cd /d "%~dp0"

:: 1. Sync time from network
for /f "usebackq tokens=*" %%a in (`powershell -command "try { (Invoke-RestMethod -Uri 'http://worldtimeapi.org/api/ip').datetime } catch { Get-Date -Format 'yyyy-MM-ddTHH:mm:ss' }"`) do (
    set "RAW_TIME=%%a"
)
set "FINAL_TIME=!RAW_TIME:~0,10! !RAW_TIME:~11,8!"

:: 2. Set temporary auth (This is the fix for the 403 error)
echo Updating remote URL with token...
git remote set-url origin "https://!GITHUB_TOKEN!@!REPO_URL!"

echo.
echo --- STEP 1: Adding files from dats folder ---
:: Adding the folder and any changes
git add dats
git add .

echo.
echo --- STEP 2: Committing ---
git commit -m "Experimental data sync, from !FINAL_TIME!"

echo.
echo --- STEP 3: Pushing to GitHub (Forcing to resolve divergence) ---
:: We use --force because the robot and cloud branches have diverged
git push origin main --force

:: 3. Clean up (Security: remove token from remote URL)
git remote set-url origin "https://!REPO_URL!"

echo.
echo ===========================================
echo   SUCCESS: Data sent to GitHub.
echo   Time: !FINAL_TIME!
echo ===========================================
pause