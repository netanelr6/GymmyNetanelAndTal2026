@echo off
setlocal enabledelayedexpansion

:: --- CONFIGURATION ---
set "GITHUB_USER=netanelr6"
set "GITHUB_TOKEN=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"
set "REPO_PATH=github.com/netanelr6/GymmyNetanelAndTal2026.git"
:: ---------------------

cd /d "%~dp0"

:: 1. Sync time from network
for /f "usebackq tokens=*" %%a in (`powershell -command "try { (Invoke-RestMethod -Uri 'http://worldtimeapi.org/api/ip').datetime } catch { Get-Date -Format 'yyyy-MM-ddTHH:mm:ss' }"`) do (
    set "RAW_TIME=%%a"
)
set "FINAL_TIME=!RAW_TIME:~0,10! !RAW_TIME:~11,8!"

:: 2. Set Auth with both Username and Token (The "Secret Sauce")
echo Authenticating as !GITHUB_USER!...
git remote set-url origin "https://!GITHUB_USER!:!GITHUB_TOKEN!@!REPO_PATH!"

:: 3. Stage changes
echo Staging changes...
git add dats
git add .

:: 4. Commit
git commit -m "Experimental data sync, from !FINAL_TIME!"

:: 5. Push with forced override
echo Pushing to GitHub...
:: Using -c credential.helper= to ensure it doesn't ask for a password
git -c credential.helper= push origin main --force

:: Check if the push was successful
if !errorlevel! neq 0 (
    echo.
    echo [ERROR] Push failed. Please check if your Token is still valid.
) else (
    echo.
    echo ===========================================
    echo   SUCCESS: Data synced to GitHub!
    echo   Time: !FINAL_TIME!
    echo ===========================================
)

:: 6. Cleanup security (Remove token from local git config)
git remote set-url origin "https://!REPO_PATH!"

pause