@echo off
setlocal enabledelayedexpansion

:: --- CONFIGURATION ---
:: Paste your token inside the quotes below
set "GITHUB_TOKEN=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"
set "REPO_URL=github.com/netanelr6/GymmyNetanelAndTal2026.git"
:: ---------------------

cd /d "%~dp0"

:: 1. Sync time
for /f "usebackq tokens=*" %%a in (`powershell -command "try { (Invoke-RestMethod -Uri 'http://worldtimeapi.org/api/ip').datetime } catch { Get-Date -Format 'yyyy-MM-ddTHH:mm:ss' }"`) do (
    set "RAW_TIME=%%a"
)
set "FINAL_TIME=!RAW_TIME:~0,10! !RAW_TIME:~11,8!"

:: 2. Forced Authentication Update
echo Using Token to bypass Windows Credential Manager...
git remote set-url origin "https://!GITHUB_TOKEN!@!REPO_URL!"

:: 3. Stage changes
git add dats
git add .

:: 4. Commit
git commit -m "Experimental data sync, from !FINAL_TIME!"

:: 5. Push WITH CREDENTIAL OVERRIDE
:: The "-c credential.helper=" part tells Git to ignore stored passwords
echo Pushing to GitHub...
git -c credential.helper= push origin main --force

:: 6. Cleanup
git remote set-url origin "https://!REPO_URL!"

echo.
echo ===========================================
echo   SUCCESS: Sync complete.
echo ===========================================
pause