@echo off
setlocal enabledelayedexpansion

:: Set working directory
cd /d "%~dp0"

:START
cls
echo =======================================
echo     Robot Update Script (With Safety)
echo =======================================
echo.

:: 1. Fetch latest branch info from remote
echo Fetching latest branch list from remote...
git fetch origin >nul 2>&1

echo.
echo --- Available Branches ---
git branch -a
echo --------------------------
echo.

:: 2. Branch selection
set /p BRANCH_NAME="Enter branch name to update from (press Enter for 'main'): "
if "!BRANCH_NAME!"=="" set BRANCH_NAME=main

:: 3. Validation: Check if the branch exists on remote
git rev-parse --verify origin/!BRANCH_NAME! >nul 2>&1
if !errorlevel! neq 0 (
    echo.
    echo [ERROR] Branch '!BRANCH_NAME!' does not exist on remote.
    echo Please check the name and try again.
    echo.
    pause
    goto START
)

:: 4. Final Confirmation
echo.
echo Branch '!BRANCH_NAME!' verified. 
echo This will update all files EXCEPT the 'dats' folder.
set /p CONFIRM="Are you sure you want to proceed? (Y/N): "

if /i "!CONFIRM!" neq "Y" (
    echo Update cancelled by user.
    pause
    exit /b
)

:: 5. Execution
echo.
echo Updating local files from origin/!BRANCH_NAME!...
git checkout origin/!BRANCH_NAME! -- . ":(exclude)dats"

:: Sync local pointer
git merge origin/!BRANCH_NAME! -X ours --no-commit >nul 2>&1

echo.
echo =======================================
echo SUCCESS: Update complete for !BRANCH_NAME!
echo 'dats' folder was not modified.
echo =======================================
pause