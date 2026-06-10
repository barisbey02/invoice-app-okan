@echo off
:: ─────────────────────────────────────────────────────────────────────────────
:: sync.bat — Pull latest code from GitHub every 60 seconds.
:: Run this as a Windows Service via NSSM (see README for setup instructions).
:: ─────────────────────────────────────────────────────────────────────────────

:: Change this to the full path of your repo on the office computer
set REPO_DIR=C:\invoice-app-okan

:: How long to wait between pulls (in seconds)
set INTERVAL=60

:: Where to write the log file
set LOG_FILE=%REPO_DIR%\sync.log

:: Maximum log size in bytes before it gets rotated (~500 KB)
set MAX_LOG_SIZE=512000

echo [%date% %time%] sync.bat started. Watching %REPO_DIR% every %INTERVAL%s. >> "%LOG_FILE%"

:loop
    :: ── Check log size and rotate if too large ────────────────────────────
    for %%A in ("%LOG_FILE%") do set LOG_SIZE=%%~zA
    if %LOG_SIZE% GTR %MAX_LOG_SIZE% (
        echo [%date% %time%] Log rotated. >> "%LOG_FILE%.old"
        type "%LOG_FILE%" >> "%LOG_FILE%.old"
        echo [%date% %time%] Log rotated — previous entries moved to sync.log.old. > "%LOG_FILE%"
    )

    :: ── Move into the repo directory ──────────────────────────────────────
    cd /d "%REPO_DIR%"
    if errorlevel 1 (
        echo [%date% %time%] ERROR: Could not cd into %REPO_DIR%. Is the path correct? >> "%LOG_FILE%"
        goto wait
    )

    :: ── Run git pull ──────────────────────────────────────────────────────
    git pull origin main >> "%LOG_FILE%" 2>&1
    set GIT_EXIT=%errorlevel%

    if %GIT_EXIT% equ 0 (
        echo [%date% %time%] git pull OK (exit 0) >> "%LOG_FILE%"
    ) else if %GIT_EXIT% equ 1 (
        :: Exit code 1 from git usually means "already up to date" — not an error
        echo [%date% %time%] git pull exit 1 (likely up to date) >> "%LOG_FILE%"
    ) else (
        :: Anything else is a real problem (no network, auth failure, merge conflict, etc.)
        echo [%date% %time%] WARNING: git pull failed with exit code %GIT_EXIT% >> "%LOG_FILE%"
        echo [%date% %time%] Check network connection and GitHub credentials. >> "%LOG_FILE%"
    )

    :: ── Install any new npm packages that arrived in the pull ─────────────
    :: Uses --prefer-offline so it doesn't block if npm registry is slow
    cd /d "%REPO_DIR%\local-backend"
    C:\node\node-v20.19.0-win-x64\npm.cmd install --prefer-offline >> "%LOG_FILE%" 2>&1

:wait
    :: ── Wait INTERVAL seconds before next pull ────────────────────────────
    timeout /t %INTERVAL% /nobreak > nul
goto loop
