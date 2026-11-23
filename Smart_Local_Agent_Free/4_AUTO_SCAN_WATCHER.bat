@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul 2>&1
title Auto Scanner Folder Watcher - Smart Document Agent ($0 Free)
color 0B

set "PATH=%LOCALAPPDATA%\Programs\nodejs;%ProgramFiles%\nodejs;%APPDATA%\npm;%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"
set "WATCHER=%USERPROFILE%\.gemini\doc_tools\scan_watcher.py"
set "CFG=%USERPROFILE%\.gemini\scan_watcher_config.json"

if not exist "%WATCHER%" (
    if exist "%~dp0doc_tools\scan_watcher.py" (
        set "WATCHER=%~dp0doc_tools\scan_watcher.py"
    ) else (
        echo [ERROR] scan_watcher.py not found! Please run 1_INSTALL_FIRST_TIME.bat first.
        pause
        exit /b 1
    )
)

set "PY_CMD="
python -c "import sys; sys.exit(0 if sys.version_info.major==3 else 1)" >nul 2>&1
if %errorlevel%==0 set "PY_CMD=python"
if "!PY_CMD!"=="" (
    py -3 -c "import sys; sys.exit(0)" >nul 2>&1
    if %errorlevel%==0 set "PY_CMD=py -3"
)
if "!PY_CMD!"=="" (
    echo [ERROR] Python 3 not found. Please run 1_INSTALL_FIRST_TIME.bat first.
    pause
    exit /b 1
)

:: If caller passed --reconfig or config doesn't exist yet, run simple 3-question setup!
set "NEED_SETUP=0"
if not exist "%CFG%" set "NEED_SETUP=1"
if /I "%~1"=="--reconfig" set "NEED_SETUP=1"

if "!NEED_SETUP!"=="1" (
    echo ======================================================================
    echo   AUTO SCANNER WATCHER - FIRST TIME FOLDER LINK (1 Minute)
    echo ======================================================================
    echo.
    echo [Step 1/4] SCANNER INBOX FOLDER (where your scanner saves stamped PDFs/JPGs):
    echo   - Drag and drop your current Scan folder here, OR press ENTER for:
    echo     %USERPROFILE%\Desktop\Smart_Scan\1_Scan_Inbox
    echo.
    set "INBOX="
    set /p "INBOX=Scan Inbox Folder > "
    set "INBOX=!INBOX:"=!"
    if "!INBOX!"=="" set "INBOX=%USERPROFILE%\Desktop\Smart_Scan\1_Scan_Inbox"

    echo.
    echo [Step 2/4] WORD / EXCEL ORIGINALS FOLDER (drop .docx / .xlsx before printing!):
    echo   - Any Word/Excel file saved here gets filed + logged automatically, and auto-matched
    echo     when its stamped copy comes back on the Scanner! Press ENTER for:
    echo     %USERPROFILE%\Desktop\Smart_Scan\2_Drafts_Word_Excel
    echo.
    set "DRAFTS="
    set /p "DRAFTS=Word/Excel Originals Folder > "
    set "DRAFTS=!DRAFTS:"=!"
    if "!DRAFTS!"=="" set "DRAFTS=%USERPROFILE%\Desktop\Smart_Scan\2_Drafts_Word_Excel"

    echo.
    echo [Step 3/4] ORGANIZED ARCHIVE FOLDER (where clean folders + Master Excel Log go):
    echo   - Tip: Pick a folder inside Google Drive Desktop (G:\My Drive\...)
    echo     to auto-upload everything to Google Drive for FREE! Press ENTER for:
    echo     %USERPROFILE%\Desktop\Smart_Scan\3_Organized_Archive
    echo.
    set "ARCH="
    set /p "ARCH=Organized Archive Folder > "
    set "ARCH=!ARCH:"=!"
    if "!ARCH!"=="" set "ARCH=%USERPROFILE%\Desktop\Smart_Scan\3_Organized_Archive"

    echo.
    echo [Step 4/4] HOW DO YOU WANT FOLDERS ORGANIZED?
    echo   [1] By Document Type  ^(Submittals / Petty Cash / Invoices / Letters / Contracts^)
    echo   [2] By Project first, then Document Type  ^(Project A\Submittals, Project A\Invoices...^)
    echo.
    set "LCHOICE=1"
    set /p "LCHOICE=Choose [1 or 2] (default 1) > "
    set "LAYOUT=by_type"
    if "!LCHOICE!"=="2" set "LAYOUT=by_project_then_type"

    echo.
    !PY_CMD! "%WATCHER%" --config "%CFG%" --inbox "!INBOX!" --drafts "!DRAFTS!" --archive "!ARCH!" --layout "!LAYOUT!" --setup-only
    if /I "%~1"=="--reconfig" (
        echo.
        echo [OK] Scanner folders updated successfully!
        pause
        exit /b 0
    )
    echo.
)

echo   Tip: Drop or scan ANY PDF/JPG/PNG into your Scan Inbox folder now!
echo   (Keep this window minimized, or use 5_AUTO_START_WITH_WINDOWS.bat
echo    to run silently in the background every time your PC starts. Ctrl+C to stop.)
echo.
!PY_CMD! "%WATCHER%" --config "%CFG%"
pause
