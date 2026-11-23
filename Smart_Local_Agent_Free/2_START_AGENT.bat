@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Smart Local Document Agent (100%% Free)
color 0A

set "PATH=%LOCALAPPDATA%\Programs\nodejs;%ProgramFiles%\nodejs;%APPDATA%\npm;%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"

:: Determine target folder:
:: 1) If a folder was dragged & dropped onto this .bat (or passed via Right-Click menu), use it.
:: 2) If run from Desktop or starter kit folder without args, prompt user to drag a folder in.
:: 3) If copied into a real project folder and double-clicked, use that folder directly!
set "TARGET=%~1"

if not "!TARGET!"=="" (
    if not exist "!TARGET!\" (
        :: If user dropped a file instead of a folder, use its parent folder
        for %%F in ("!TARGET!") do set "TARGET=%%~dpF"
    )
) else (
    set "HERE=%~dp0"
    if exist "!HERE!1_INSTALL_FIRST_TIME.bat" (
        set "ASK=1"
    ) else if /I "!HERE!"=="%USERPROFILE%\Desktop\" (
        set "ASK=1"
    ) else if /I "!HERE!"=="%USERPROFILE%\.gemini\" (
        set "ASK=1"
    ) else (
        set "TARGET=!HERE!"
    )
)

if "!ASK!"=="1" (
    echo ======================================================================
    echo   SMART LOCAL DOCUMENT AGENT - CHOOSE PROJECT FOLDER
    echo ======================================================================
    echo   Drag and drop ANY folder into this window and press ENTER:
    echo   (Submittals, Petty Cash, Invoices, Letters, Contracts, or Google Drive)
    echo.
    set /p "TARGET=Folder Path > "
    set "TARGET=!TARGET:"=!"
)

if "!TARGET!"=="" (
    echo No folder selected. Exiting.
    pause
    exit /b 0
)

cd /d "!TARGET!" 2>nul
if %errorlevel% NEQ 0 (
    echo [ERROR] Cannot open folder: !TARGET!
    pause
    exit /b 1
)

cls
echo ======================================================================
echo   SMART LOCAL AGENT ACTIVE IN:
echo   %CD%
echo ======================================================================
echo   READY SLASH COMMANDS (or just write anything in Arabic!):
echo     /fillform    - Fill YOUR real Word/Excel project form + auto-number + Log
echo     /revise      - Make R01/R02 directly from original Word/Excel + scan notes
echo     /writeletter - Draft official Arabic/English letter in Word (.docx)
echo     /replyletter - Read incoming scanned letter + draft Word reply (.docx)
echo     /excelform   - Build ANY Excel form/sheet (IPC, PO, VO, MIR, BOQ...)
echo     /organize    - Smart clean, dedup, rename, sort folders + Excel
echo     /submittals  - Audit Submittals (MAR/MIR/WIR/SD) + Update Log/XLSM
echo     /pettycash   - Sort receipts/photos + Petty Cash Excel (KWD 0.000)
echo     /invoices    - Audit supplier/contractor invoices + Register Excel
echo     /letters     - Archive Incoming/Outgoing letters + Correspondence Log
echo     /contracts   - Organize Contracts, IPCs, VOs, Guarantees + Tracker
echo     /linklog     - Match existing Excel to PDFs + add clickable links
echo ======================================================================
echo.

where gemini >nul 2>&1
if %errorlevel% NEQ 0 (
    echo [ERROR] Gemini CLI is not installed yet.
    echo         Please run 1_INSTALL_FIRST_TIME.bat first!
    pause
    exit /b 1
)

:: Load saved free Google AI Studio key + switch Gemini CLI auth mode + select best free model automatically
set "MODEL_ARGS="
set "PY_CMD="
python -c "import sys" >nul 2>&1
if %errorlevel%==0 set "PY_CMD=python"
if "!PY_CMD!"=="" (
    py -3 -c "import sys" >nul 2>&1
    if !errorlevel!==0 set "PY_CMD=py -3"
)
if exist "%USERPROFILE%\.gemini\ai_key.json" (
    if not "!PY_CMD!"=="" (
        for /f "usebackq tokens=1,2 delims=|" %%A in (`!PY_CMD! -c "import json,pathlib;g=pathlib.Path.home()/'.gemini';d=json.loads((g/'ai_key.json').read_text('utf-8'));k=(d.get('gemini_api_key') or '').strip();m=(d.get('top_model') or '').strip();sf=g/'settings.json';s=json.loads(sf.read_text('utf-8')) if sf.exists() else {};s.setdefault('security',{}).setdefault('auth',{})['selectedType']='gemini-api-key';sf.write_text(json.dumps(s,indent=2),'utf-8') if k else None;print((k or '__NONE__')+'|'+(m or '__NONE__'))" 2^>nul`) do (
            if not "%%A"=="__NONE__" if not "%%A"=="" set "GEMINI_API_KEY=%%A"
            if not "%%B"=="__NONE__" if not "%%B"=="" set "MODEL_ARGS=-m %%B"
        )
    )
)

call gemini %MODEL_ARGS%
pause
