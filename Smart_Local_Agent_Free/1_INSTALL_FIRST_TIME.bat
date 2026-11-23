@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Smart Local Document Agent - First Time Free Setup ($0)
color 0B

echo ======================================================================
echo   SMART LOCAL DOCUMENT AGENT - 100%% FREE SETUP (ZERO COST)
echo   Multi-Project Visual Hub + Auto Scanner + WhatsApp + Word/Excel
echo   Tip: Do NOT click inside this window while installing.
echo ======================================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "GEMINI_HOME=%USERPROFILE%\.gemini"
set "LOCAL_NODE=%LOCALAPPDATA%\Programs\nodejs"
set "LOCAL_PY=%LOCALAPPDATA%\Programs\Python\Python312"
set "PATH=%LOCAL_NODE%;%ProgramFiles%\nodejs;%APPDATA%\npm;%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts;%LOCAL_PY%;%LOCAL_PY%\Scripts;%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"

:: ------------------------------------------------------------------
:: [1/5] & [2/5] Verify Node.js 20+ and Python 3 (fast curl download if missing)
:: ------------------------------------------------------------------
echo [1/5] Checking Node.js 20+ ...
set "NEED_REQ=0"
set "NODE_OK=0"
where node >nul 2>&1
if %errorlevel%==0 (
    for /f "tokens=1 delims=." %%v in ('node -v 2^>nul') do (
        set "NVER=%%v"
        set "NVER=!NVER:v=!"
        if !NVER! GEQ 20 set "NODE_OK=1"
    )
)
if "!NODE_OK!"=="0" set "NEED_REQ=1"

set "PY_CMD="
python -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
if %errorlevel%==0 set "PY_CMD=python"
if "!PY_CMD!"=="" (
    py -3 -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
    if !errorlevel!==0 set "PY_CMD=py -3"
)
if "!PY_CMD!"=="" set "NEED_REQ=1"

if "!NEED_REQ!"=="1" (
    if not "%~1"=="--after-dl" (
        if exist "%SCRIPT_DIR%00_DOWNLOAD_REQUIREMENTS.bat" (
            echo       Missing Node.js or Python - running fast downloader now...
            echo.
            call "%SCRIPT_DIR%00_DOWNLOAD_REQUIREMENTS.bat"
            exit /b !errorlevel!
        )
    )
    echo [ERROR] Node.js 20+ or Python 3 is still missing. Run 00_DOWNLOAD_REQUIREMENTS.bat first!
    pause
    exit /b 1
)

for /f "delims=" %%v in ('node -v') do echo       OK - Node.js %%v detected.
echo.
echo [2/5] Checking Python 3...
for /f "delims=" %%v in ('!PY_CMD! --version') do echo       OK - %%v detected.

:: ------------------------------------------------------------------
:: [3/5] Install Gemini CLI (Free Official Google Agent)
:: ------------------------------------------------------------------
echo.
echo [3/5] Installing official Google Gemini CLI (@google/gemini-cli) [Optional if using Ollama/Manual]...
if not exist "%APPDATA%\npm" mkdir "%APPDATA%\npm"
call npm.cmd config set prefix "%APPDATA%\npm" >nul 2>&1
if "%GEMINI_CLI_VER%"=="" set "GEMINI_CLI_VER=0"
call npm.cmd install -g @google/gemini-cli@%GEMINI_CLI_VER% --no-audit --no-fund
if %errorlevel% NEQ 0 (
    echo [WARN] Could not install @google/gemini-cli. You can still use Ollama or Gemini API keys directly.
) else (
    set "PATH=%APPDATA%\npm;%LOCAL_NODE%;%ProgramFiles%\nodejs;%PATH%"
    echo       OK - Gemini CLI installed successfully.
)

:: ------------------------------------------------------------------
:: [4/5] Install Python Document & Security Libraries from requirements.txt
:: ------------------------------------------------------------------
echo.
echo [4/5] Installing local PDF/Image/Excel/Word/XML libraries from requirements.txt...
!PY_CMD! -m pip install --upgrade --user -r "%SCRIPT_DIR%requirements.txt"
if %errorlevel% NEQ 0 (
    !PY_CMD! -m pip install --upgrade -r "%SCRIPT_DIR%requirements.txt"
)
echo       OK - Document, Excel, and Word libraries ready.

:: ------------------------------------------------------------------
:: [5/5] Install Multi-Project Visual Hub, 12 Commands & Offline Tools
:: ------------------------------------------------------------------
echo.
echo [5/5] Installing Projects Visual Hub, WhatsApp Bridge, and Offline Tools...
if not exist "%GEMINI_HOME%" mkdir "%GEMINI_HOME%"
if not exist "%GEMINI_HOME%\commands" mkdir "%GEMINI_HOME%\commands"
if not exist "%GEMINI_HOME%\doc_tools" mkdir "%GEMINI_HOME%\doc_tools"
if not exist "%GEMINI_HOME%\templates" mkdir "%GEMINI_HOME%\templates"

if exist "%GEMINI_HOME%\GEMINI.md" (
    copy /Y "%GEMINI_HOME%\GEMINI.md" "%GEMINI_HOME%\GEMINI.md.bak" >nul
)
copy /Y "%SCRIPT_DIR%GEMINI.md" "%GEMINI_HOME%\GEMINI.md" >nul
copy /Y "%SCRIPT_DIR%commands\*.toml" "%GEMINI_HOME%\commands\" >nul
copy /Y "%SCRIPT_DIR%doc_tools\*.py" "%GEMINI_HOME%\doc_tools\" >nul
copy /Y "%SCRIPT_DIR%doc_tools\*.html" "%GEMINI_HOME%\doc_tools\" >nul
copy /Y "%SCRIPT_DIR%doc_tools\*.json" "%GEMINI_HOME%\doc_tools\" >nul
if exist "%SCRIPT_DIR%doc_tools\app_icon.ico" copy /Y "%SCRIPT_DIR%doc_tools\app_icon.ico" "%GEMINI_HOME%\doc_tools\app_icon.ico" >nul
if exist "%SCRIPT_DIR%doc_tools\app_icon_192.png" copy /Y "%SCRIPT_DIR%doc_tools\app_icon_192.png" "%GEMINI_HOME%\doc_tools\app_icon_192.png" >nul
if exist "%SCRIPT_DIR%doc_tools\app_icon_512.png" copy /Y "%SCRIPT_DIR%doc_tools\app_icon_512.png" "%GEMINI_HOME%\doc_tools\app_icon_512.png" >nul
copy /Y "%SCRIPT_DIR%0_OPEN_PROJECTS_APP.bat" "%GEMINI_HOME%\Open_Projects_App.bat" >nul
copy /Y "%SCRIPT_DIR%0_OPEN_SILENT.vbs" "%GEMINI_HOME%\Open_Projects_Silent.vbs" >nul
copy /Y "%SCRIPT_DIR%2_START_AGENT.bat" "%GEMINI_HOME%\Start_Smart_Agent.bat" >nul

set "TARGET_VBS=%GEMINI_HOME%\Open_Projects_Silent.vbs"
set "ICON_PATH=%GEMINI_HOME%\doc_tools\app_icon.ico"
powershell -NoProfile -Command "$dt=[Environment]::GetFolderPath('Desktop'); $ws=New-Object -ComObject WScript.Shell; $sc=$ws.CreateShortcut((Join-Path $dt 'Smart Projects Hub.lnk')); $sc.TargetPath='wscript.exe'; $sc.Arguments='//nologo ""' + $env:TARGET_VBS + '""'; $sc.WorkingDirectory=(Split-Path $env:TARGET_VBS); if(Test-Path $env:ICON_PATH){$sc.IconLocation=$env:ICON_PATH}; $sc.Description='Smart Projects Hub v10.13.2'; $sc.Save(); $oldAgent=Join-Path $dt 'Smart_Folder_Agent.bat'; if(Test-Path $oldAgent){Remove-Item -Force $oldAgent -ErrorAction SilentlyContinue}" >nul 2>&1

if exist "%USERPROFILE%\Desktop" (
    if not exist "%USERPROFILE%\Desktop\Smart_Projects_Hub" mkdir "%USERPROFILE%\Desktop\Smart_Projects_Hub"
    if exist "%USERPROFILE%\Desktop\Smart_Folder_Agent.bat" del /F /Q "%USERPROFILE%\Desktop\Smart_Folder_Agent.bat" >nul 2>&1
    echo       OK - Silent Desktop Icon created: Smart Projects Hub
    echo            ^(Double-click anytime to open without a black window!^)
)

echo.
echo ======================================================================
echo   SETUP COMPLETE! ($0 Cost - No Credit Card Needed)
echo ======================================================================
echo   Launching your Arabic Smart Projects Hub Visual App silently now...
echo   ^(On first open, click the yellow bar to paste your FREE 10-second
echo    Google AI Studio key from https://aistudio.google.com/apikey ^)
echo ======================================================================
echo.
wscript.exe //nologo "%TARGET_VBS%"
exit /b 0
