@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Background Startup Control - Smart Projects Hub & Scanner Watcher
color 0E

set "PATH=%LOCALAPPDATA%\Programs\nodejs;%ProgramFiles%\nodejs;%APPDATA%\npm;%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"
if exist "%USERPROFILE%\.gemini\doc_tools\project_dashboard.py" (
    set "APP=%USERPROFILE%\.gemini\doc_tools\project_dashboard.py"
) else (
    set "APP=%~dp0doc_tools\project_dashboard.py"
)
set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_FILE=%STARTUP_DIR%\Smart_Projects_Hub_AutoStart.vbs"

if not exist "%APP%" (
    echo [ERROR] Please run 1_INSTALL_FIRST_TIME.bat first!
    pause
    exit /b 1
)

echo ======================================================================
echo   AUTO-START PROJECTS HUB + ALL PROJECT SCANNERS WITH WINDOWS ($0)
echo ======================================================================
echo   [1] ENABLE  Auto-Start silently with Windows + Open Projects Hub NOW
echo   [2] DISABLE Auto-Start and Stop Background Service
echo   [3] Cancel
echo ======================================================================
echo.
set /p "CH=Choose [1-3]: "

if "%CH%"=="1" (
    set "PY_CMD="
    python -c "import sys" >nul 2>&1
    if %errorlevel%==0 set "PY_CMD=python"
    if "!PY_CMD!"=="" (
        py -3 -c "import sys" >nul 2>&1
        if !errorlevel!==0 set "PY_CMD=py -3"
    )
    if "!PY_CMD!"=="" (
        echo [ERROR] Python 3 was not found. Please run 1_INSTALL_FIRST_TIME.bat first!
        pause
        exit /b 1
    )
    set "PYW="
    for /f "delims=" %%P in ('!PY_CMD! -c "import sys,os; d=os.path.dirname(sys.executable); pw=os.path.join(d,'pythonw.exe'); print(pw if os.path.exists(pw) else sys.executable)"') do set "PYW=%%P"
    if "!PYW!"=="" (
        echo [ERROR] Could not locate pythonw.exe / python.exe!
        pause
        exit /b 1
    )

    (
        echo Set WshShell = CreateObject^("WScript.Shell"^)
        echo WshShell.Run """!PYW!"" ""%APP%"" --host 127.0.0.1 --port 8765 --no-browser", 0, False
    ) > "%VBS_FILE%"

    powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*project_dashboard.py*' -or $_.CommandLine -like '*scan_watcher.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }" >nul 2>&1
    wscript.exe "%VBS_FILE%"
    timeout /t 1 >nul
    start "" "http://127.0.0.1:8765"
    echo.
    echo [OK] Projects Hub + 24/7 Scanner Auto-Watcher is RUNNING in background!
    echo      Every project folder is watched automatically whenever your PC is on.
    pause
    exit /b 0
)

if "%CH%"=="2" (
    if exist "%VBS_FILE%" del /F /Q "%VBS_FILE%"
    powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*project_dashboard.py*' -or $_.CommandLine -like '*scan_watcher.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }" >nul 2>&1
    echo.
    echo [OK] Background service stopped and removed from Windows Startup.
    pause
    exit /b 0
)

exit /b 0
