@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul 2>&1
title Smart Projects Hub - Visual App v10.13.2 ($0 Free)
color 0B

set "PATH=%LOCALAPPDATA%\Programs\nodejs;%ProgramFiles%\nodejs;%APPDATA%\npm;%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"

:: Always prefer the newest bundled app alongside this .bat; fall back to ~/.gemini copy
set "LOCAL_APP=%~dp0doc_tools\project_dashboard.py"
set "INSTALLED_APP=%USERPROFILE%\.gemini\doc_tools\project_dashboard.py"

if exist "%LOCAL_APP%" (
    set "APP=%LOCAL_APP%"
) else if exist "%INSTALLED_APP%" (
    set "APP=%INSTALLED_APP%"
) else (
    echo [ERROR] Please run 1_INSTALL_FIRST_TIME.bat first!
    pause
    exit /b 1
)

set "PY_CMD=python"
python -c "import sys; sys.exit(0 if sys.version_info.major==3 else 1)" >nul 2>&1
if %errorlevel% NEQ 0 set "PY_CMD=py -3"

:: Check if port 8765 or 8766 is already running AND serving build 2026-11-18-v10.13.2
for %%P in (8765 8766) do (
    !PY_CMD! -c "import json, urllib.request, sys; d=json.loads(urllib.request.urlopen('http://127.0.0.1:%%P/api/state',timeout=1.0).read().decode('utf-8')); sys.exit(0 if d.get('build_tag')=='2026-11-18-v10.13.2' else 2)" >nul 2>&1
    if !errorlevel!==0 (
        start "" "http://127.0.0.1:%%P"
        exit /b 0
    )
)

:: If an older version is holding port 8765, pick port 8766 so the updated UI opens cleanly
!PY_CMD! -c "import socket; s=socket.create_connection(('127.0.0.1',8765),0.5); s.close()" >nul 2>&1
if %errorlevel%==0 (
    echo [INFO] Port 8765 is held by an older instance. Starting updated v10.13.2 on port 8766...
    set "PORT=8766"
) else (
    set "PORT=8765"
)

echo Starting Smart Projects Hub UI (build 2026-11-18-v10.13.2) on port %PORT%...
echo (Minimize this window while working, or use 5_AUTO_START_WITH_WINDOWS.bat to run hidden.)
echo.
!PY_CMD! "%APP%" --host 127.0.0.1 --port %PORT%
pause
