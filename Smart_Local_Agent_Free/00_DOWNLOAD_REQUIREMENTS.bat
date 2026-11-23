@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Auto Downloader - Node.js 22 LTS + Python 3.12 (Verified Checksum)
color 0A

echo ======================================================================
echo   FAST DIRECT DOWNLOADER ^& VERIFIER (Node.js 22 LTS + Python 3.12.10)
echo   Uses native curl.exe + certutil SHA256/MD5 checksum verification
echo   Tip: Do NOT click inside this window while downloading.
echo ======================================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "DL_DIR=%SCRIPT_DIR%Downloaded_Installers"
if not exist "%DL_DIR%" mkdir "%DL_DIR%"

set "NODE_VER=v22.22.0"
set "NODE_URL=https://nodejs.org/dist/%NODE_VER%/node-%NODE_VER%-x64.msi"
set "NODE_MSI_SHA256=b10f88c6ded24ca487839b3eccb8870a08d7f9fc2b9bb3b463fc72a3a40bcdb1"
set "NODE_ZIP_URL=https://nodejs.org/dist/%NODE_VER%/node-%NODE_VER%-win-x64.zip"
set "NODE_ZIP_SHA256=c97fa376d2becdc8863fcd3ca2dd9a83a9f3468ee7ccf7a6d076ec66a645c77a"
set "NODE_MSI=%DL_DIR%\node-%NODE_VER%-x64.msi"
set "NODE_ZIP=%DL_DIR%\node-%NODE_VER%-win-x64.zip"
set "LOCAL_NODE=%LOCALAPPDATA%\Programs\nodejs"

set "PY_URL=https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"
set "PY_MD5=5eddb0b6f12c852725de071ae681dde4"
set "PY_EXE=%DL_DIR%\python-3.12.10-amd64.exe"
set "LOCAL_PY=%LOCALAPPDATA%\Programs\Python\Python312"

set "PATH=%LOCAL_NODE%;%ProgramFiles%\nodejs;%APPDATA%\npm;%LOCALAPPDATA%\Programs\Python\Python313;%LOCAL_PY%;%LOCALAPPDATA%\Programs\Python\Python311;%PATH%"

set "HAS_CURL=0"
where curl.exe >nul 2>&1
if %errorlevel%==0 set "HAS_CURL=1"

:: ==================================================================
:: [1/2] DOWNLOAD, VERIFY SHA256 & INSTALL NODE.JS 22 LTS
:: ==================================================================
echo [1/2] Checking Node.js 20+ ...
set "NEED_NODE=1"
where node >nul 2>&1
if %errorlevel%==0 (
    for /f "tokens=1 delims=." %%v in ('node -v 2^>nul') do (
        set "NVER=%%v"
        set "NVER=!NVER:v=!"
        if !NVER! GEQ 20 set "NEED_NODE=0"
    )
)

if "!NEED_NODE!"=="0" (
    for /f "delims=" %%v in ('node -v') do echo       Already installed: Node.js %%v - Skipping download!
) else (
    if not exist "%NODE_MSI%" (
        echo       Downloading Node.js %NODE_VER% LTS ^(x64 MSI^) into Downloaded_Installers\ ...
        if "!HAS_CURL!"=="1" (
            curl.exe -fL --progress-bar --connect-timeout 20 --retry 3 "%NODE_URL%" -o "%NODE_MSI%.part"
        ) else (
            bitsadmin /transfer NodeDownload /download /priority FOREGROUND "%NODE_URL%" "%NODE_MSI%.part"
        )
        if exist "%NODE_MSI%.part" move /Y "%NODE_MSI%.part" "%NODE_MSI%" >nul
    ) else (
        echo       Found cached installer: %NODE_MSI%
    )

    if exist "%NODE_MSI%" (
        call :VerifyHash "%NODE_MSI%" SHA256 "%NODE_MSI_SHA256%"
        if errorlevel 1 (
            echo [ERROR] Node.js MSI checksum mismatch! Deleting untrusted file.
            del /f /q "%NODE_MSI%" >nul 2>&1
        ) else (
            echo       [OK] Node.js MSI SHA256 verified. Installing...
            msiexec.exe /i "%NODE_MSI%" /passive /norestart
            set "PATH=%ProgramFiles%\nodejs;%APPDATA%\npm;%PATH%"
        )
    )

    :: Fallback: if user has no Admin rights (cancelled UAC), install portable ZIP with tar.exe
    if not exist "%ProgramFiles%\nodejs\node.exe" (
        where tar.exe >nul 2>&1
        if !errorlevel!==0 (
            echo       Installing Node.js %NODE_VER% without Admin rights ^(verified ZIP^)...
            if not exist "%NODE_ZIP%" (
                curl.exe -fL --progress-bar --connect-timeout 20 --retry 3 "%NODE_ZIP_URL%" -o "%NODE_ZIP%.part"
                if exist "%NODE_ZIP%.part" move /Y "%NODE_ZIP%.part" "%NODE_ZIP%" >nul
            )
            if exist "%NODE_ZIP%" (
                call :VerifyHash "%NODE_ZIP%" SHA256 "%NODE_ZIP_SHA256%"
                if errorlevel 1 (
                    echo [ERROR] Node.js ZIP checksum mismatch! Deleting untrusted file.
                    del /f /q "%NODE_ZIP%" >nul 2>&1
                ) else (
                    echo       [OK] Node.js ZIP SHA256 verified. Extracting...
                    if not exist "%LOCALAPPDATA%\Programs" mkdir "%LOCALAPPDATA%\Programs"
                    if exist "%LOCAL_NODE%" rmdir /S /Q "%LOCAL_NODE%"
                    tar.exe -xf "%NODE_ZIP%" -C "%LOCALAPPDATA%\Programs"
                    if exist "%LOCALAPPDATA%\Programs\node-%NODE_VER%-win-x64" (
                        ren "%LOCALAPPDATA%\Programs\node-%NODE_VER%-win-x64" "nodejs"
                    )
                    set "PATH=%LOCAL_NODE%;%APPDATA%\npm;%PATH%"
                )
            )
        )
    )

    where node >nul 2>&1
    if !errorlevel!==0 (
        for /f "delims=" %%v in ('node -v') do echo       [OK] Node.js %%v ready!
    ) else (
        echo       [ERROR] Could not download/verify Node.js. Opening official download page...
        start "" "%NODE_URL%"
        pause
        exit /b 1
    )
)

:: ==================================================================
:: [2/2] DOWNLOAD, VERIFY MD5 & INSTALL PYTHON 3.12.10
:: ==================================================================
echo.
echo [2/2] Checking Python 3 ...
set "PY_OK=0"
python -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
if %errorlevel%==0 set "PY_OK=1"
if "!PY_OK!"=="0" (
    py -3 -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
    if !errorlevel!==0 set "PY_OK=1"
)
if "!PY_OK!"=="0" (
    if exist "%LOCAL_PY%\python.exe" set "PY_OK=1"
)

if "!PY_OK!"=="1" (
    echo       Already installed: Python 3 - Skipping download!
) else (
    if not exist "%PY_EXE%" (
        echo       Downloading Python 3.12.10 ^(25.7 MB^) into Downloaded_Installers\ ...
        if "!HAS_CURL!"=="1" (
            curl.exe -fL --progress-bar --connect-timeout 20 --retry 3 "%PY_URL%" -o "%PY_EXE%.part"
        ) else (
            bitsadmin /transfer PyDownload /download /priority FOREGROUND "%PY_URL%" "%PY_EXE%.part"
        )
        if exist "%PY_EXE%.part" move /Y "%PY_EXE%.part" "%PY_EXE%" >nul
    ) else (
        echo       Found cached installer: %PY_EXE%
    )

    if not exist "%PY_EXE%" (
        echo       [ERROR] Could not download Python. Opening browser link instead...
        start "" "%PY_URL%"
        pause
        exit /b 1
    )

    call :VerifyHash "%PY_EXE%" MD5 "%PY_MD5%"
    if errorlevel 1 (
        echo [ERROR] Python 3.12.10 installer checksum mismatch! Deleting untrusted file.
        del /f /q "%PY_EXE%" >nul 2>&1
        pause
        exit /b 1
    )
    echo       [OK] Python 3.12.10 checksum verified.

    echo       Installing Python 3.12.10 automatically ^(progress window will open^)...
    "%PY_EXE%" /passive InstallAllUsers=0 PrependPath=1 Include_test=0 Include_pip=1 Include_launcher=0
    set "PATH=%LOCAL_PY%;%LOCAL_PY%\Scripts;%PATH%"

    if exist "%LOCAL_PY%\python.exe" (
        for /f "delims=" %%v in ('"%LOCAL_PY%\python.exe" --version') do echo       [OK] %%v ready!
    ) else (
        python -c "import sys; sys.exit(0)" >nul 2>&1
        if !errorlevel!==0 (
            echo       [OK] Python 3 ready!
        ) else (
            echo       [ERROR] Python install did not finish. Double-click:
            echo               %PY_EXE%
            pause
            exit /b 1
        )
    )
)

echo.
echo ======================================================================
echo   [SUCCESS] Node.js %NODE_VER% LTS + Python 3.12.10 are verified and ready!
echo   Saved copies kept in: %DL_DIR%
echo ======================================================================
echo.

if exist "%SCRIPT_DIR%1_INSTALL_FIRST_TIME.bat" (
    echo   Launching 1_INSTALL_FIRST_TIME.bat to finish setup...
    call "%SCRIPT_DIR%1_INSTALL_FIRST_TIME.bat" --after-dl
)
exit /b %errorlevel%

:VerifyHash
set "_VF=%~1"
set "_ALG=%~2"
set "_EXP=%~3"
set "_ACTUAL="
for /f "skip=1 delims=" %%H in ('certutil -hashfile "%_VF%" %_ALG% 2^>nul') do (
    if not defined _ACTUAL (
        set "_LINE=%%H"
        set "_LINE=!_LINE: =!"
        set "_ACTUAL=!_LINE!"
    )
)
if /i "!_ACTUAL!"=="%_EXP%" exit /b 0
exit /b 1
