@echo off
setlocal EnableExtensions
title Right-Click Menu Setup - Smart Local Document Agent
color 0E

set "LAUNCHER=%USERPROFILE%\.gemini\Start_Smart_Agent.bat"
if not exist "%LAUNCHER%" (
    if not exist "%USERPROFILE%\.gemini" mkdir "%USERPROFILE%\.gemini"
    copy /Y "%~dp02_START_AGENT.bat" "%LAUNCHER%" >nul
)

echo ======================================================================
echo   WINDOWS RIGHT-CLICK FOLDER MENU (No Admin Required)
echo ======================================================================
echo   [1] ADD    "Open with Smart Document Agent" to Folder Right-Click
echo   [2] REMOVE "Open with Smart Document Agent" from Right-Click
echo   [3] Cancel
echo ======================================================================
echo.
set /p "CHOICE=Choose [1, 2, or 3]: "

if "%CHOICE%"=="1" (
    reg add "HKCU\Software\Classes\Directory\shell\SmartDocAgent" /ve /d "Open with Smart Document Agent (Free)" /f >nul
    reg add "HKCU\Software\Classes\Directory\shell\SmartDocAgent" /v "Icon" /d "imageres.dll,-5302" /f >nul
    reg add "HKCU\Software\Classes\Directory\shell\SmartDocAgent\command" /ve /d "cmd.exe /c \"\"%LAUNCHER%\" \"%%1\"\"" /f >nul

    reg add "HKCU\Software\Classes\Directory\Background\shell\SmartDocAgent" /ve /d "Open with Smart Document Agent (Free)" /f >nul
    reg add "HKCU\Software\Classes\Directory\Background\shell\SmartDocAgent" /v "Icon" /d "imageres.dll,-5302" /f >nul
    reg add "HKCU\Software\Classes\Directory\Background\shell\SmartDocAgent\command" /ve /d "cmd.exe /c \"\"%LAUNCHER%\" \"%%V\"\"" /f >nul

    echo.
    echo [OK] Added! Now right-click ANY folder in Windows Explorer and pick:
    echo      "Open with Smart Document Agent (Free)"
    echo      (On Windows 11: Right-click ^> Show more options, or Shift+Right-click)
    pause
    exit /b 0
)

if "%CHOICE%"=="2" (
    reg delete "HKCU\Software\Classes\Directory\shell\SmartDocAgent" /f >nul 2>&1
    reg delete "HKCU\Software\Classes\Directory\Background\shell\SmartDocAgent" /f >nul 2>&1
    echo.
    echo [OK] Removed from Windows Right-Click menu.
    pause
    exit /b 0
)

exit /b 0
