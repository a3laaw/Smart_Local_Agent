@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1
title Create Silent Desktop Icon - Smart Projects Hub v10.13.2
color 0B

set "SCRIPT_DIR=%~dp0"
set "GEMINI_HOME=%USERPROFILE%\.gemini"

if exist "%GEMINI_HOME%\doc_tools\project_dashboard.py" (
    copy /Y "%SCRIPT_DIR%0_OPEN_SILENT.vbs" "%GEMINI_HOME%\Open_Projects_Silent.vbs" >nul 2>&1
    if exist "%SCRIPT_DIR%doc_tools\app_icon.ico" (
        copy /Y "%SCRIPT_DIR%doc_tools\app_icon.ico" "%GEMINI_HOME%\doc_tools\app_icon.ico" >nul 2>&1
    )
    if exist "%SCRIPT_DIR%doc_tools\app_icon_192.png" (
        copy /Y "%SCRIPT_DIR%doc_tools\app_icon_192.png" "%GEMINI_HOME%\doc_tools\app_icon_192.png" >nul 2>&1
    )
    if exist "%SCRIPT_DIR%doc_tools\app_icon_512.png" (
        copy /Y "%SCRIPT_DIR%doc_tools\app_icon_512.png" "%GEMINI_HOME%\doc_tools\app_icon_512.png" >nul 2>&1
    )
    set "TARGET_VBS=%GEMINI_HOME%\Open_Projects_Silent.vbs"
    set "ICON_PATH=%GEMINI_HOME%\doc_tools\app_icon.ico"
) else (
    set "TARGET_VBS=%SCRIPT_DIR%0_OPEN_SILENT.vbs"
    set "ICON_PATH=%SCRIPT_DIR%doc_tools\app_icon.ico"
)

powershell -NoProfile -Command "$dt=[Environment]::GetFolderPath('Desktop'); $ws=New-Object -ComObject WScript.Shell; $sc=$ws.CreateShortcut((Join-Path $dt 'Smart Projects Hub.lnk')); $sc.TargetPath='wscript.exe'; $sc.Arguments='//nologo ""' + $env:TARGET_VBS + '""'; $sc.WorkingDirectory=(Split-Path $env:TARGET_VBS); if(Test-Path $env:ICON_PATH){$sc.IconLocation=$env:ICON_PATH}; $sc.Description='Smart Projects Hub v10.13.2 (Silent Launcher)'; $sc.Save(); Write-Host ('[OK] Desktop Icon created: ' + $sc.FullName)"

if %errorlevel% NEQ 0 (
    echo [ERROR] Could not create Desktop shortcut automatically.
    echo         You can double-click 0_OPEN_SILENT.vbs directly without any black window!
    pause
    exit /b 1
)

echo.
echo [OK] Done! Look at your Desktop for the "Smart Projects Hub" icon.
echo      Double-clicking it opens the app cleanly with NO black background window!
echo.
wscript.exe //nologo "%TARGET_VBS%"
exit /b 0
