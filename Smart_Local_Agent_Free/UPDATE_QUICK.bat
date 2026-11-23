@echo off
setlocal EnableExtensions
title Quick 1-Second Update - Smart Projects Hub v10.13.2
color 0A

set "SCRIPT_DIR=%~dp0"
set "GEMINI_HOME=%USERPROFILE%\.gemini"

echo ======================================================================
echo   UPDATING SMART PROJECTS HUB TO v10.13.2 (SAFE - ZERO DATA DELETION)
echo ======================================================================
echo   - Persistent System-Wide AI Cache (Zero Duplicate Quota Calls!)
echo   - Collapsible Sidebar + Collapsible Sections Group (Saved Locally)
echo   - Signature Amber/Orange/Gold/Blue/Coral Brand Theme & TopBar
echo   - TopMost Foreground Folder & File Picker Dialogs (Always On Top)
echo ======================================================================
if not exist "%GEMINI_HOME%\doc_tools" mkdir "%GEMINI_HOME%\doc_tools"
if not exist "%GEMINI_HOME%\commands" mkdir "%GEMINI_HOME%\commands"

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
powershell -NoProfile -Command "$dt=[Environment]::GetFolderPath('Desktop'); $ws=New-Object -ComObject WScript.Shell; $sc=$ws.CreateShortcut((Join-Path $dt 'Smart Projects Hub.lnk')); $sc.TargetPath='wscript.exe'; $sc.Arguments='//nologo ""' + $env:TARGET_VBS + '""'; $sc.WorkingDirectory=(Split-Path $env:TARGET_VBS); if(Test-Path $env:ICON_PATH){$sc.IconLocation=$env:ICON_PATH}; $sc.Description='Smart Projects Hub v10.13.2'; $sc.Save(); foreach($old in @('Smart_Folder_Agent.bat','Smart_Projects_App.bat')){$p=Join-Path $dt $old; if(Test-Path $p){Remove-Item -Force $p -ErrorAction SilentlyContinue}}" >nul 2>&1

if exist "%USERPROFILE%\Desktop" (
    if exist "%USERPROFILE%\Desktop\Smart_Projects_App.bat" del /F /Q "%USERPROFILE%\Desktop\Smart_Projects_App.bat" >nul 2>&1
    if exist "%USERPROFILE%\Desktop\Smart_Folder_Agent.bat" del /F /Q "%USERPROFILE%\Desktop\Smart_Folder_Agent.bat" >nul 2>&1
)

:: Stop older dashboard server on ports 8765 / 8766 so v10.13.2 takes over immediately
for /f "tokens=5" %%p in ('netstat -aon ^| findstr ":8765" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%p >nul 2>&1
)
for /f "tokens=5" %%p in ('netstat -aon ^| findstr ":8766" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%p >nul 2>&1
)
timeout /t 1 /nobreak >nul

echo [OK] Update to v10.13.2 complete (all your projects, logs, and templates preserved).
echo      Silent Desktop Icon: "Smart Projects Hub" (no black background window).
wscript.exe //nologo "%TARGET_VBS%"
exit /b 0
