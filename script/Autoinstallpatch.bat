@echo off
title ArcGIS Patch Installer
setlocal enabledelayedexpansion

:: Define the log file
set LOG_FILE=%~dp0ArcGIS_Patch_Install.log

:: Clear previous log
echo ArcGIS Patch Installation Log > "%LOG_FILE%"
echo ------------------------------------------- >> "%LOG_FILE%"

:: Loop through all .msp files in the directory
for %%F in (*.msp) do (
    echo Installing patch: %%F...
    echo Installing patch: %%F >> "%LOG_FILE%"
    
    msiexec /p "%%F" /qn /norestart /l*v "%~dp0%%~nF_install_log.txt"

    if !errorlevel! == 0 (
        echo SUCCESS: %%F installed successfully! >> "%LOG_FILE%"
    ) else (
        echo ERROR: Failed to install %%F! Check %%~nF_install_log.txt for details. >> "%LOG_FILE%"
    )
)

echo Installation complete. Check %LOG_FILE% for details.
pause
exit /b 0