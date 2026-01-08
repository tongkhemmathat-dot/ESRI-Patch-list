@echo off
setlocal EnableExtensions EnableDelayedExpansion
title ArcGIS Offline Patch Installer

REM ==================================================
REM CONFIG
REM ==================================================
set "PATCH_DIR=%~dp0"
set "LOG_FILE=%PATCH_DIR%patch_install.log"
set "BAR_WIDTH=30"

REM ==================================================
REM Header
REM ==================================================
cls
echo ==================================================
echo   ArcGIS Offline Patch Installer (MSP)
echo ==================================================
echo Patch folder : %PATCH_DIR%
echo Start time   : %DATE% %TIME%
echo.

echo ================================================== >> "%LOG_FILE%"
echo Patch install started at %DATE% %TIME% >> "%LOG_FILE%"
echo Patch folder: %PATCH_DIR% >> "%LOG_FILE%"
echo ================================================== >> "%LOG_FILE%"

REM ==================================================
REM Check MSP files exist
REM ==================================================
dir "%PATCH_DIR%*.msp" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No .msp files found in this folder.
    echo [ERROR] No .msp files found >> "%LOG_FILE%"
    call :Popup "ArcGIS Patch Installer" "No .msp files found in:%PATCH_DIR%"
    pause
    exit /b 1
)

REM ==================================================
REM Count total patches
REM ==================================================
set /a TOTAL=0, DONE=0, SUCCESS=0, FAIL=0, SKIP=0
for %%F in ("%PATCH_DIR%*.msp") do set /a TOTAL+=1

echo Total patches found: %TOTAL%
echo.

REM ==================================================
REM Main loop
REM ==================================================
for %%F in ("%PATCH_DIR%*.msp") do (
    set /a DONE+=1
    call :DrawProgress !DONE! !TOTAL! %BAR_WIDTH%

    echo --------------------------------------------------
    echo [!DONE!/!TOTAL!] START  : %%~nxF
    echo Start time: %DATE% %TIME%
    echo.

    echo Installing %%~nxF >> "%LOG_FILE%"

    REM Per-patch log
    set "PATCH_LOG=%PATCH_DIR%%%~nF.log"

    REM Install
    msiexec /p "%%F" /qn /norestart /log "!PATCH_LOG!"
    set "RC=!ERRORLEVEL!"

    REM Interpret exit codes (practical for MSP patching)
    REM 0     = success
    REM 3010  = success, reboot required
    REM 1641  = success, reboot initiated/required (rare with /norestart but handle)
    REM 1638  = another version already installed (often means already applied)
    REM 1642  = update not applicable (e.g., wrong product/version) -> treat as skip
    REM others= failed
    if "!RC!"=="0" (
        echo [SUCCESS] %%~nxF
        echo Result : SUCCESS (RC=0) >> "%LOG_FILE%"
        set /a SUCCESS+=1
    ) else if "!RC!"=="3010" (
        echo [SUCCESS] %%~nxF  ^(Reboot required: RC=3010^)
        echo Result : SUCCESS - REBOOT REQUIRED (RC=3010) >> "%LOG_FILE%"
        set /a SUCCESS+=1
    ) else if "!RC!"=="1641" (
        echo [SUCCESS] %%~nxF  ^(Reboot required: RC=1641^)
        echo Result : SUCCESS - REBOOT REQUIRED (RC=1641) >> "%LOG_FILE%"
        set /a SUCCESS+=1
    ) else if "!RC!"=="1638" (
        echo [SKIPPED] %%~nxF  ^(Already installed / another version present: RC=1638^)
        echo Result : SKIPPED - ALREADY INSTALLED/ANOTHER VERSION (RC=1638) >> "%LOG_FILE%"
        set /a SKIP+=1
    ) else if "!RC!"=="1642" (
        echo [SKIPPED] %%~nxF  ^(Not applicable: RC=1642^)
        echo Result : SKIPPED - NOT APPLICABLE (RC=1642) >> "%LOG_FILE%"
        set /a SKIP+=1
    ) else (
        echo [FAILED ] %%~nxF  ^(RC=!RC!^)
        echo Result : FAILED (RC=!RC!) >> "%LOG_FILE%"
        set /a FAIL+=1
    )

    echo.
)

REM ==================================================
REM Summary
REM ==================================================
call :DrawProgress %TOTAL% %TOTAL% %BAR_WIDTH%
echo ==================================================
echo   INSTALLATION SUMMARY
echo ==================================================
echo Total   : %TOTAL%
echo Success : %SUCCESS%
echo Skipped : %SKIP%
echo Failed  : %FAIL%
echo End time: %DATE% %TIME%
echo Log     : %LOG_FILE%
echo ==================================================

echo Summary - Total:%TOTAL% Success:%SUCCESS% Skipped:%SKIP% Failed:%FAIL% >> "%LOG_FILE%"
echo Patch install finished at %DATE% %TIME% >> "%LOG_FILE%"
echo ================================================== >> "%LOG_FILE%"

REM Popup summary
set "POPMSG=Total: %TOTAL%`nSuccess: %SUCCESS%`nSkipped: %SKIP%`nFailed: %FAIL%`n`nLog: %LOG_FILE%"
call :Popup "ArcGIS Patch Installer - Completed" "%POPMSG%"

pause
exit /b 0

REM ==================================================
REM DrawProgress current total width
REM ==================================================
:DrawProgress
REM Args: %1=current %2=total %3=width
setlocal EnableDelayedExpansion
set /a CUR=%~1, TOT=%~2, W=%~3

REM Prevent divide by zero
if %TOT% LEQ 0 set /a TOT=1

set /a PCT=(CUR*100)/TOT
set /a FILLED=(CUR*W)/TOT
set "BAR="

for /L %%i in (1,1,!FILLED!) do set "BAR=!BAR!#"
for /L %%i in (!FILLED!,1,%W%-1) do set "BAR=!BAR!-"

REM Carriage return style line refresh (works well in most cmd)
<nul set /p "=Progress: [!BAR!] !PCT!%% (!CUR!/!TOT!)   "
echo.
endlocal
exit /b 0

REM ==================================================
REM Popup Title Message
REM ==================================================
:Popup
REM Uses PowerShell MessageBox (no external files required)
setlocal
set "PTITLE=%~1"
set "PMESSAGE=%~2"
powershell -NoProfile -Command ^
"Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('%PMESSAGE%','%PTITLE%','OK','Information') | Out-Null"
endlocal
exit /b 0
