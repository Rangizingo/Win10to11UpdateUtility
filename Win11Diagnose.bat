@echo off
setlocal EnableDelayedExpansion

:: Configuration
set "PC=01INVENTORY-PC"
set "LOGFILE=%~dp0Win11Diagnose.log"

:: Clear log file
echo. > "%LOGFILE%"

call :LOG "============================================================"
call :LOG "Windows 11 Upgrade Diagnostic Report"
call :LOG "============================================================"
call :LOG "Timestamp: %DATE% %TIME%"
call :LOG "Target PC: %PC%"
call :LOG "============================================================"
call :LOG ""

:: ============================================================
:: SECTION 1: CONNECTIVITY
:: ============================================================
call :LOG "[SECTION 1] CONNECTIVITY CHECKS"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Ping test..."
ping -n 1 -w 1000 %PC% >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    call :LOG "[OK] %PC% is reachable"
) else (
    call :LOG "[FAIL] %PC% is NOT reachable"
    call :LOG "[SOLUTION] Ensure PC is powered on and connected to network"
    goto :END
)

call :LOG "[CHECK] Admin share accessible..."
if exist "\\%PC%\C$" (
    call :LOG "[OK] Admin share \\%PC%\C$ is accessible"
) else (
    call :LOG "[FAIL] Cannot access admin share"
    call :LOG "[SOLUTION] Ensure you have admin rights and File Sharing is enabled"
)
call :LOG ""

:: ============================================================
:: SECTION 2: INSTALLER FILE
:: ============================================================
call :LOG "[SECTION 2] INSTALLER FILE CHECKS"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Installer exists on remote PC..."
if exist "\\%PC%\C$\Temp\Windows11InstallationAssistant.exe" (
    call :LOG "[OK] Installer found at C:\Temp\Windows11InstallationAssistant.exe"

    for %%A in ("\\%PC%\C$\Temp\Windows11InstallationAssistant.exe") do set "FILESIZE=%%~zA"
    call :LOG "[INFO] File size: !FILESIZE! bytes"

    if !FILESIZE! LSS 1000000 (
        call :LOG "[WARN] File size seems too small - may be corrupted"
        call :LOG "[SOLUTION] Re-download the installer"
    ) else (
        call :LOG "[OK] File size appears valid"
    )
) else (
    call :LOG "[FAIL] Installer NOT found"
    call :LOG "[SOLUTION] Run the upgrade script to download installer first"
)
call :LOG ""

:: ============================================================
:: SECTION 3: REGISTRY BYPASS
:: ============================================================
call :LOG "[SECTION 3] WINDOWS 11 BYPASS REGISTRY CHECK"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] TPM/CPU bypass registry key..."
psexec \\%PC% -s cmd /c "reg query HKLM\SYSTEM\Setup\MoSetup /v AllowUpgradesWithUnsupportedTPMOrCPU" 2>&1 | findstr "0x1" >nul
if %ERRORLEVEL% EQU 0 (
    call :LOG "[OK] Bypass registry key is set (0x1)"
) else (
    call :LOG "[FAIL] Bypass registry key NOT set or incorrect"
    call :LOG "[SOLUTION] Run: psexec \\%PC% -s reg add HKLM\SYSTEM\Setup\MoSetup /v AllowUpgradesWithUnsupportedTPMOrCPU /t REG_DWORD /d 1 /f"
)
call :LOG ""

:: ============================================================
:: SECTION 4: RUNNING PROCESSES
:: ============================================================
call :LOG "[SECTION 4] PROCESS CHECKS"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Windows11 related processes..."
psexec \\%PC% -s cmd /c "tasklist | findstr /i Windows11" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    call :LOG "[INFO] Windows11 process IS running"
    psexec \\%PC% -s cmd /c "tasklist | findstr /i Windows11" >> "%LOGFILE%" 2>&1
) else (
    call :LOG "[INFO] No Windows11 process currently running"
)

call :LOG "[CHECK] Setup related processes..."
psexec \\%PC% -s cmd /c "tasklist | findstr /i setup" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    call :LOG "[INFO] Setup process IS running:"
    psexec \\%PC% -s cmd /c "tasklist | findstr /i setup" >> "%LOGFILE%" 2>&1
) else (
    call :LOG "[INFO] No setup process currently running"
)
call :LOG ""

:: ============================================================
:: SECTION 5: UPGRADE STAGING FOLDER
:: ============================================================
call :LOG "[SECTION 5] UPGRADE STAGING FOLDER"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Windows upgrade staging folder..."
if exist "\\%PC%\C$\$WINDOWS.~BT" (
    call :LOG "[OK] Staging folder C:\$WINDOWS.~BT EXISTS - upgrade has started"
    call :LOG "[INFO] Checking for setup log..."
    if exist "\\%PC%\C$\$WINDOWS.~BT\Sources\Panther\setupact.log" (
        call :LOG "[OK] Setup log exists - upgrade is in progress or completed"
    ) else (
        call :LOG "[INFO] Setup log not yet created - upgrade may be early stage"
    )
) else (
    call :LOG "[INFO] Staging folder does NOT exist - upgrade has not started"
    call :LOG "[SOLUTION] This means the installer never began the upgrade process"
)
call :LOG ""

:: ============================================================
:: SECTION 6: DISK SPACE
:: ============================================================
call :LOG "[SECTION 6] DISK SPACE CHECK"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Free disk space on C: drive..."
for /f "tokens=2" %%a in ('psexec \\%PC% -s cmd /c "wmic logicaldisk where DeviceID='C:' get FreeSpace /format:value" ^| findstr FreeSpace') do set "FREESPACE=%%a"
set /a "FREEGB=!FREESPACE:~0,-9!"
call :LOG "[INFO] Free space: approximately !FREEGB! GB"

if !FREEGB! LSS 64 (
    call :LOG "[WARN] Less than 64GB free - may not be enough for upgrade"
    call :LOG "[SOLUTION] Free up disk space on C: drive"
) else (
    call :LOG "[OK] Sufficient disk space available"
)
call :LOG ""

:: ============================================================
:: SECTION 7: CURRENT OS VERSION
:: ============================================================
call :LOG "[SECTION 7] CURRENT OS VERSION"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Current Windows version..."
psexec \\%PC% -s cmd /c "wmic os get Caption,Version,BuildNumber /format:list" >> "%LOGFILE%" 2>&1
call :LOG "[INFO] See above for OS details"
call :LOG ""

:: ============================================================
:: SECTION 8: EVENT LOGS
:: ============================================================
call :LOG "[SECTION 8] RECENT EVENT LOG ENTRIES"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Recent Application errors..."
psexec \\%PC% -s cmd /c "wevtutil qe Application /c:5 /f:text /rd:true /q:*[System[Level=2]]" >> "%LOGFILE%" 2>&1

call :LOG ""
call :LOG "[CHECK] Recent System errors..."
psexec \\%PC% -s cmd /c "wevtutil qe System /c:5 /f:text /rd:true /q:*[System[Level=2]]" >> "%LOGFILE%" 2>&1
call :LOG ""

:: ============================================================
:: SECTION 9: INSTALLER LOGS
:: ============================================================
call :LOG "[SECTION 9] INSTALLER LOG LOCATIONS"
call :LOG "------------------------------------------------------------"

call :LOG "[CHECK] Checking common log locations..."

if exist "\\%PC%\C$\Windows\Logs\MoSetup" (
    call :LOG "[FOUND] C:\Windows\Logs\MoSetup exists"
    dir "\\%PC%\C$\Windows\Logs\MoSetup" >> "%LOGFILE%" 2>&1
) else (
    call :LOG "[INFO] C:\Windows\Logs\MoSetup does not exist"
)

if exist "\\%PC%\C$\Windows\Panther" (
    call :LOG "[FOUND] C:\Windows\Panther exists"
    dir "\\%PC%\C$\Windows\Panther\*.log" >> "%LOGFILE%" 2>&1
) else (
    call :LOG "[INFO] C:\Windows\Panther does not exist"
)
call :LOG ""

:: ============================================================
:: SECTION 10: DIAGNOSIS SUMMARY
:: ============================================================
call :LOG "============================================================"
call :LOG "DIAGNOSIS SUMMARY"
call :LOG "============================================================"
call :LOG ""
call :LOG "LIKELY ISSUES:"
call :LOG ""
call :LOG "1. INSTALLATION ASSISTANT SILENT MODE"
call :LOG "   The Windows 11 Installation Assistant may not support"
call :LOG "   true silent installation. The /quietinstall flag may"
call :LOG "   not work as expected."
call :LOG ""
call :LOG "   SOLUTION: Use Windows 11 ISO setup.exe instead:"
call :LOG "   setup.exe /auto upgrade /quiet /eula accept"
call :LOG ""
call :LOG "2. SESSION 0 ISOLATION"
call :LOG "   Running with -s (SYSTEM account) runs in Session 0"
call :LOG "   which blocks GUI applications from starting."
call :LOG ""
call :LOG "   SOLUTION: Run without -s flag or use -i flag with"
call :LOG "   an active user session:"
call :LOG "   psexec \\%PC% -i \"C:\Temp\Windows11InstallationAssistant.exe\""
call :LOG ""
call :LOG "3. USER INTERACTION REQUIRED"
call :LOG "   The Installation Assistant may require clicking through"
call :LOG "   prompts even with /skipeula flag."
call :LOG ""
call :LOG "   SOLUTION: Use ISO-based setup.exe which has better"
call :LOG "   unattended support, or RDP in to complete manually."
call :LOG ""
call :LOG "RECOMMENDED NEXT STEPS:"
call :LOG ""
call :LOG "Option A - Test interactively:"
call :LOG "   1. RDP into %PC%"
call :LOG "   2. Run: C:\Temp\Windows11InstallationAssistant.exe"
call :LOG "   3. Observe what happens"
call :LOG ""
call :LOG "Option B - Switch to ISO method:"
call :LOG "   1. Download Windows 11 ISO from Microsoft"
call :LOG "   2. Extract to network share"
call :LOG "   3. Run: setup.exe /auto upgrade /quiet /eula accept"
call :LOG ""
call :LOG "============================================================"
call :LOG "END OF DIAGNOSTIC REPORT"
call :LOG "============================================================"

goto :END

:LOG
echo %~1
echo %~1 >> "%LOGFILE%"
goto :EOF

:END
call :LOG ""
call :LOG "Report generated at: %DATE% %TIME%"
echo.
echo Full report saved to: %LOGFILE%
echo.
pause
exit /b
