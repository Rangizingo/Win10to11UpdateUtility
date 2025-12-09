@echo off
setlocal EnableDelayedExpansion

:: Configuration
set "PC=01INVENTORY-PC"
set "REMOTE_PATH=C:\Temp"
set "INSTALLER_NAME=Windows11InstallationAssistant.exe"
set "SCRIPT_DIR=%~dp0"
set "LOGFILE=%~dp0Win11Upgrade.log"

:: Clear log file (fresh each run)
echo. > "%LOGFILE%"

:: Logging function - outputs to both console and file
call :LOG "=========================================="
call :LOG "Windows 11 Upgrade Script Started"
call :LOG "=========================================="
call :LOG "Timestamp: %DATE% %TIME%"
call :LOG "Target PC: %PC%"
call :LOG "Remote Path: %REMOTE_PATH%"
call :LOG "Log File: %LOGFILE%"
call :LOG "------------------------------------------"

:: Step 1: Verify remote PC is online
call :LOG "[STEP 1] Checking if %PC% is online..."
ping -n 1 -w 1000 %PC% >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    call :LOG "[ERROR] %PC% is not reachable. Aborting."
    goto :END
)
call :LOG "[DEBUG] Ping successful. %PC% is online."

:: Step 2: Create remote temp directory
call :LOG "------------------------------------------"
call :LOG "[STEP 2] Creating remote directory %REMOTE_PATH%..."
call :LOG "[DEBUG] Checking if \\%PC%\C$\Temp exists..."
if not exist "\\%PC%\C$\Temp" (
    call :LOG "[DEBUG] Directory does not exist. Creating..."
    mkdir "\\%PC%\C$\Temp" 2>&1
    if %ERRORLEVEL% NEQ 0 (
        call :LOG "[ERROR] Failed to create remote directory. Error code: %ERRORLEVEL%"
        goto :END
    )
    call :LOG "[DEBUG] Directory created successfully."
) else (
    call :LOG "[DEBUG] Directory already exists."
)

:: Step 3: Copy download script to remote PC
call :LOG "------------------------------------------"
call :LOG "[STEP 3] Copying download script to %PC%..."
call :LOG "[DEBUG] Source: %SCRIPT_DIR%DownloadWin11.ps1"
call :LOG "[DEBUG] Destination: \\%PC%\C$\Temp\DownloadWin11.ps1"

copy "%SCRIPT_DIR%DownloadWin11.ps1" "\\%PC%\C$\Temp\DownloadWin11.ps1" /Y >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    call :LOG "[ERROR] Failed to copy download script. Error code: %ERRORLEVEL%"
    goto :END
)
call :LOG "[DEBUG] Download script copied successfully."

:: Step 4: Execute download script on remote PC
call :LOG "------------------------------------------"
call :LOG "[STEP 4] Downloading Windows 11 Installation Assistant on %PC%..."
call :LOG "[DEBUG] Executing PowerShell script remotely..."

psexec \\%PC% -s powershell -ExecutionPolicy Bypass -File "C:\Temp\DownloadWin11.ps1" 2>&1
set "DL_RESULT=%ERRORLEVEL%"
call :LOG "[DEBUG] PsExec returned: %DL_RESULT%"

if %DL_RESULT% NEQ 0 (
    call :LOG "[WARNING] Download may have failed. Error code: %DL_RESULT%"
    call :LOG "[DEBUG] Checking if file exists anyway..."
)

:: Step 5: Verify installer exists on remote PC
call :LOG "------------------------------------------"
call :LOG "[STEP 5] Verifying installer exists on %PC%..."
call :LOG "[DEBUG] Checking for file: \\%PC%\C$\Temp\%INSTALLER_NAME%"

if not exist "\\%PC%\C$\Temp\%INSTALLER_NAME%" (
    call :LOG "[ERROR] Installer not found on remote PC after download."
    goto :END
)

call :LOG "[DEBUG] File exists. Checking file size..."
for %%A in ("\\%PC%\C$\Temp\%INSTALLER_NAME%") do set "FILESIZE=%%~zA"
call :LOG "[DEBUG] File size: %FILESIZE% bytes"

if %FILESIZE% LSS 1000000 (
    call :LOG "[ERROR] File size too small (%FILESIZE% bytes). Download may have failed."
    goto :END
)
call :LOG "[DEBUG] File size OK. Installer verified."

:: Step 6: Launch the upgrade
call :LOG "------------------------------------------"
call :LOG "[STEP 6] Launching Windows 11 upgrade on %PC%..."
call :LOG "[DEBUG] Command: %REMOTE_PATH%\%INSTALLER_NAME% /quietinstall /skipeula /auto upgrade"
call :LOG "[DEBUG] Using -d flag to run detached..."

psexec \\%PC% -s -d "%REMOTE_PATH%\%INSTALLER_NAME%" /quietinstall /skipeula /auto upgrade 2>&1
set "LAUNCH_RESULT=%ERRORLEVEL%"
call :LOG "[DEBUG] PsExec launch returned: %LAUNCH_RESULT%"

if %LAUNCH_RESULT% NEQ 0 (
    call :LOG "[WARNING] PsExec returned error code: %LAUNCH_RESULT%"
    call :LOG "[DEBUG] This may be normal if process started successfully in background."
) else (
    call :LOG "[DEBUG] Upgrade process launched successfully."
)

:: Step 7: Verify process is running
call :LOG "------------------------------------------"
call :LOG "[STEP 7] Verifying upgrade process is running..."
call :LOG "[DEBUG] Waiting 5 seconds for process to start..."
timeout /t 5 /nobreak >nul

psexec \\%PC% -s cmd /c "tasklist | findstr /i Windows11" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    call :LOG "[DEBUG] Windows11 process found running on %PC%."
) else (
    call :LOG "[DEBUG] Windows11 process not found. Checking for setup..."
    psexec \\%PC% -s cmd /c "tasklist | findstr /i setup" >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        call :LOG "[DEBUG] Setup process found running."
    ) else (
        call :LOG "[WARNING] No upgrade process detected yet. May still be initializing."
    )
)

:: Complete
call :LOG "------------------------------------------"
call :LOG "[COMPLETE] Script finished successfully."
call :LOG "=========================================="
call :LOG ""
call :LOG "TO MONITOR PROGRESS:"
call :LOG "  psexec \\%PC% -s powershell \"Get-Content 'C:\$WINDOWS.~BT\Sources\Panther\setupact.log' -Tail 20 -Wait\""
call :LOG ""
call :LOG "TO VERIFY AFTER REBOOT:"
call :LOG "  psexec \\%PC% -s cmd /c \"wmic os get Caption\""
call :LOG "=========================================="

goto :END

:: Logging function
:LOG
echo %~1
echo %~1 >> "%LOGFILE%"
goto :EOF

:END
call :LOG "[DEBUG] Script execution ended at %DATE% %TIME%"
echo.
echo Log saved to: %LOGFILE%
pause
exit /b
