@echo off
echo ==========================================
echo Windows 11 Bypass - Registry Deployment
echo ==========================================
echo.

:: List of all PCs
set "PCLIST=01INVENTORY-PC 03STAGE-PC ALESOVICH-LAP BCHASTEEN-LAP CALLDATA-PC EAGLE5 PSMITH-LAP SECURITY99 SHIP2-PC2 SVANHOLLEN-PC"

for %%P in (%PCLIST%) do (
    echo ------------------------------------------
    echo Processing: %%P
    echo ------------------------------------------

    psexec \\%%P -s reg add "HKLM\SYSTEM\Setup\MoSetup" /v AllowUpgradesWithUnsupportedTPMOrCPU /t REG_DWORD /d 1 /f

    psexec \\%%P -s reg query "HKLM\SYSTEM\Setup\MoSetup" /v AllowUpgradesWithUnsupportedTPMOrCPU

    echo %%P complete.
    echo.
)

echo ==========================================
echo All PCs processed!
echo ==========================================
pause
