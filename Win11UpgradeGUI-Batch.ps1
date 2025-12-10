Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
# CONFIGURATION
# ============================================================
$script:ISOUrl = "https://software-static.download.prss.microsoft.com/dbazure/888969d5-f34g-4e03-ac9d-1f9786c66749/26200.6584.250915-1905.25h2_ge_release_svc_refresh_CLIENT_CONSUMER_x64FRE_en-us.iso"
$script:RemotePath = "C:\Win11Upgrade"
$script:ISOName = "Win11.iso"
$script:MinSpaceGB = 30
$script:AutoReboot = $true
$script:PCStatus = @{}  # Hashtable to track status of each PC

# Status constants
$script:STATUS_PENDING = "Pending"
$script:STATUS_TESTING = "Testing..."
$script:STATUS_ONLINE = "Online"
$script:STATUS_OFFLINE = "Offline"
$script:STATUS_BYPASSING = "Applying Bypass..."
$script:STATUS_BYPASSED = "Bypass Applied"
$script:STATUS_DOWNLOADING = "Downloading..."
$script:STATUS_DOWNLOADED = "ISO Downloaded"
$script:STATUS_EXTRACTING = "Extracting..."
$script:STATUS_EXTRACTED = "Extracted"
$script:STATUS_UPGRADING = "Upgrading..."
$script:STATUS_COMPLETE = "Complete!"
$script:STATUS_REBOOTING = "Rebooting..."
$script:STATUS_FAILED = "FAILED"

# ============================================================
# CREATE MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows 11 Batch Upgrade Tool - Parallel Mode"
$form.Size = New-Object System.Drawing.Size(950, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# ============================================================
# PC LIST INPUT (Left Panel)
# ============================================================
$lblPCList = New-Object System.Windows.Forms.Label
$lblPCList.Location = New-Object System.Drawing.Point(10, 10)
$lblPCList.Size = New-Object System.Drawing.Size(200, 20)
$lblPCList.Text = "Target PCs (one per line):"
$form.Controls.Add($lblPCList)

$txtPCList = New-Object System.Windows.Forms.TextBox
$txtPCList.Location = New-Object System.Drawing.Point(10, 35)
$txtPCList.Size = New-Object System.Drawing.Size(200, 200)
$txtPCList.Multiline = $true
$txtPCList.ScrollBars = "Vertical"
$txtPCList.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtPCList.Text = @"
01INVENTORY-PC
03STAGE-PC
ALESOVICH-LAP
BCHASTEEN-LAP
CALLDATA-PC
PSMITH-LAP
SECURITY99
SHIP2-PC2
SVANHOLLEN-PC
"@
$form.Controls.Add($txtPCList)

# ============================================================
# STATUS LIST (Middle Panel)
# ============================================================
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(220, 10)
$lblStatus.Size = New-Object System.Drawing.Size(200, 20)
$lblStatus.Text = "PC Status:"
$form.Controls.Add($lblStatus)

$listStatus = New-Object System.Windows.Forms.ListView
$listStatus.Location = New-Object System.Drawing.Point(220, 35)
$listStatus.Size = New-Object System.Drawing.Size(350, 200)
$listStatus.View = [System.Windows.Forms.View]::Details
$listStatus.FullRowSelect = $true
$listStatus.GridLines = $true
$listStatus.Font = New-Object System.Drawing.Font("Consolas", 9)
$listStatus.Columns.Add("PC Name", 150) | Out-Null
$listStatus.Columns.Add("Status", 180) | Out-Null
$form.Controls.Add($listStatus)

# ============================================================
# ACTION BUTTONS (Right Panel)
# ============================================================
$btnLoadPCs = New-Object System.Windows.Forms.Button
$btnLoadPCs.Location = New-Object System.Drawing.Point(580, 35)
$btnLoadPCs.Size = New-Object System.Drawing.Size(170, 30)
$btnLoadPCs.Text = "1. Load PC List"
$form.Controls.Add($btnLoadPCs)

$btnTestAll = New-Object System.Windows.Forms.Button
$btnTestAll.Location = New-Object System.Drawing.Point(580, 70)
$btnTestAll.Size = New-Object System.Drawing.Size(170, 30)
$btnTestAll.Text = "2. Test All Connections"
$form.Controls.Add($btnTestAll)

$btnBypassAll = New-Object System.Windows.Forms.Button
$btnBypassAll.Location = New-Object System.Drawing.Point(580, 105)
$btnBypassAll.Size = New-Object System.Drawing.Size(170, 30)
$btnBypassAll.Text = "3. Apply Bypass (All)"
$form.Controls.Add($btnBypassAll)

$btnDownloadAll = New-Object System.Windows.Forms.Button
$btnDownloadAll.Location = New-Object System.Drawing.Point(580, 140)
$btnDownloadAll.Size = New-Object System.Drawing.Size(170, 30)
$btnDownloadAll.Text = "4. Download ISO (All)"
$form.Controls.Add($btnDownloadAll)

$btnExtractAll = New-Object System.Windows.Forms.Button
$btnExtractAll.Location = New-Object System.Drawing.Point(580, 175)
$btnExtractAll.Size = New-Object System.Drawing.Size(170, 30)
$btnExtractAll.Text = "5. Extract ISO (All)"
$form.Controls.Add($btnExtractAll)

$btnUpgradeAll = New-Object System.Windows.Forms.Button
$btnUpgradeAll.Location = New-Object System.Drawing.Point(760, 35)
$btnUpgradeAll.Size = New-Object System.Drawing.Size(170, 30)
$btnUpgradeAll.Text = "6. Start Upgrade (All)"
$btnUpgradeAll.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnUpgradeAll)

$btnMonitorAll = New-Object System.Windows.Forms.Button
$btnMonitorAll.Location = New-Object System.Drawing.Point(760, 70)
$btnMonitorAll.Size = New-Object System.Drawing.Size(170, 30)
$btnMonitorAll.Text = "7. Monitor All"
$form.Controls.Add($btnMonitorAll)

$btnRebootAll = New-Object System.Windows.Forms.Button
$btnRebootAll.Location = New-Object System.Drawing.Point(760, 105)
$btnRebootAll.Size = New-Object System.Drawing.Size(170, 30)
$btnRebootAll.Text = "Reboot All Ready"
$btnRebootAll.BackColor = [System.Drawing.Color]::OrangeRed
$btnRebootAll.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnRebootAll)

$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Location = New-Object System.Drawing.Point(760, 140)
$btnClearLog.Size = New-Object System.Drawing.Size(170, 30)
$btnClearLog.Text = "Clear Log"
$form.Controls.Add($btnClearLog)

# Verify OS button
$btnVerifyAll = New-Object System.Windows.Forms.Button
$btnVerifyAll.Location = New-Object System.Drawing.Point(760, 175)
$btnVerifyAll.Size = New-Object System.Drawing.Size(170, 30)
$btnVerifyAll.Text = "8. Verify OS Version"
$form.Controls.Add($btnVerifyAll)

# Auto-reboot checkbox
$chkAutoReboot = New-Object System.Windows.Forms.CheckBox
$chkAutoReboot.Location = New-Object System.Drawing.Point(760, 210)
$chkAutoReboot.Size = New-Object System.Drawing.Size(170, 25)
$chkAutoReboot.Text = "Auto-reboot when ready"
$chkAutoReboot.Checked = $script:AutoReboot
$chkAutoReboot.Add_CheckedChanged({ $script:AutoReboot = $chkAutoReboot.Checked })
$form.Controls.Add($chkAutoReboot)

# ============================================================
# LOG OUTPUT
# ============================================================
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(10, 245)
$lblLog.Size = New-Object System.Drawing.Size(100, 20)
$lblLog.Text = "Debug Log:"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 270)
$txtLog.Size = New-Object System.Drawing.Size(920, 380)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = [System.Drawing.Color]::Black
$txtLog.ForeColor = [System.Drawing.Color]::LightGreen
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

# ============================================================
# HELPER FUNCTIONS
# ============================================================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    $txtLog.AppendText("$logEntry`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-PCStatus {
    param([string]$PCName, [string]$Status)
    $script:PCStatus[$PCName] = $Status

    # Update ListView
    foreach ($item in $listStatus.Items) {
        if ($item.Text -eq $PCName) {
            $item.SubItems[1].Text = $Status

            # Color coding
            switch -Wildcard ($Status) {
                "*FAILED*" { $item.BackColor = [System.Drawing.Color]::LightCoral }
                "*Complete*" { $item.BackColor = [System.Drawing.Color]::LightGreen }
                "*Upgrading*" { $item.BackColor = [System.Drawing.Color]::LightBlue }
                "*Downloading*" { $item.BackColor = [System.Drawing.Color]::LightYellow }
                "*Offline*" { $item.BackColor = [System.Drawing.Color]::LightGray }
                default { $item.BackColor = [System.Drawing.Color]::White }
            }
            break
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-PCList {
    $pcs = $txtPCList.Text -split "`r`n|`n" | Where-Object { $_.Trim() -ne "" } | ForEach-Object { $_.Trim() }
    return $pcs
}

# ============================================================
# AUTO-DETECT PC STATE FUNCTION
# ============================================================
function Detect-PCState {
    param([string]$PCName)

    # Check if PC is reachable first
    if (-not (Test-Connection -ComputerName $PCName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        return $script:STATUS_OFFLINE
    }

    # Check if admin share accessible
    if (-not (Test-Path "\\$PCName\C$" -ErrorAction SilentlyContinue)) {
        return "Online (No Admin)"
    }

    $remotePath = "\\$PCName\C$\Win11Upgrade"
    $isoPath = "$remotePath\Win11.iso"
    $extractedPath = "$remotePath\Extracted\sources\setupprep.exe"
    $stagingPath = "\\$PCName\C$\`$WINDOWS.~BT"

    # Check if setup process is running (upgrading)
    $setupRunning = $false
    try {
        $taskCheck = & psexec "\\$PCName" -s cmd /c "tasklist | findstr /i SetupHost" 2>&1
        if ($taskCheck -match "SetupHost") {
            $setupRunning = $true
        }
    } catch {}

    # Only consider "Upgrading" if BOTH setup is running AND we have our extracted files
    if ($setupRunning -and (Test-Path $extractedPath)) {
        return $script:STATUS_UPGRADING
    }

    # Check if staging folder exists with our upgrade (check for our specific install path in logs)
    if ((Test-Path $stagingPath) -and (Test-Path $extractedPath)) {
        $setupLog = "$stagingPath\Sources\Panther\setupact.log"
        if (Test-Path $setupLog) {
            $logTail = Get-Content $setupLog -Tail 50 -ErrorAction SilentlyContinue
            $logText = $logTail -join " "

            # Verify it's OUR upgrade by checking for our path
            if ($logText -match "Win11Upgrade") {
                if ($logText -match "RebootMachine" -or $logText -match "Reboot is required" -or $logText -match "Overall progress: \[100%\]") {
                    return $script:STATUS_COMPLETE
                }
                # Our upgrade in progress but setup not running = stalled or complete
                if (-not $setupRunning) {
                    return $script:STATUS_COMPLETE
                }
                return $script:STATUS_UPGRADING
            }
        }
    }

    # Check if extracted
    if (Test-Path $extractedPath) {
        return $script:STATUS_EXTRACTED
    }

    # Check if ISO downloaded
    if (Test-Path $isoPath) {
        $size = (Get-Item $isoPath -ErrorAction SilentlyContinue).Length / 1GB
        if ($size -gt 5) {
            return $script:STATUS_DOWNLOADED
        } else {
            return $script:STATUS_DOWNLOADING
        }
    }

    # Check if directory exists (bypass may have been applied)
    if (Test-Path $remotePath) {
        return $script:STATUS_BYPASSED
    }

    # Just online
    return $script:STATUS_ONLINE
}

# ============================================================
# 1. LOAD PC LIST (with auto-detect)
# ============================================================
$btnLoadPCs.Add_Click({
    $listStatus.Items.Clear()
    $script:PCStatus.Clear()

    $pcs = Get-PCList

    Write-Log "Loading $($pcs.Count) PCs and detecting current state..." "INFO"

    foreach ($pc in $pcs) {
        $item = New-Object System.Windows.Forms.ListViewItem($pc)
        $item.SubItems.Add("Detecting...")
        $listStatus.Items.Add($item) | Out-Null
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Detect state for each PC
    foreach ($pc in $pcs) {
        Write-Log "[DETECT] Checking $pc..." "DEBUG"
        [System.Windows.Forms.Application]::DoEvents()

        $state = Detect-PCState $pc
        Update-PCStatus $pc $state
        Write-Log "[DETECT] $pc = $state" "INFO"
    }

    Write-Log "========================================" "INFO"
    Write-Log "Auto-detection complete. States:" "INFO"

    $summary = $script:PCStatus.Values | Group-Object | ForEach-Object { "$($_.Name): $($_.Count)" }
    Write-Log ($summary -join ", ") "INFO"
})

# ============================================================
# 2. TEST ALL CONNECTIONS (Parallel)
# ============================================================
$btnTestAll.Add_Click({
    $pcs = Get-PCList
    Write-Log "========================================" "INFO"
    Write-Log "Testing connections to $($pcs.Count) PCs in parallel..." "INFO"

    foreach ($pc in $pcs) {
        Update-PCStatus $pc $script:STATUS_TESTING
    }

    # Test all in parallel using jobs
    $jobs = @()
    foreach ($pc in $pcs) {
        $jobs += Start-Job -ScriptBlock {
            param($pcName)
            $result = @{ PC = $pcName; Online = $false; AdminShare = $false }

            if (Test-Connection -ComputerName $pcName -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                $result.Online = $true
                if (Test-Path "\\$pcName\C$") {
                    $result.AdminShare = $true
                }
            }
            return $result
        } -ArgumentList $pc
    }

    # Wait and collect results
    $results = $jobs | Wait-Job | Receive-Job
    $jobs | Remove-Job

    $onlineCount = 0
    foreach ($result in $results) {
        if ($result.Online -and $result.AdminShare) {
            Update-PCStatus $result.PC $script:STATUS_ONLINE
            Write-Log "[OK] $($result.PC) - Online and accessible" "INFO"
            $onlineCount++
        } elseif ($result.Online) {
            Update-PCStatus $result.PC "Online (No Admin Share)"
            Write-Log "[WARN] $($result.PC) - Online but admin share not accessible" "WARN"
        } else {
            Update-PCStatus $result.PC $script:STATUS_OFFLINE
            Write-Log "[FAIL] $($result.PC) - Offline or unreachable" "ERROR"
        }
    }

    Write-Log "Connection test complete: $onlineCount/$($pcs.Count) PCs online and accessible" "INFO"
})

# ============================================================
# 3. APPLY BYPASS TO ALL (Parallel)
# ============================================================
$btnBypassAll.Add_Click({
    $pcs = Get-PCList | Where-Object { $script:PCStatus[$_] -eq $script:STATUS_ONLINE -or $script:PCStatus[$_] -eq $script:STATUS_BYPASSED }

    if ($pcs.Count -eq 0) {
        Write-Log "[ERROR] No online PCs to apply bypass. Run 'Test All' first." "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Applying bypass to $($pcs.Count) PCs in parallel..." "INFO"

    foreach ($pc in $pcs) {
        Update-PCStatus $pc $script:STATUS_BYPASSING
    }

    $jobs = @()
    foreach ($pc in $pcs) {
        $jobs += Start-Job -ScriptBlock {
            param($pcName)
            $result = @{ PC = $pcName; Success = $false; Error = "" }

            try {
                # Apply registry keys via PsExec
                $regKeys = @(
                    "HKLM\SYSTEM\Setup\MoSetup /v AllowUpgradesWithUnsupportedTPMOrCPU /t REG_DWORD /d 1 /f",
                    "HKLM\SYSTEM\Setup\LabConfig /v BypassTPMCheck /t REG_DWORD /d 1 /f",
                    "HKLM\SYSTEM\Setup\LabConfig /v BypassSecureBootCheck /t REG_DWORD /d 1 /f",
                    "HKLM\SYSTEM\Setup\LabConfig /v BypassRAMCheck /t REG_DWORD /d 1 /f",
                    "HKLM\SYSTEM\Setup\LabConfig /v BypassCPUCheck /t REG_DWORD /d 1 /f",
                    "HKLM\SYSTEM\Setup\LabConfig /v BypassStorageCheck /t REG_DWORD /d 1 /f",
                    "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v DisableWUfBSafeguards /t REG_DWORD /d 1 /f"
                )

                foreach ($key in $regKeys) {
                    & psexec "\\$pcName" -s reg add $key.Split(" ") 2>&1 | Out-Null
                }

                $result.Success = $true
            } catch {
                $result.Error = $_.Exception.Message
            }

            return $result
        } -ArgumentList $pc
    }

    $results = $jobs | Wait-Job | Receive-Job
    $jobs | Remove-Job

    $successCount = 0
    foreach ($result in $results) {
        if ($result.Success) {
            Update-PCStatus $result.PC $script:STATUS_BYPASSED
            Write-Log "[OK] $($result.PC) - Bypass applied" "INFO"
            $successCount++
        } else {
            Update-PCStatus $result.PC $script:STATUS_FAILED
            Write-Log "[FAIL] $($result.PC) - $($result.Error)" "ERROR"
        }
    }

    Write-Log "Bypass complete: $successCount/$($pcs.Count) PCs" "INFO"
})

# ============================================================
# 4. DOWNLOAD ISO TO ALL (Parallel)
# ============================================================
$btnDownloadAll.Add_Click({
    $pcs = Get-PCList | Where-Object {
        $script:PCStatus[$_] -eq $script:STATUS_BYPASSED -or
        $script:PCStatus[$_] -eq $script:STATUS_ONLINE -or
        $script:PCStatus[$_] -eq $script:STATUS_DOWNLOADED
    }

    if ($pcs.Count -eq 0) {
        Write-Log "[ERROR] No ready PCs. Run 'Test All' and 'Apply Bypass' first." "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Starting ISO download on $($pcs.Count) PCs..." "INFO"
    Write-Log "[INFO] Downloads run in background on each PC. Monitor file sizes." "INFO"

    foreach ($pc in $pcs) {
        Update-PCStatus $pc $script:STATUS_DOWNLOADING

        # Create remote directory
        $remoteDirUNC = "\\$pc\C$\Win11Upgrade"
        if (!(Test-Path $remoteDirUNC)) {
            New-Item -Path $remoteDirUNC -ItemType Directory -Force | Out-Null
        }

        # Check if already downloaded
        $isoPath = "$remoteDirUNC\$($script:ISOName)"
        if (Test-Path $isoPath) {
            $size = [math]::Round((Get-Item $isoPath).Length / 1GB, 2)
            if ($size -gt 4) {
                Write-Log "[INFO] $pc - ISO already exists ($size GB)" "INFO"
                Update-PCStatus $pc $script:STATUS_DOWNLOADED
                continue
            }
        }

        # Create download script
        $downloadScript = @"
`$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri '$($script:ISOUrl)' -OutFile 'C:\Win11Upgrade\Win11.iso' -UseBasicParsing
"@
        $scriptPath = "$remoteDirUNC\DownloadISO.ps1"
        $downloadScript | Out-File -FilePath $scriptPath -Encoding UTF8 -Force

        # Start download in background
        $cmd = "psexec \\$pc -s -d powershell -ExecutionPolicy Bypass -File `"C:\Win11Upgrade\DownloadISO.ps1`""
        cmd /c $cmd 2>&1 | Out-Null

        Write-Log "[STARTED] $pc - Download started in background" "INFO"
    }

    Write-Log "[INFO] Use 'Monitor All' to check download progress" "INFO"
})

# ============================================================
# 5. EXTRACT ISO ON ALL (Parallel)
# ============================================================
$btnExtractAll.Add_Click({
    $pcs = Get-PCList | Where-Object {
        $script:PCStatus[$_] -eq $script:STATUS_DOWNLOADED -or
        $script:PCStatus[$_] -eq $script:STATUS_EXTRACTED
    }

    if ($pcs.Count -eq 0) {
        Write-Log "[ERROR] No PCs with downloaded ISO. Run 'Download ISO' first." "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Extracting ISO on $($pcs.Count) PCs..." "INFO"

    foreach ($pc in $pcs) {
        Update-PCStatus $pc $script:STATUS_EXTRACTING
    }

    $jobs = @()
    foreach ($pc in $pcs) {
        $jobs += Start-Job -ScriptBlock {
            param($pcName)
            $result = @{ PC = $pcName; Success = $false; Error = "" }

            try {
                $extractScript = @'
$isoFile = 'C:\Win11Upgrade\Win11.iso'
$extractPath = 'C:\Win11Upgrade\Extracted'
if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
$mount = Mount-DiskImage -ImagePath $isoFile -PassThru
$drive = ($mount | Get-Volume).DriveLetter
Copy-Item -Path "${drive}:\*" -Destination $extractPath -Recurse -Force
Dismount-DiskImage -ImagePath $isoFile
'@
                $scriptPath = "\\$pcName\C$\Win11Upgrade\ExtractISO.ps1"
                $extractScript | Out-File -FilePath $scriptPath -Encoding UTF8 -Force

                & psexec "\\$pcName" -s powershell -ExecutionPolicy Bypass -File "C:\Win11Upgrade\ExtractISO.ps1" 2>&1 | Out-Null

                # Verify
                if (Test-Path "\\$pcName\C$\Win11Upgrade\Extracted\setup.exe") {
                    $result.Success = $true
                } else {
                    $result.Error = "setup.exe not found after extraction"
                }
            } catch {
                $result.Error = $_.Exception.Message
            }

            return $result
        } -ArgumentList $pc
    }

    $results = $jobs | Wait-Job | Receive-Job
    $jobs | Remove-Job

    foreach ($result in $results) {
        if ($result.Success) {
            Update-PCStatus $result.PC $script:STATUS_EXTRACTED
            Write-Log "[OK] $($result.PC) - Extraction complete" "INFO"
        } else {
            Update-PCStatus $result.PC $script:STATUS_FAILED
            Write-Log "[FAIL] $($result.PC) - $($result.Error)" "ERROR"
        }
    }
})

# ============================================================
# 6. START UPGRADE ON ALL (Parallel)
# ============================================================
$btnUpgradeAll.Add_Click({
    $pcs = Get-PCList | Where-Object { $script:PCStatus[$_] -eq $script:STATUS_EXTRACTED }

    if ($pcs.Count -eq 0) {
        Write-Log "[ERROR] No PCs ready for upgrade. Run previous steps first." "ERROR"
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Start Windows 11 upgrade on $($pcs.Count) PCs?`n`nPCs: $($pcs -join ', ')",
        "Confirm Batch Upgrade",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log "[INFO] Upgrade cancelled by user" "INFO"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Starting upgrade on $($pcs.Count) PCs..." "INFO"

    foreach ($pc in $pcs) {
        Update-PCStatus $pc $script:STATUS_UPGRADING

        $cmd = "psexec \\$pc -s -d `"C:\Win11Upgrade\Extracted\sources\setupprep.exe`" /product server /auto upgrade /quiet /eula accept /dynamicupdate disable"
        cmd /c $cmd 2>&1 | Out-Null

        Write-Log "[STARTED] $pc - Upgrade process launched" "INFO"
    }

    Write-Log "[INFO] All upgrades started. Use 'Monitor All' to track progress." "INFO"
})

# ============================================================
# 7. MONITOR ALL
# ============================================================
$btnMonitorAll.Add_Click({
    $pcs = Get-PCList

    Write-Log "========================================" "INFO"
    Write-Log "Checking status of all PCs..." "INFO"

    foreach ($pc in $pcs) {
        [System.Windows.Forms.Application]::DoEvents()

        $remoteDirUNC = "\\$pc\C$\Win11Upgrade"
        $isoPath = "$remoteDirUNC\Win11.iso"
        $setupPath = "$remoteDirUNC\Extracted\sources\setupprep.exe"
        $stagingPath = "\\$pc\C$\`$WINDOWS.~BT"

        # Check current state
        $currentStatus = $script:PCStatus[$pc]

        # If downloading, check ISO size
        if ($currentStatus -eq $script:STATUS_DOWNLOADING) {
            if (Test-Path $isoPath) {
                $size = [math]::Round((Get-Item $isoPath).Length / 1GB, 2)
                if ($size -gt 5.5) {
                    Update-PCStatus $pc $script:STATUS_DOWNLOADED
                    Write-Log "[OK] $pc - Download complete ($size GB)" "INFO"
                } else {
                    Write-Log "[PROGRESS] $pc - Downloading: $size GB" "DEBUG"
                }
            }
        }

        # If upgrading, check progress
        if ($currentStatus -eq $script:STATUS_UPGRADING) {
            $setupLog = "$stagingPath\Sources\Panther\setupact.log"
            if (Test-Path $setupLog) {
                $logTail = Get-Content $setupLog -Tail 5 -ErrorAction SilentlyContinue
                $logText = $logTail -join " "

                # Extract progress
                if ($logText -match "Overall progress: \[(\d+)%\]") {
                    $progress = $matches[1]
                    Write-Log "[PROGRESS] $pc - Upgrade: $progress%" "INFO"

                    if ([int]$progress -ge 100) {
                        Update-PCStatus $pc $script:STATUS_COMPLETE
                        Write-Log "[COMPLETE] $pc - Upgrade finished! Reboot required." "INFO"

                        if ($script:AutoReboot) {
                            Write-Log "[INFO] $pc - Auto-rebooting..." "INFO"
                            & psexec "\\$pc" -s shutdown /r /f /t 30 /c "Windows 11 upgrade complete" 2>&1 | Out-Null
                            Update-PCStatus $pc $script:STATUS_REBOOTING
                        }
                    }
                }

                # Check for completion markers
                if ($logText -match "RebootMachine" -or $logText -match "Reboot is required") {
                    Update-PCStatus $pc $script:STATUS_COMPLETE
                    Write-Log "[COMPLETE] $pc - Ready for reboot" "INFO"
                }
            }

            # Check if setup process is still running
            $taskCheck = & psexec "\\$pc" -s cmd /c "tasklist | findstr /i setup" 2>&1
            if (-not ($taskCheck -match "setup")) {
                if (Test-Path $stagingPath) {
                    Update-PCStatus $pc $script:STATUS_COMPLETE
                    Write-Log "[COMPLETE] $pc - Upgrade finished (setup process ended)" "INFO"
                }
            }
        }
    }

    Write-Log "Status check complete." "INFO"
})

# ============================================================
# REBOOT ALL READY PCs
# ============================================================
$btnRebootAll.Add_Click({
    $pcs = Get-PCList | Where-Object { $script:PCStatus[$_] -eq $script:STATUS_COMPLETE }

    if ($pcs.Count -eq 0) {
        Write-Log "[INFO] No PCs ready for reboot" "INFO"
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Reboot $($pcs.Count) PCs now?`n`nPCs: $($pcs -join ', ')",
        "Confirm Batch Reboot",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Rebooting $($pcs.Count) PCs..." "INFO"

    foreach ($pc in $pcs) {
        & psexec "\\$pc" -s shutdown /r /f /t 10 /c "Windows 11 upgrade - rebooting" 2>&1 | Out-Null
        Update-PCStatus $pc $script:STATUS_REBOOTING
        Write-Log "[REBOOT] $pc - Reboot command sent" "INFO"
    }
})

# ============================================================
# CLEAR LOG
# ============================================================
$btnClearLog.Add_Click({
    $txtLog.Clear()
    Write-Log "Log cleared." "INFO"
})

# ============================================================
# 8. VERIFY OS VERSION (All PCs)
# ============================================================
$btnVerifyAll.Add_Click({
    $pcs = Get-PCList

    Write-Log "========================================" "INFO"
    Write-Log "Checking OS version on all PCs..." "INFO"

    $win11Count = 0
    $win10Count = 0
    $offlineCount = 0

    foreach ($pc in $pcs) {
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $os = Get-WmiObject Win32_OperatingSystem -ComputerName $pc -ErrorAction Stop

            if ($os.Caption -like "*Windows 11*") {
                Write-Log "[WIN11] $pc - $($os.Caption) (Build $($os.BuildNumber))" "INFO"
                Update-PCStatus $pc "Windows 11"
                $win11Count++
            } else {
                Write-Log "[WIN10] $pc - $($os.Caption) (Build $($os.BuildNumber))" "INFO"
                Update-PCStatus $pc "Windows 10"
                $win10Count++
            }
        } catch {
            Write-Log "[OFFLINE] $pc - Cannot query OS" "WARN"
            Update-PCStatus $pc $script:STATUS_OFFLINE
            $offlineCount++
        }
    }

    Write-Log "========================================" "INFO"
    Write-Log "Summary: Windows 11: $win11Count | Windows 10: $win10Count | Offline: $offlineCount" "INFO"
})

# ============================================================
# STARTUP
# ============================================================
Write-Log "Windows 11 Batch Upgrade Tool - Parallel Mode" "INFO"
Write-Log "========================================" "INFO"
Write-Log "Instructions:" "INFO"
Write-Log "  1. Edit PC list (left panel) or keep defaults" "INFO"
Write-Log "  2. Click 'Load PC List' to populate status grid" "INFO"
Write-Log "  3. Click 'Test All' to verify connectivity" "INFO"
Write-Log "  4. Run steps 3-6 in order for parallel upgrade" "INFO"
Write-Log "  5. Use 'Monitor All' to check progress" "INFO"
Write-Log "========================================" "INFO"

# Show form
[void]$form.ShowDialog()
