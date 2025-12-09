Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
# CONFIGURATION
# ============================================================
$script:TargetPC = ""
$script:ISOUrl = "https://software-static.download.prss.microsoft.com/dbazure/888969d5-f34g-4e03-ac9d-1f9786c66749/26200.6584.250915-1905.25h2_ge_release_svc_refresh_CLIENT_CONSUMER_x64FRE_en-us.iso"
$script:RemotePath = "C:\Win11Upgrade"
$script:ISOName = "Win11.iso"
$script:MinSpaceGB = 30
$script:AutoReboot = $true

# ============================================================
# CREATE MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows 11 Remote Upgrade Tool"
$form.Size = New-Object System.Drawing.Size(800, 650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# ============================================================
# TARGET PC INPUT
# ============================================================
$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Location = New-Object System.Drawing.Point(10, 15)
$lblTarget.Size = New-Object System.Drawing.Size(80, 20)
$lblTarget.Text = "Target PC:"
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(95, 12)
$txtTarget.Size = New-Object System.Drawing.Size(200, 20)
$txtTarget.Text = "01INVENTORY-PC"
$form.Controls.Add($txtTarget)

$btnSetTarget = New-Object System.Windows.Forms.Button
$btnSetTarget.Location = New-Object System.Drawing.Point(305, 10)
$btnSetTarget.Size = New-Object System.Drawing.Size(80, 25)
$btnSetTarget.Text = "Set Target"
$form.Controls.Add($btnSetTarget)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(400, 15)
$lblStatus.Size = New-Object System.Drawing.Size(350, 20)
$lblStatus.Text = "Status: No target set"
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblStatus)

# ============================================================
# BUTTONS PANEL
# ============================================================
$btnPing = New-Object System.Windows.Forms.Button
$btnPing.Location = New-Object System.Drawing.Point(10, 50)
$btnPing.Size = New-Object System.Drawing.Size(150, 35)
$btnPing.Text = "1. Test Connection"
$form.Controls.Add($btnPing)

$btnStorage = New-Object System.Windows.Forms.Button
$btnStorage.Location = New-Object System.Drawing.Point(170, 50)
$btnStorage.Size = New-Object System.Drawing.Size(150, 35)
$btnStorage.Text = "2. Check Storage"
$form.Controls.Add($btnStorage)

$btnRegistry = New-Object System.Windows.Forms.Button
$btnRegistry.Location = New-Object System.Drawing.Point(330, 50)
$btnRegistry.Size = New-Object System.Drawing.Size(150, 35)
$btnRegistry.Text = "3. Apply Bypass"
$form.Controls.Add($btnRegistry)

$btnDownload = New-Object System.Windows.Forms.Button
$btnDownload.Location = New-Object System.Drawing.Point(490, 50)
$btnDownload.Size = New-Object System.Drawing.Size(150, 35)
$btnDownload.Text = "4. Download ISO"
$form.Controls.Add($btnDownload)

$btnExtract = New-Object System.Windows.Forms.Button
$btnExtract.Location = New-Object System.Drawing.Point(650, 50)
$btnExtract.Size = New-Object System.Drawing.Size(130, 35)
$btnExtract.Text = "5. Extract ISO"
$form.Controls.Add($btnExtract)

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Location = New-Object System.Drawing.Point(10, 95)
$btnInstall.Size = New-Object System.Drawing.Size(150, 35)
$btnInstall.Text = "6. Start Upgrade"
$btnInstall.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnInstall)

$btnMonitor = New-Object System.Windows.Forms.Button
$btnMonitor.Location = New-Object System.Drawing.Point(170, 95)
$btnMonitor.Size = New-Object System.Drawing.Size(150, 35)
$btnMonitor.Text = "7. Monitor Progress"
$form.Controls.Add($btnMonitor)

$btnVerify = New-Object System.Windows.Forms.Button
$btnVerify.Location = New-Object System.Drawing.Point(330, 95)
$btnVerify.Size = New-Object System.Drawing.Size(150, 35)
$btnVerify.Text = "8. Verify OS Version"
$form.Controls.Add($btnVerify)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(650, 95)
$btnClear.Size = New-Object System.Drawing.Size(130, 35)
$btnClear.Text = "Clear Log"
$form.Controls.Add($btnClear)

# Auto-reboot checkbox
$chkAutoReboot = New-Object System.Windows.Forms.CheckBox
$chkAutoReboot.Location = New-Object System.Drawing.Point(490, 100)
$chkAutoReboot.Size = New-Object System.Drawing.Size(150, 25)
$chkAutoReboot.Text = "Auto-reboot when ready"
$chkAutoReboot.Checked = $script:AutoReboot
$chkAutoReboot.Add_CheckedChanged({ $script:AutoReboot = $chkAutoReboot.Checked })
$form.Controls.Add($chkAutoReboot)

# Force Reboot button
$btnReboot = New-Object System.Windows.Forms.Button
$btnReboot.Location = New-Object System.Drawing.Point(10, 130)
$btnReboot.Size = New-Object System.Drawing.Size(120, 30)
$btnReboot.Text = "Force Reboot"
$btnReboot.BackColor = [System.Drawing.Color]::OrangeRed
$btnReboot.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnReboot)

# ============================================================
# LOG OUTPUT
# ============================================================
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(140, 135)
$lblLog.Size = New-Object System.Drawing.Size(100, 20)
$lblLog.Text = "Debug Log:"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 165)
$txtLog.Size = New-Object System.Drawing.Size(765, 435)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = [System.Drawing.Color]::Black
$txtLog.ForeColor = [System.Drawing.Color]::LightGreen
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

# ============================================================
# LOGGING FUNCTION
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

# ============================================================
# SET TARGET BUTTON
# ============================================================
$btnSetTarget.Add_Click({
    $script:TargetPC = $txtTarget.Text.Trim()
    if ($script:TargetPC -eq "") {
        $lblStatus.Text = "Status: No target set"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "ERROR: No target PC specified" "ERROR"
    } else {
        $lblStatus.Text = "Status: Target = $($script:TargetPC)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Blue
        Write-Log "Target PC set to: $($script:TargetPC)" "INFO"
    }
})

# ============================================================
# 1. TEST CONNECTION
# ============================================================
$btnPing.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Testing connection to $($script:TargetPC)..." "INFO"

    # Ping test
    Write-Log "[DEBUG] Running ping test..." "DEBUG"
    $ping = Test-Connection -ComputerName $script:TargetPC -Count 1 -Quiet

    if ($ping) {
        Write-Log "[OK] Ping successful - PC is reachable" "INFO"

        # Admin share test
        Write-Log "[DEBUG] Testing admin share access..." "DEBUG"
        $sharePath = "\\$($script:TargetPC)\C$"
        if (Test-Path $sharePath) {
            Write-Log "[OK] Admin share accessible: $sharePath" "INFO"
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            Write-Log "[FAIL] Cannot access admin share: $sharePath" "ERROR"
            Write-Log "[SOLUTION] Check admin rights and firewall settings" "INFO"
        }
    } else {
        Write-Log "[FAIL] Ping failed - PC is not reachable" "ERROR"
        Write-Log "[SOLUTION] Check if PC is online and network connectivity" "INFO"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }

    Write-Log "Connection test complete." "INFO"
})

# ============================================================
# 2. CHECK STORAGE
# ============================================================
$btnStorage.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Checking storage on $($script:TargetPC)..." "INFO"

    try {
        Write-Log "[DEBUG] Querying disk space via WMI..." "DEBUG"
        $disk = Get-WmiObject Win32_LogicalDisk -ComputerName $script:TargetPC -Filter "DeviceID='C:'" -ErrorAction Stop

        $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        $totalGB = [math]::Round($disk.Size / 1GB, 2)
        $usedGB = $totalGB - $freeGB
        $percentFree = [math]::Round(($freeGB / $totalGB) * 100, 1)

        Write-Log "[INFO] Total Space: $totalGB GB" "INFO"
        Write-Log "[INFO] Used Space: $usedGB GB" "INFO"
        Write-Log "[INFO] Free Space: $freeGB GB ($percentFree%)" "INFO"

        if ($freeGB -ge $script:MinSpaceGB) {
            Write-Log "[OK] Sufficient space available (need $($script:MinSpaceGB) GB, have $freeGB GB)" "INFO"
        } else {
            Write-Log "[FAIL] Insufficient space! Need $($script:MinSpaceGB) GB, only have $freeGB GB" "ERROR"
            Write-Log "[SOLUTION] Free up at least $([math]::Ceiling($script:MinSpaceGB - $freeGB)) GB on C: drive" "INFO"
        }
    } catch {
        Write-Log "[ERROR] Failed to query storage: $($_.Exception.Message)" "ERROR"
    }

    Write-Log "Storage check complete." "INFO"
})

# ============================================================
# 3. APPLY BYPASS (Registry + appraiserres.dll)
# ============================================================
$btnRegistry.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Applying ALL bypasses on $($script:TargetPC)..." "INFO"

    # PART 1: Registry keys
    Write-Log "[STEP 1/2] Applying registry bypasses..." "INFO"

    $regKeys = @(
        @{Path="HKLM:\SYSTEM\Setup\MoSetup"; Name="AllowUpgradesWithUnsupportedTPMOrCPU"; Value=1},
        @{Path="HKLM:\SYSTEM\Setup\LabConfig"; Name="BypassTPMCheck"; Value=1},
        @{Path="HKLM:\SYSTEM\Setup\LabConfig"; Name="BypassSecureBootCheck"; Value=1},
        @{Path="HKLM:\SYSTEM\Setup\LabConfig"; Name="BypassRAMCheck"; Value=1},
        @{Path="HKLM:\SYSTEM\Setup\LabConfig"; Name="BypassCPUCheck"; Value=1},
        @{Path="HKLM:\SYSTEM\Setup\LabConfig"; Name="BypassStorageCheck"; Value=1},
        @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"; Name="DisableWUfBSafeguards"; Value=1}
    )

    foreach ($key in $regKeys) {
        Write-Log "[DEBUG] Setting $($key.Path)\$($key.Name) = $($key.Value)" "DEBUG"

        try {
            $result = Invoke-Command -ComputerName $script:TargetPC -ScriptBlock {
                param($path, $name, $value)

                # Create path if not exists
                $parentPath = Split-Path $path
                $leafPath = Split-Path $path -Leaf
                if (!(Test-Path $path)) {
                    New-Item -Path $parentPath -Name $leafPath -Force | Out-Null
                }

                New-ItemProperty -Path $path -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
                return "OK"
            } -ArgumentList $key.Path, $key.Name, $key.Value -ErrorAction Stop

            Write-Log "[OK] $($key.Name) set successfully" "INFO"
        } catch {
            Write-Log "[WARN] PowerShell remoting failed, trying PsExec..." "WARN"

            # Fallback to psexec
            $regPath = $key.Path -replace "HKLM:", "HKLM"
            $cmd = "psexec \\$($script:TargetPC) -s reg add `"$regPath`" /v $($key.Name) /t REG_DWORD /d $($key.Value) /f 2>&1"
            Write-Log "[DEBUG] Running: $cmd" "DEBUG"

            $output = cmd /c $cmd
            Write-Log "[DEBUG] Output: $output" "DEBUG"
        }
    }

    Write-Log "[OK] Registry bypasses complete." "INFO"

    # PART 2: Disable appraiserres.dll (hardware check DLL)
    Write-Log "[STEP 2/2] Disabling hardware check DLL (appraiserres.dll)..." "INFO"

    $remoteDirUNC = "\\$($script:TargetPC)\C$\Win11Upgrade"
    $appraiserPath = "$remoteDirUNC\Extracted\sources\appraiserres.dll"
    $appraiserBackup = "$remoteDirUNC\Extracted\sources\appraiserres.dll.bak"

    if (Test-Path $appraiserPath) {
        try {
            # Rename to disable it
            if (Test-Path $appraiserBackup) {
                Remove-Item $appraiserBackup -Force
                Write-Log "[DEBUG] Removed old backup file" "DEBUG"
            }
            Rename-Item -Path $appraiserPath -NewName "appraiserres.dll.bak" -Force
            Write-Log "[OK] appraiserres.dll disabled (renamed to .bak)" "INFO"
            Write-Log "[INFO] This forces setup to skip ALL hardware compatibility checks" "INFO"
        } catch {
            Write-Log "[ERROR] Failed to rename appraiserres.dll: $($_.Exception.Message)" "ERROR"
            Write-Log "[DEBUG] Trying via PsExec..." "DEBUG"

            $cmd = "psexec \\$($script:TargetPC) -s cmd /c `"ren C:\Win11Upgrade\Extracted\sources\appraiserres.dll appraiserres.dll.bak`" 2>&1"
            $output = cmd /c $cmd
            Write-Log "[DEBUG] Output: $output" "DEBUG"
        }
    } else {
        Write-Log "[INFO] appraiserres.dll not found - run Extract ISO first, then Apply Bypass again" "INFO"
        Write-Log "[INFO] Registry bypasses have been applied. Run this step again after extraction." "INFO"
    }

    Write-Log "========================================" "INFO"
    Write-Log "All bypasses applied!" "INFO"
    Write-Log "[INFO] You can now run Start Upgrade" "INFO"
})

# ============================================================
# 4. DOWNLOAD ISO
# ============================================================
$script:ExpectedISOSizeGB = 6.6
$script:DownloadCancelled = $false

$btnDownload.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    $script:DownloadCancelled = $false

    Write-Log "========================================" "INFO"
    Write-Log "Downloading Windows 11 ISO to $($script:TargetPC)..." "INFO"
    Write-Log "[WARN] Note: Microsoft direct links expire. If download fails, you may need to manually get ISO URL." "WARN"

    # Create remote directory
    Write-Log "[DEBUG] Creating remote directory $($script:RemotePath)..." "DEBUG"
    $remoteDirUNC = "\\$($script:TargetPC)\C$\Win11Upgrade"

    if (!(Test-Path $remoteDirUNC)) {
        New-Item -Path $remoteDirUNC -ItemType Directory -Force | Out-Null
        Write-Log "[OK] Created directory: $($script:RemotePath)" "INFO"
    } else {
        Write-Log "[INFO] Directory already exists" "INFO"
    }

    # Check if ISO already exists
    $isoPath = "$remoteDirUNC\$($script:ISOName)"
    if (Test-Path $isoPath) {
        $size = [math]::Round((Get-Item $isoPath).Length / 1GB, 2)
        if ($size -gt 4) {
            Write-Log "[INFO] ISO already exists ($size GB). Delete it first to re-download." "INFO"
            return
        } else {
            Write-Log "[INFO] Partial download found ($size GB). Removing and restarting..." "INFO"
            Remove-Item $isoPath -Force
        }
    }

    Write-Log "[INFO] Starting download (this may take 30-60 minutes)..." "INFO"
    Write-Log "[INFO] Expected size: ~$($script:ExpectedISOSizeGB) GB" "INFO"
    Write-Log "[DEBUG] URL: $($script:ISOUrl)" "DEBUG"

    # Create download script on remote PC
    $downloadScript = @"
`$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
`$url = '$($script:ISOUrl)'
`$out = '$($script:RemotePath)\$($script:ISOName)'
Write-Host "[DEBUG] Downloading from: `$url"
Write-Host "[DEBUG] Saving to: `$out"
try {
    Invoke-WebRequest -Uri `$url -OutFile `$out -UseBasicParsing
    Write-Host "[OK] Download complete"
} catch {
    Write-Host "[ERROR] Download failed: `$(`$_.Exception.Message)"
}
"@

    $scriptPath = "$remoteDirUNC\DownloadISO.ps1"
    $downloadScript | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
    Write-Log "[DEBUG] Download script created on remote PC" "DEBUG"

    # Start download in background using a job
    Write-Log "[INFO] Starting download in background..." "INFO"

    $cmd = "psexec \\$($script:TargetPC) -s -d powershell -ExecutionPolicy Bypass -File `"$($script:RemotePath)\DownloadISO.ps1`" 2>&1"
    Write-Log "[DEBUG] Running: $cmd" "DEBUG"
    cmd /c $cmd | Out-Null

    Write-Log "[INFO] Download started. Monitoring progress..." "INFO"
    Write-Log "[INFO] Press any button to stop monitoring (download continues in background)" "INFO"

    # Monitor progress by checking file size
    $expectedBytes = $script:ExpectedISOSizeGB * 1GB
    $lastSize = 0
    $stuckCount = 0
    $checkInterval = 5  # seconds

    while (-not $script:DownloadCancelled) {
        Start-Sleep -Seconds $checkInterval
        [System.Windows.Forms.Application]::DoEvents()

        if (Test-Path $isoPath) {
            $currentSize = (Get-Item $isoPath).Length
            $currentSizeMB = [math]::Round($currentSize / 1MB, 1)
            $currentSizeGB = [math]::Round($currentSize / 1GB, 2)
            $percentComplete = [math]::Round(($currentSize / $expectedBytes) * 100, 1)

            # Calculate speed
            $speedMBps = [math]::Round(($currentSize - $lastSize) / 1MB / $checkInterval, 2)

            # Estimate time remaining
            if ($speedMBps -gt 0) {
                $remainingBytes = $expectedBytes - $currentSize
                $remainingSeconds = $remainingBytes / ($speedMBps * 1MB)
                $remainingMinutes = [math]::Round($remainingSeconds / 60, 1)
                $etaString = "ETA: ~$remainingMinutes min"
            } else {
                $etaString = "ETA: calculating..."
            }

            # Create progress bar (cap at 100% to avoid negative values)
            $barLength = 30
            $displayPercent = [math]::Min($percentComplete, 100)
            $filledLength = [math]::Max(0, [math]::Min($barLength, [math]::Floor($displayPercent / 100 * $barLength)))
            $emptyLength = [math]::Max(0, $barLength - $filledLength)
            $bar = ("=" * $filledLength) + ("-" * $emptyLength)

            Write-Log "[PROGRESS] [$bar] $percentComplete% ($currentSizeGB GB / $($script:ExpectedISOSizeGB) GB) @ $speedMBps MB/s - $etaString" "INFO"

            # Check if download is complete
            if ($percentComplete -ge 99) {
                Write-Log "[OK] Download appears complete!" "INFO"
                break
            }

            # Check if download is stuck
            if ($currentSize -eq $lastSize) {
                $stuckCount++
                if ($stuckCount -ge 6) {  # 30 seconds of no progress
                    Write-Log "[WARN] Download appears stuck. May have failed or completed." "WARN"
                    break
                }
            } else {
                $stuckCount = 0
            }

            $lastSize = $currentSize
        } else {
            Write-Log "[DEBUG] Waiting for download to start..." "DEBUG"
        }
    }

    # Final verification
    if (Test-Path $isoPath) {
        $size = [math]::Round((Get-Item $isoPath).Length / 1GB, 2)
        if ($size -gt 4) {
            Write-Log "[OK] ISO downloaded successfully ($size GB)" "INFO"
        } else {
            Write-Log "[WARN] ISO file seems too small ($size GB) - download may have failed" "WARN"
        }
    } else {
        Write-Log "[FAIL] ISO not found after download" "ERROR"
        Write-Log "[SOLUTION] Microsoft direct links expire. Download ISO manually and place at $isoPath" "INFO"
    }
})

# ============================================================
# 5. EXTRACT ISO
# ============================================================
$btnExtract.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Extracting ISO on $($script:TargetPC)..." "INFO"

    $remoteDirUNC = "\\$($script:TargetPC)\C$\Win11Upgrade"
    $isoPath = "$remoteDirUNC\$($script:ISOName)"

    if (!(Test-Path $isoPath)) {
        Write-Log "[ERROR] ISO not found at $isoPath" "ERROR"
        Write-Log "[SOLUTION] Run Download ISO step first" "INFO"
        return
    }

    Write-Log "[DEBUG] ISO found, extracting..." "DEBUG"

    # Mount and copy using PowerShell on remote
    try {
        $result = Invoke-Command -ComputerName $script:TargetPC -ScriptBlock {
            $isoFile = "C:\Win11Upgrade\Win11.iso"
            $extractPath = "C:\Win11Upgrade\Extracted"

            # Remove old extraction
            if (Test-Path $extractPath) {
                Remove-Item $extractPath -Recurse -Force
            }
            New-Item -Path $extractPath -ItemType Directory -Force | Out-Null

            # Mount ISO
            Write-Host "[DEBUG] Mounting ISO..."
            $mount = Mount-DiskImage -ImagePath $isoFile -PassThru
            $driveLetter = ($mount | Get-Volume).DriveLetter

            # Copy files
            Write-Host "[DEBUG] Copying files from ${driveLetter}:\ to $extractPath..."
            Copy-Item -Path "${driveLetter}:\*" -Destination $extractPath -Recurse -Force

            # Dismount
            Dismount-DiskImage -ImagePath $isoFile

            return "OK"
        } -ErrorAction Stop

        Write-Log "[OK] ISO extracted successfully to C:\Win11Upgrade\Extracted\" "INFO"
    } catch {
        Write-Log "[ERROR] Extraction failed: $($_.Exception.Message)" "ERROR"
        Write-Log "[DEBUG] Trying PsExec method..." "DEBUG"

        # Fallback - use PowerShell via PsExec
        $extractScript = @"

`$isoFile = 'C:\Win11Upgrade\Win11.iso'
`$extractPath = 'C:\Win11Upgrade\Extracted'
if (Test-Path `$extractPath) { Remove-Item `$extractPath -Recurse -Force }
New-Item -Path `$extractPath -ItemType Directory -Force | Out-Null
`$mount = Mount-DiskImage -ImagePath `$isoFile -PassThru
`$drive = (`$mount | Get-Volume).DriveLetter
Copy-Item -Path "`${drive}:\*" -Destination `$extractPath -Recurse -Force
Dismount-DiskImage -ImagePath `$isoFile
Write-Host 'Extraction complete'
"@
        $scriptPath = "$remoteDirUNC\ExtractISO.ps1"
        $extractScript | Out-File -FilePath $scriptPath -Encoding UTF8 -Force

        $cmd = "psexec \\$($script:TargetPC) -s powershell -ExecutionPolicy Bypass -File `"C:\Win11Upgrade\ExtractISO.ps1`" 2>&1"
        $output = cmd /c $cmd
        foreach ($line in $output) {
            Write-Log $line "DEBUG"
        }
    }

    # Verify extraction
    $setupPath = "$remoteDirUNC\Extracted\setup.exe"
    if (Test-Path $setupPath) {
        Write-Log "[OK] setup.exe found - extraction successful" "INFO"
    } else {
        Write-Log "[WARN] setup.exe not found - check extraction" "WARN"
    }
})

# ============================================================
# 6. START UPGRADE
# ============================================================
$btnInstall.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Starting Windows 11 upgrade on $($script:TargetPC)..." "INFO"

    $remoteDirUNC = "\\$($script:TargetPC)\C$\Win11Upgrade"
    $setupprepPath = "$remoteDirUNC\Extracted\sources\setupprep.exe"

    if (!(Test-Path $setupprepPath)) {
        Write-Log "[ERROR] setupprep.exe not found at $setupprepPath" "ERROR"
        Write-Log "[SOLUTION] Run Extract ISO step first" "INFO"
        return
    }

    Write-Log "[OK] setupprep.exe found" "INFO"
    Write-Log "[DEBUG] Launching upgrade using setupprep.exe /product server..." "DEBUG"
    Write-Log "[INFO] This method bypasses ALL hardware checks" "INFO"

    $cmd = "psexec \\$($script:TargetPC) -s -d `"C:\Win11Upgrade\Extracted\sources\setupprep.exe`" /product server /auto upgrade /quiet /eula accept /dynamicupdate disable 2>&1"
    Write-Log "[DEBUG] Command: $cmd" "DEBUG"

    $output = cmd /c $cmd
    foreach ($line in $output) {
        if ($line -ne "") {
            Write-Log $line "DEBUG"
        }
    }

    Write-Log "[INFO] Upgrade process launched in background" "INFO"
    Write-Log "[INFO] Use 'Monitor Progress' to track the upgrade" "INFO"

    # Start auto-watch
    Start-UpgradeWatch
})

# ============================================================
# 7. MONITOR PROGRESS
# ============================================================
$btnMonitor.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Checking upgrade progress on $($script:TargetPC)..." "INFO"

    # Check for running processes
    Write-Log "[DEBUG] Checking for upgrade processes..." "DEBUG"

    $cmd = "psexec \\$($script:TargetPC) -s cmd /c `"tasklist | findstr /i setup`" 2>&1"
    $output = cmd /c $cmd
    $hasSetup = $output -match "setup"

    if ($hasSetup) {
        Write-Log "[INFO] Setup process is running:" "INFO"
        foreach ($line in $output) {
            if ($line -match "setup") {
                Write-Log "  $line" "INFO"
            }
        }
    } else {
        Write-Log "[INFO] No setup process currently running" "INFO"
    }

    # Check staging folder
    $stagingPath = "\\$($script:TargetPC)\C$\`$WINDOWS.~BT"
    if (Test-Path $stagingPath) {
        Write-Log "[INFO] Staging folder exists - upgrade has started" "INFO"

        # Check for setup log
        $setupLog = "$stagingPath\Sources\Panther\setupact.log"
        if (Test-Path $setupLog) {
            Write-Log "[DEBUG] Reading last 10 lines of setup log..." "DEBUG"
            $logContent = Get-Content $setupLog -Tail 10
            foreach ($line in $logContent) {
                Write-Log "  $line" "DEBUG"
            }
        }
    } else {
        Write-Log "[INFO] Staging folder not found - upgrade may not have started yet" "INFO"
    }

    # Check MoSetup log
    $mosetupLog = "\\$($script:TargetPC)\C$\Windows\Logs\MoSetup\BlueBox.log"
    if (Test-Path $mosetupLog) {
        Write-Log "[DEBUG] Reading last 5 lines of BlueBox.log..." "DEBUG"
        $logContent = Get-Content $mosetupLog -Tail 5
        foreach ($line in $logContent) {
            Write-Log "  $line" "DEBUG"
        }
    }
})

# ============================================================
# 8. VERIFY OS VERSION
# ============================================================
$btnVerify.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Checking OS version on $($script:TargetPC)..." "INFO"

    try {
        $os = Get-WmiObject Win32_OperatingSystem -ComputerName $script:TargetPC -ErrorAction Stop

        Write-Log "[INFO] OS: $($os.Caption)" "INFO"
        Write-Log "[INFO] Version: $($os.Version)" "INFO"
        Write-Log "[INFO] Build: $($os.BuildNumber)" "INFO"

        if ($os.Caption -like "*Windows 11*") {
            Write-Log "[SUCCESS] Windows 11 is installed!" "INFO"
            $lblStatus.Text = "Status: Windows 11 CONFIRMED"
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            Write-Log "[INFO] Still running Windows 10 (or upgrade pending reboot)" "INFO"
        }
    } catch {
        Write-Log "[ERROR] Failed to query OS: $($_.Exception.Message)" "ERROR"
    }
})

# ============================================================
# CLEAR LOG
# ============================================================
$btnClear.Add_Click({
    $txtLog.Clear()
    Write-Log "Log cleared." "INFO"
})

# ============================================================
# FORCE REBOOT BUTTON
# ============================================================
$btnReboot.Add_Click({
    if ($script:TargetPC -eq "") {
        Write-Log "ERROR: Set target PC first!" "ERROR"
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Force reboot $($script:TargetPC)?`n`nThis will immediately restart the PC!",
        "Confirm Reboot",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log "[INFO] Reboot cancelled by user" "INFO"
        return
    }

    Write-Log "========================================" "INFO"
    Write-Log "Force rebooting $($script:TargetPC)..." "INFO"

    # Send reboot command
    $cmd = "psexec \\$($script:TargetPC) -s shutdown /r /f /t 5 /c `"Remote reboot initiated`" 2>&1"
    Write-Log "[DEBUG] Command: $cmd" "DEBUG"
    $output = cmd /c $cmd
    foreach ($line in $output) {
        if ($line -ne "") { Write-Log $line "DEBUG" }
    }

    Write-Log "[INFO] Reboot command sent. PC will restart in 5 seconds." "INFO"
    Write-Log "[INFO] Monitoring for PC to come back online..." "INFO"

    # Wait a moment for PC to start shutting down
    Start-Sleep -Seconds 10

    # Start ping monitoring
    $maxAttempts = 60  # 5 minutes max
    $attempt = 0
    $wasOffline = $false

    while ($attempt -lt $maxAttempts) {
        $attempt++
        [System.Windows.Forms.Application]::DoEvents()

        $ping = Test-Connection -ComputerName $script:TargetPC -Count 1 -Quiet -ErrorAction SilentlyContinue

        if ($ping) {
            if ($wasOffline) {
                Write-Log "[PING] $($script:TargetPC) is ONLINE! (attempt $attempt)" "INFO"
                Write-Log "[SUCCESS] PC has rebooted and is back online!" "INFO"
                [System.Media.SystemSounds]::Exclamation.Play()
                [System.Windows.Forms.MessageBox]::Show("$($script:TargetPC) is back online!", "PC Online", "OK", "Information")
                break
            } else {
                Write-Log "[PING] $($script:TargetPC) still online (shutting down...)" "DEBUG"
            }
        } else {
            if (-not $wasOffline) {
                Write-Log "[PING] $($script:TargetPC) went OFFLINE (rebooting...)" "INFO"
                $wasOffline = $true
            } else {
                Write-Log "[PING] $($script:TargetPC) offline... waiting (attempt $attempt)" "DEBUG"
            }
        }

        Start-Sleep -Seconds 5
    }

    if ($attempt -ge $maxAttempts) {
        Write-Log "[WARN] Timeout waiting for PC to come back online" "WARN"
    }
})

# ============================================================
# AUTO-WATCH FUNCTION (checks if upgrade complete, triggers reboot)
# ============================================================
$script:WatchTimer = New-Object System.Windows.Forms.Timer
$script:WatchTimer.Interval = 30000  # 30 seconds

$script:WatchTimer.Add_Tick({
    if ($script:TargetPC -eq "") { return }

    # Check if setup processes are still running
    $cmd = "psexec \\$($script:TargetPC) -s cmd /c `"tasklist | findstr /i setup`" 2>&1"
    $output = cmd /c $cmd
    $hasSetup = $output -match "setup"

    # Check staging folder for completion markers
    $stagingPath = "\\$($script:TargetPC)\C$\`$WINDOWS.~BT"
    $setupLog = "$stagingPath\Sources\Panther\setupact.log"

    $upgradeComplete = $false
    $upgradeFailed = $false

    if (Test-Path $setupLog) {
        $logTail = Get-Content $setupLog -Tail 50 -ErrorAction SilentlyContinue
        $logText = $logTail -join "`n"

        # Check for completion indicators
        if ($logText -match "MOUPG.*Finalize phase completed" -or
            $logText -match "Overall progress: \[100%\]" -or
            $logText -match "RebootMachine" -or
            $logText -match "Reboot is required") {
            $upgradeComplete = $true
        }

        # Check for failure
        if ($logText -match "MOUPG.*failed" -or $logText -match "FatalError") {
            $upgradeFailed = $true
        }
    }

    # If no setup process and staging exists, might be done
    if (-not $hasSetup -and (Test-Path $stagingPath)) {
        # Double-check by looking for pending reboot state
        $upgradeComplete = $true
    }

    if ($upgradeFailed) {
        $script:WatchTimer.Stop()
        Write-Log "[ERROR] Upgrade appears to have FAILED!" "ERROR"
        Write-Log "[INFO] Check Monitor Progress for details" "INFO"
        [System.Media.SystemSounds]::Hand.Play()
        [System.Windows.Forms.MessageBox]::Show("Upgrade FAILED on $($script:TargetPC)!`nCheck logs for details.", "Upgrade Failed", "OK", "Error")
    }
    elseif ($upgradeComplete) {
        $script:WatchTimer.Stop()
        Write-Log "[SUCCESS] Upgrade pre-reboot phase COMPLETE!" "INFO"
        [System.Media.SystemSounds]::Exclamation.Play()

        if ($script:AutoReboot) {
            Write-Log "[INFO] Auto-reboot enabled - rebooting $($script:TargetPC) now..." "INFO"
            $rebootCmd = "psexec \\$($script:TargetPC) -s shutdown /r /t 30 /c `"Windows 11 upgrade complete - rebooting in 30 seconds`" 2>&1"
            cmd /c $rebootCmd
            Write-Log "[INFO] Reboot command sent. PC will restart in 30 seconds." "INFO"
            [System.Windows.Forms.MessageBox]::Show("Upgrade complete on $($script:TargetPC)!`nPC is rebooting in 30 seconds to finish installation.", "Upgrade Complete", "OK", "Information")
        } else {
            Write-Log "[INFO] Auto-reboot disabled. Reboot manually to complete." "INFO"
            [System.Windows.Forms.MessageBox]::Show("Upgrade complete on $($script:TargetPC)!`nReboot the PC manually to finish installation.", "Upgrade Complete - Reboot Required", "OK", "Information")
        }
    }
    else {
        # Still in progress - log a brief status
        Write-Log "[WATCH] Upgrade still in progress... (checking every 30s)" "DEBUG"
    }
})

# Start watching after upgrade begins
function Start-UpgradeWatch {
    Write-Log "[INFO] Auto-watch started - will notify when complete" "INFO"
    if ($script:AutoReboot) {
        Write-Log "[INFO] Auto-reboot is ENABLED - PC will reboot automatically" "INFO"
    } else {
        Write-Log "[INFO] Auto-reboot is DISABLED - manual reboot required" "INFO"
    }
    $script:WatchTimer.Start()
}

# ============================================================
# STARTUP
# ============================================================
Write-Log "Windows 11 Remote Upgrade Tool Started" "INFO"
Write-Log "========================================" "INFO"
Write-Log "Steps:" "INFO"
Write-Log "  1. Enter target PC name and click 'Set Target'" "INFO"
Write-Log "  2. Test Connection" "INFO"
Write-Log "  3. Check Storage (need 30+ GB free)" "INFO"
Write-Log "  4. Apply Bypass (registry keys)" "INFO"
Write-Log "  5. Download ISO (or manually copy)" "INFO"
Write-Log "  6. Extract ISO" "INFO"
Write-Log "  7. Start Upgrade" "INFO"
Write-Log "  8. Monitor Progress / Verify OS" "INFO"
Write-Log "========================================" "INFO"

# Show form
[void]$form.ShowDialog()
