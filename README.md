# Win10 to 11 Update Utility

A PowerShell-based GUI tool for remotely upgrading Windows 10 PCs to Windows 11, with full support for bypassing Microsoft's hardware requirements (TPM, Secure Boot, CPU checks).

## Features

- **Remote deployment** - Upgrade PCs across your network without physical access
- **Hardware bypass** - Automatically applies all known registry bypasses for unsupported hardware
- **Dual-grid PC selection** - Available PCs list + Target PCs list with drag-and-drop style interaction
- **Parallel operations** - Test, bypass, download, extract, and upgrade multiple PCs simultaneously
- **Selective targeting** - Select specific PCs for actions, or run on all targets
- **Auto-detection** - Detects current state of each PC (Downloading, Extracted, Upgrading, Complete)
- **Live status tracking** - Color-coded status for each PC in real-time
- **60-day rollback window** - Automatically extended during bypass (default is only 10 days)
- **Auto-reboot option** - Automatically reboot PCs when upgrade completes
- **Force reboot with ping monitoring** - Remote reboot with live status until PC comes back online
- **Debug logging** - Full terminal output for troubleshooting
- **Domain-ready** - Uses PsExec and admin shares for enterprise environments
- **Session persistence** - Downloads/upgrades continue even if GUI is closed

## Requirements

### On Your Admin PC
- Windows with PowerShell 5.1+
- [PsExec](https://docs.microsoft.com/en-us/sysinternals/downloads/psexec) (from Sysinternals) in PATH or System32
- Domain admin credentials (or local admin on target PCs)
- Network access to target PCs

### On Target PCs
- Windows 10 (any version)
- File and Printer Sharing enabled (default in domain environments)
- Admin share accessible (`\\PCNAME\C$`)
- ~30GB free disk space
- Internet access (for ISO download)

## Files

| File | Description |
|------|-------------|
| `Win11UpgradeGUI.ps1` | **Main GUI** - supports single or multiple PCs with parallel upgrades |
| `Win11Bypass.bat` | Standalone batch script to apply registry bypass to multiple PCs |
| `Win11Upgrade-Test.bat` | Test script for single PC upgrade (downloads and installs) |
| `Win11Diagnose.bat` | Diagnostic tool to troubleshoot upgrade failures |
| `DownloadWin11.ps1` | Helper script copied to remote PCs for ISO download |

## Quick Start

1. Run PowerShell as Administrator
2. Execute:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "Win11UpgradeGUI.ps1"
   ```

## User Interface

### Left Panel - Available PCs
- Pre-populated list of PCs you can upgrade
- **Double-click** a PC to add it to the Target list
- **Add All >>** button adds all available PCs at once
- **Clear Targets** button removes all PCs from target list

### Middle Panel - Target PCs
- Shows PCs you've selected for upgrade with their current status
- **Double-click** a PC to remove it from targets
- **Multi-select** (Ctrl+click or Shift+click) to select specific PCs
- Color-coded status indicators

### Right Panel - Actions
All actions run on **selected PCs** if any are selected, otherwise runs on **all target PCs**.

| Button | Description |
|--------|-------------|
| 1. Test Connections | Ping and verify admin share access |
| 2. Apply Bypass | Apply registry bypasses + extend rollback to 60 days |
| 3. Download ISO | Download Windows 11 ISO to each PC (background) |
| 4. Extract ISO | Mount and extract ISO contents |
| 5. Start Upgrade | Launch silent upgrade with hardware bypass |
| 6. Monitor Progress | Check upgrade status and progress % |
| 7. Verify OS | Confirm Windows 11 is installed |
| Reboot Ready | Reboot all PCs that completed upgrade |
| Check Storage | Show free disk space on target PCs |
| Force Reboot | Reboot selected PC with ping monitoring |
| Rollback | Roll back selected PC to Windows 10 (remote) |

## Workflow

1. **Add PCs** - Double-click PCs in Available list, or click "Add All >>"
2. **Test Connections** - Verify all target PCs are online and accessible
3. **Apply Bypass** - Set registry keys + extend rollback window to 60 days
4. **Download ISO** - Downloads run in background on each PC (~6.5 GB)
5. **Extract ISO** - Mounts ISO and copies files locally
6. **Start Upgrade** - Launches silent upgrade (takes 30-60 minutes)
7. **Monitor Progress** - Check status periodically
8. **Reboot** - Auto-reboot when ready, or use Reboot Ready button
9. **Verify OS** - Confirm Windows 11 is installed

## Status Colors

| Color | Status |
|-------|--------|
| White | Pending / Online |
| Gray | Offline |
| Yellow | Downloading |
| Blue | Upgrading |
| Green | Complete (ready for reboot) |
| Red | Failed |

## Registry Bypasses Applied

The tool applies all known Windows 11 requirement bypasses:

| Registry Key | Value | Purpose |
|-------------|-------|---------|
| `HKLM\SYSTEM\Setup\MoSetup\AllowUpgradesWithUnsupportedTPMOrCPU` | 1 | Main bypass flag |
| `HKLM\SYSTEM\Setup\LabConfig\BypassTPMCheck` | 1 | Skip TPM 2.0 requirement |
| `HKLM\SYSTEM\Setup\LabConfig\BypassSecureBootCheck` | 1 | Skip Secure Boot requirement |
| `HKLM\SYSTEM\Setup\LabConfig\BypassRAMCheck` | 1 | Skip 4GB RAM requirement |
| `HKLM\SYSTEM\Setup\LabConfig\BypassCPUCheck` | 1 | Skip CPU compatibility check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassStorageCheck` | 1 | Skip storage requirement |
| `HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DisableWUfBSafeguards` | 1 | Disable Windows Update safeguard holds |

Additionally, the tool runs:
```
DISM /Online /Set-OSUninstallWindow /Value:60
```
This extends the rollback window from 10 days to **60 days**.

## Upgrade Command

The tool uses `setupprep.exe` with the `/product server` flag to bypass hardware checks:

```
setupprep.exe /product server /auto upgrade /quiet /eula accept /dynamicupdate disable
```

| Flag | Purpose |
|------|---------|
| `/product server` | Treats upgrade as Server SKU (bypasses consumer hardware checks) |
| `/auto upgrade` | In-place upgrade keeping files and apps |
| `/quiet` | No UI (silent operation) |
| `/eula accept` | Auto-accept license agreement |
| `/dynamicupdate disable` | Skip downloading updates during setup (prevents hangs) |

## Rollback to Windows 10

After upgrading, users have **60 days** to rollback (extended from default 10 days):

### Via GUI (Recommended)
1. Select the PC in the Target PCs list
2. Click the **Rollback** button
3. The tool will:
   - Check if rollback is available
   - Show days remaining in rollback window
   - Initiate the rollback remotely
   - Monitor the PC until it comes back online
   - Verify Windows 10 is restored

### Via Settings (Manual)
- Settings → System → Recovery → "Go back to Windows 10"

### Via Command (Remote)
```powershell
# Check rollback availability
psexec \\PCNAME -s dism /Online /Get-OSUninstallWindow

# Initiate rollback
psexec \\PCNAME -s dism /Online /Initiate-OSUninstall /Quiet
```

### Requirements
- `C:\Windows.old` folder must exist (don't run Disk Cleanup!)
- Must be within the 60-day rollback window
- PC must be on Windows 11 (upgraded from Windows 10)

## Session Persistence

- **Downloads** run as background processes on target PCs - closing GUI doesn't interrupt them
- **Upgrades** run as background processes on target PCs - closing GUI doesn't interrupt them
- **Auto-detect** when adding PCs finds their current state automatically
- Safe to close and reopen the tool at any time
- Can run multiple GUI instances for different PC groups

## Troubleshooting

### "PC not reachable"
- Verify PC is powered on and connected to network
- Check firewall allows ICMP ping and SMB (port 445)
- Try `ping PCNAME` from command prompt

### "Cannot access admin share"
- Ensure you're running PowerShell as Administrator
- Verify you have domain admin or local admin rights
- Check File and Printer Sharing is enabled on target PC
- Try `dir \\PCNAME\C$` from command prompt

### Download stuck at "waiting"
- Microsoft ISO links may expire - check URL is still valid
- Verify target PC has internet access
- Manually download ISO and copy to `\\PCNAME\C$\Win11Upgrade\Win11.iso`

### Upgrade fails with 0xC1900200
- Hardware requirements not bypassed - run "Apply Bypass" step again
- The `/product server` flag should handle this automatically

### Upgrade stuck at 0%
- The `/dynamicupdate disable` flag (already included) prevents this
- Check if `SetupHost.exe` is running on target PC via Task Manager
- Use "Monitor Progress" to see detailed log output

### Use the Diagnostic Tool
Run `Win11Diagnose.bat` on the target PC to generate a comprehensive report of potential issues.

## How It Works

1. **Registry Bypass**: Adds registry keys that tell Windows Setup to skip hardware checks
2. **Rollback Extension**: Uses DISM to extend rollback window to 60 days
3. **ISO Download**: Downloads Windows 11 ISO directly to target PC via PowerShell (background process)
4. **Extraction**: Mounts the ISO and copies all files to `C:\Win11Upgrade\Extracted\`
5. **Silent Upgrade**: Runs `setupprep.exe /product server` which bypasses all consumer hardware checks
6. **Auto-Reboot**: When upgrade completes, automatically reboots (if enabled) to finish installation
7. **Verification**: Queries WMI to confirm Windows 11 is installed

## Security Notes

- Scripts run with SYSTEM privileges via PsExec
- Admin share access requires domain admin or equivalent rights
- ISO is downloaded from official Microsoft servers
- No third-party tools or modifications to Windows system files
- All operations are logged for audit purposes

## License

MIT License - Use at your own risk. Not affiliated with Microsoft.

## Disclaimer

Bypassing Windows 11 hardware requirements means Microsoft may not fully support your installation. While upgrades typically work fine, you may encounter:
- Potential issues with future updates on very old hardware (pre-2010 CPUs)
- No official support from Microsoft for hardware-related issues

**Recommendations:**
- Test on a single PC before mass deployment
- Keep the 60-day rollback window in mind for quick recovery
- Document which PCs were upgraded for future reference
