# Win10 to 11 Update Utility

A PowerShell-based GUI tool for remotely upgrading Windows 10 PCs to Windows 11, with full support for bypassing Microsoft's hardware requirements (TPM, Secure Boot, CPU checks).

## Features

- **Remote deployment** - Upgrade PCs across your network without physical access
- **Hardware bypass** - Automatically applies all known registry bypasses for unsupported hardware
- **Batch/Parallel mode** - Upgrade multiple PCs simultaneously with status tracking
- **Auto-detection** - Detects current state of each PC on session restart
- **Step-by-step workflow** - Each operation is a separate button for controlled execution
- **Live progress monitoring** - Real-time download progress with speed and ETA
- **Auto-reboot option** - Automatically reboot PCs when upgrade completes
- **Force reboot with ping monitoring** - Remote reboot with live status until PC comes back online
- **Debug logging** - Full terminal output for troubleshooting
- **Domain-ready** - Uses PsExec and admin shares for enterprise environments

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
3. Edit the PC list in the left panel (one PC per line)
4. Click **Load PC List** - auto-detects current state of each PC
5. Click **Test All Connections** - verifies which PCs are online
6. Follow steps 3-6 in order:
   - **Apply Bypass (All)** - Registry keys applied in parallel
   - **Download ISO (All)** - Downloads start on all PCs simultaneously
   - **Extract ISO (All)** - Extracts in parallel
   - **Start Upgrade (All)** - Launches upgrades on all ready PCs
7. Use **Monitor All** to check progress
8. **Reboot All Ready** or let auto-reboot handle it
9. **Verify OS** to confirm Windows 11 is installed

**Tip:** For a single PC, just enter one PC name in the list.

## Additional Features

- **Check Storage (All)** - Shows free disk space on all PCs
- **Force Reboot Selected** - Click a PC in the list, then force reboot with ping monitoring
- **60-day rollback window** - Automatically extended during bypass (default is 10 days)

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
| `HKLM\SYSTEM\Setup\MoSetup\AllowUpgradesWithUnsupportedTPMOrCPU` | 1 | Main bypass |
| `HKLM\SYSTEM\Setup\LabConfig\BypassTPMCheck` | 1 | Skip TPM 2.0 check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassSecureBootCheck` | 1 | Skip Secure Boot check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassRAMCheck` | 1 | Skip 4GB RAM check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassCPUCheck` | 1 | Skip CPU compatibility check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassStorageCheck` | 1 | Skip storage check |
| `HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DisableWUfBSafeguards` | 1 | Disable safeguard holds |

## Upgrade Command

The tool uses `setupprep.exe` with the `/product server` flag to bypass hardware checks:

```
setupprep.exe /product server /auto upgrade /quiet /eula accept /dynamicupdate disable
```

- `/product server` - Treats upgrade as Server SKU (bypasses consumer hardware checks)
- `/auto upgrade` - In-place upgrade keeping files/apps
- `/quiet` - No UI (silent)
- `/eula accept` - Auto-accept license
- `/dynamicupdate disable` - Skip downloading updates during setup (prevents hangs)

## Rollback

Windows keeps a rollback option for **10 days** after upgrade:
- Settings → System → Recovery → "Go back to Windows 10"
- Requires `C:\Windows.old` folder (don't run Disk Cleanup!)

## Troubleshooting

### "PC not reachable"
- Verify PC is powered on and connected to network
- Check firewall allows ICMP ping and SMB (port 445)

### "Cannot access admin share"
- Ensure you're running as domain admin
- Verify File and Printer Sharing is enabled on target PC

### Upgrade fails with 0xC1900200
- Hardware requirements not bypassed - run "Apply Bypass" step
- The `/product server` flag should handle this

### Download fails / stuck at "waiting"
- Microsoft ISO links may expire - check URL is still valid
- Check target PC has internet access
- Manually download ISO and copy to `\\PCNAME\C$\Win11Upgrade\Win11.iso`

### Upgrade stuck at 0%
- Add `/dynamicupdate disable` flag (already included)
- Check if `SetupHost.exe` is running on target PC

### Use the Diagnostic Tool
Run `Win11Diagnose.bat` to generate a comprehensive report of potential issues.

## Session Persistence

- **Downloads** run on target PCs - closing GUI doesn't interrupt them
- **Upgrades** run on target PCs - closing GUI doesn't interrupt them
- **Auto-detect** on "Load PC List" finds current state of each PC
- Safe to close and reopen the tool at any time

## How It Works

1. **Registry Bypass**: Adds registry keys that tell Windows Setup to skip hardware checks
2. **ISO Download**: Downloads Windows 11 ISO directly to target PC (background process)
3. **Extraction**: Mounts the ISO and copies files to a local folder
4. **Silent Upgrade**: Runs `setupprep.exe /product server` to bypass all hardware checks
5. **Auto-Watch**: Monitors setup logs for completion, triggers reboot when ready

## Security Notes

- Scripts run with SYSTEM privileges via PsExec
- Admin share access requires domain admin or equivalent rights
- ISO is downloaded from official Microsoft servers
- No third-party tools or modifications to Windows files

## License

MIT License - Use at your own risk. Not affiliated with Microsoft.

## Disclaimer

Bypassing Windows 11 hardware requirements means Microsoft may not support your installation. While upgrades typically work fine, you may encounter:
- Potential issues with future updates on very old hardware
- No official support from Microsoft

Test on a single PC before mass deployment.
