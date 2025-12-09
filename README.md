# Win10 to 11 Update Utility

A PowerShell-based GUI tool for remotely upgrading Windows 10 PCs to Windows 11, with full support for bypassing Microsoft's hardware requirements (TPM, Secure Boot, CPU checks).

## Features

- **Remote deployment** - Upgrade PCs across your network without physical access
- **Hardware bypass** - Automatically applies all known registry bypasses for unsupported hardware
- **Step-by-step workflow** - Each operation is a separate button for controlled execution
- **Live progress monitoring** - Real-time download progress with speed and ETA
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
| `Win11UpgradeGUI.ps1` | Main GUI application - use this for most operations |
| `Win11Bypass.bat` | Standalone batch script to apply registry bypass to multiple PCs |
| `Win11Upgrade-Test.bat` | Test script for single PC upgrade (downloads and installs) |
| `Win11Diagnose.bat` | Diagnostic tool to troubleshoot upgrade failures |
| `DownloadWin11.ps1` | Helper script copied to remote PCs for ISO download |

## Quick Start

### Using the GUI (Recommended)

1. Run PowerShell as Administrator
2. Execute:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "Win11UpgradeGUI.ps1"
   ```
3. Enter target PC hostname and click "Set Target"
4. Follow steps 1-8 in order:
   - **Test Connection** - Verify PC is reachable
   - **Check Storage** - Ensure 30GB+ free space
   - **Apply Bypass** - Set all registry keys
   - **Download ISO** - Download Windows 11 directly to target PC
   - **Extract ISO** - Mount and extract ISO contents
   - **Start Upgrade** - Begin silent upgrade
   - **Monitor Progress** - Check upgrade status
   - **Verify OS** - Confirm Windows 11 is installed

### Using Batch Scripts (Alternative)

For bulk operations on multiple PCs:

1. Edit `Win11Bypass.bat` to include your PC list
2. Run as Administrator to apply registry bypass to all PCs
3. Use `Win11Upgrade-Test.bat` to test on a single PC first

## Registry Bypasses Applied

The tool applies all known Windows 11 requirement bypasses:

| Registry Key | Value | Purpose |
|-------------|-------|---------|
| `HKLM\SYSTEM\Setup\MoSetup\AllowUpgradesWithUnsupportedTPMOrCPU` | 1 | Main bypass |
| `HKLM\SYSTEM\Setup\LabConfig\BypassTPMCheck` | 1 | Skip TPM 2.0 check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassSecureBootCheck` | 1 | Skip Secure Boot check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassRAMCheck` | 1 | Skip 4GB RAM check |
| `HKLM\SYSTEM\Setup\LabConfig\BypassCPUCheck` | 1 | Skip CPU compatibility check |

## Troubleshooting

### "PC not reachable"
- Verify PC is powered on and connected to network
- Check firewall allows ICMP ping and SMB (port 445)

### "Cannot access admin share"
- Ensure you're running as domain admin
- Verify File and Printer Sharing is enabled on target PC

### Upgrade fails with 0xC1900200
- Hardware requirements not bypassed - run "Apply Bypass" step
- Use the `/compat IgnoreWarning` flag (already included in scripts)

### Download fails
- Microsoft ISO links may expire - script uses static URLs
- Check target PC has internet access
- Manually download ISO and copy to `\\PCNAME\C$\Win11Upgrade\Win11.iso`

### Use the Diagnostic Tool
Run `Win11Diagnose.bat` to generate a comprehensive report of potential issues.

## How It Works

1. **Registry Bypass**: Adds registry keys that tell Windows Setup to skip hardware checks
2. **ISO Download**: Downloads Windows 11 ISO directly to the target PC from Microsoft's static servers
3. **Extraction**: Mounts the ISO and copies files to a local folder (required for remote execution)
4. **Silent Upgrade**: Runs `setup.exe` with flags:
   - `/auto upgrade` - In-place upgrade keeping files/apps
   - `/eula accept` - Auto-accept license
   - `/compat IgnoreWarning` - Ignore compatibility warnings
   - `/quiet` - No UI
   - `/noreboot` - Don't auto-reboot (user controls timing)

## Security Notes

- Scripts run with SYSTEM privileges via PsExec
- Admin share access requires domain admin or equivalent rights
- ISO is downloaded from official Microsoft servers
- No third-party tools or modifications to Windows files

## License

MIT License - Use at your own risk. Not affiliated with Microsoft.

## Disclaimer

Bypassing Windows 11 hardware requirements means Microsoft may not support your installation. While upgrades typically work fine, you may encounter:
- No security updates (Microsoft has stated they may withhold updates)
- Potential stability issues on very old hardware
- No official support from Microsoft

Test on a single PC before mass deployment.
