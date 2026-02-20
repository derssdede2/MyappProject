# DLack — Pre-Deployment Audit Report

**Application:** Endpoint Diagnostics & Remediation v2.0  
**Platform:** WPF (.NET 8, Windows)  
**Build Status:** ✅ Compiles successfully (net8.0-windows)  
**Audit Date:** 2025  
**Repository:** https://github.com/DilerKami/DLack (branch: master)

---

## Table of Contents

1. [Application Overview](#1-application-overview)
2. [Project Structure & Dependencies](#2-project-structure--dependencies)
3. [Screen-by-Screen Walkthrough](#3-screen-by-screen-walkthrough)
4. [Scan Module — Detailed Audit](#4-scan-module--detailed-audit)
5. [Optimizer Module — Detailed Audit](#5-optimizer-module--detailed-audit)
6. [Verification Module](#6-verification-module)
7. [PDF Report Generator](#7-pdf-report-generator)
8. [Theme & UI System](#8-theme--ui-system)
9. [Bugs, Risks & Inconsistencies](#9-bugs-risks--inconsistencies)
10. [Concrete Fixes & Recommendations](#10-concrete-fixes--recommendations)
11. [Security Audit](#11-security-audit)
12. [Deployment Checklist](#12-deployment-checklist)

---

## 1. Application Overview

### What it does (end-to-end)

DLack ("Endpoint Diagnostics & Remediation") is a **Windows admin tool** that:

1. **Launches** — Requires administrator privileges (enforced via `app.manifest` `requireAdministrator` and `App.xaml.cs` check). Initializes theme (dark/light) from `%LOCALAPPDATA%\DLack\theme.json`.

2. **Scans** — Runs up to 20 diagnostic phases sequentially (per-phase timeouts: 15–60s depending on WMI risk):
   - System Overview, CPU, RAM, GPU, Disk, Thermal, Battery, Startup, Visual Settings, Antivirus, Windows Update, Network, Internet Speed Test, Network Drives, Outlook, Browser, User Profile, Office, Installed Software, Event Log Analysis.

3. **Displays Results** — Shows a health score (0–100), health dashboard cards with animated arc gauges, and detailed metrics for every category. Flagged issues are sorted by severity (Critical → Warning → Info).

4. **Optimizes** — Builds a plan of automatable fixes. User selects which actions to run. Executes them with progress tracking and live countdown timers. After execution, runs a verification pass to confirm each fix was applied.

5. **Exports PDF** — Generates a comprehensive A4 PDF report using QuestPDF with branded Pillsbury headers, color-coded severity, tables, and recommendations. Includes optimization results if optimization was performed.

### Architecture

| Layer | File(s) | Responsibility |
|-------|---------|----------------|
| Entry | `App.xaml.cs` | Admin check, theme init |
| Main UI | `MainWindow.xaml/.cs` | Scan trigger, results display, export, optimize button |
| Scanner | `Scanner.cs` | Up to 20 diagnostic phases (speed test skippable), health score calculation |
| Data Models | `DiagnosticResult.cs`, `OptimizationAction.cs` | DTOs for all scan/optimization data |
| Optimizer | `Optimizer.cs` | Plan builder, action dispatcher, verification |
| Optimization UI | `OptimizationWindow.xaml/.cs` | Action selection, execution, verification overlay |
| Results UI | `OptimizationResultsWindow.xaml/.cs` | Post-optimization results with auto-rescan countdown |
| PDF | `PDFReportGenerator.cs` | QuestPDF-based report generation |
| Theming | `ThemeManager.cs`, `Themes/*.xaml` | Light/dark theme with persisted preference |
| Formatting | `DiagnosticFormatters.cs` | Shared formatting helpers (MB, uptime, DNS/ping rating) |
| Shared Helpers | `NativeHelpers.cs` | P/Invoke (Recycle Bin) and folder-size utilities |

---

## 2. Project Structure & Dependencies

### Project File (`DLack.csproj`)

```xml
<TargetFramework>net8.0-windows</TargetFramework>
<UseWPF>true</UseWPF>
<ApplicationManifest>app.manifest</ApplicationManifest>
```

### NuGet Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `QuestPDF` | 2025.12.4 | PDF report generation (Community license) |
| `System.Management` | 10.0.3 | WMI queries for hardware diagnostics |
| `System.ServiceProcess.ServiceController` | 9.0.3 | Windows service management (start/stop wuauserv, FontCache) |

### App Manifest

- `requireAdministrator` UAC level — enforced at OS level
- Per-monitor DPI awareness (`PerMonitorV2`)
- Supported OS: Windows Vista through Windows 10+

---

## 3. Screen-by-Screen Walkthrough

### MainWindow
- **Header:** Pillsbury logo, app version, device context subtitle, theme toggle (Ctrl+T)
- **Body:** Scrollable metrics panel showing all 20 scan categories with gauge rows, info rows, and flagged issue cards
- **Health Dashboard:** 4 animated cards — Health Score arc gauge, Critical count, Warning count, Scanned categories
- **Action Strip:** Quick-glance bar showing issue count, estimated fix time, and reclaimable disk space
- **Buttons:** Scan (Ctrl+R), Export Report (Ctrl+E), Optimize (Ctrl+O), Cancel (Esc)
- **Before/After:** After optimization + rescan, shows side-by-side comparison of Health Score, Disk Free, RAM Usage, Startup Items, Issues

### OptimizationWindow
- **Sections:** Automatic Fixes (checkboxes), Manual Actions (Open buttons / info icons)
- **Cards:** Left accent bar (green/yellow/red), risk badges (SAFE/MODERATE/REBOOT), estimated duration/space, verification badges
- **Execution:** Animated overlay with comet-tail spinner, per-action countdown timer, progress bar
- **Post-execution:** Inline result messages, verification status per action, summary counts

### OptimizationResultsWindow
- **Hero badge:** Animated elastic pop with checkmark (all good) or warning (failures)
- **Summary banner:** Action count breakdown, freed space count-up animation
- **Detail cards:** Ordered by status (Failed → Partial → NoChange → Success), hover lift animation
- **Auto-verify:** 5-second countdown to trigger rescan, or manual "Verify Now" button

---

## 4. Scan Module — Detailed Audit

### Phase 1: System Overview
| Check | Source | Details |
|-------|--------|---------|
| Computer Name | `Environment.MachineName` | |
| Manufacturer/Model | WMI `Win32_ComputerSystem` | 10s timeout |
| Serial Number | WMI `Win32_BIOS` | 10s timeout |
| Windows Version/Build | WMI `Win32_OperatingSystem` | |
| CPU Model/Clock | WMI `Win32_Processor` (Name, MaxClockSpeed, CurrentClockSpeed) | Cached for throttle detection |
| Total RAM | P/Invoke `GlobalMemoryStatusEx` | |
| Uptime | `Environment.TickCount64` | **Flagged** if > 7 days |

### Phase 2: CPU Diagnostics
| Check | Source | Details |
|-------|--------|---------|
| CPU Load | WMI `Win32_Processor.LoadPercentage` | 3 samples × 500ms apart, averaged |
| **Flagged** | Load > 30% | |

### Phase 3: RAM Diagnostics
| Check | Source | Details |
|-------|--------|---------|
| Usage | P/Invoke `GlobalMemoryStatusEx` | Total, Used, Available, PercentUsed |
| Top Processes | `Process.GetProcesses()` sorted by `WorkingSet64` | Top 10 |
| Memory Diagnostic | Registry `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsMemoryDiagnostic\LastResult` | |
| **Flagged** | Usage > 85%, or TotalRAM < 8 GB | |

### Phase 4: Disk Diagnostics
| Check | Source | Details |
|-------|--------|---------|
| Drive info | `DriveInfo.GetDrives()` | Total, Free, Used, PercentUsed |
| Drive type (HDD/SSD) | WMI `MSFT_PhysicalDisk.MediaType` via `\\.\root\microsoft\windows\storage` | Maps logical → partition → physical disk |
| Drive health | WMI `Win32_DiskDrive.Status` via partition association | Falls back to single-disk check |
| Disk activity | WMI `Win32_PerfFormattedData_PerfDisk_PhysicalDisk` (`PercentDiskTime`) | |
| Windows Temp | `%WINDIR%\Temp` (depth 3) | Resolved via `Environment.SpecialFolder.Windows` |
| User Temp | `Path.GetTempPath()` (depth 3) | |
| SW Distribution | `%WINDIR%\SoftwareDistribution` (depth 4) | |
| Prefetch | `%WINDIR%\Prefetch` (depth 1) | |
| Windows.old | `%SystemDrive%\Windows.old` (depth 3) | Root from `Path.GetPathRoot()` |
| Upgrade Logs | `%WINDIR%\Logs\WindowsUpdate`, `%WINDIR%\Panther`, `%SystemDrive%\$Windows.~BT`, `%SystemDrive%\$Windows.~WS` | |
| Recycle Bin | P/Invoke `SHQueryRecycleBin` | |
| **Flagged** | Drive > 85% used, health != Healthy, activity > 50% | |

### Phase 5: Thermal Diagnostics
| Check | Source | Details |
|-------|--------|---------|
| CPU Temperature | WMI `MSAcpi_ThermalZoneTemperature` (`root\WMI`) | Kelvin/10 → Celsius. First zone = CPU |
| Throttle detection | CurrentClockSpeed < 50% of MaxClockSpeed AND temp > 70°C | Dual-condition prevents false positives from power management |
| Fan status | WMI `Win32_Fan` | Count of detected fans |
| **Flagged** | Temp > 80°C, throttling detected | |

### Phase 6: GPU Diagnostics
| Check | Source | Details |
|-------|--------|---------|
| GPU enumeration | WMI `Win32_VideoController` | All adapters |
| VRAM (accurate) | Registry `HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\*` | 64-bit `qwMemorySize`, fallback to `MemorySize` DWORD |
| Primary GPU selection | Prefers dedicated (NVIDIA/AMD/Quadro/Arc) over integrated | |
| Temperature (NVIDIA) | CLI `nvidia-smi --query-gpu=temperature.gpu` | 5s timeout |
| Temperature (fallback) | Cached thermal zones from Phase 5 | Second zone if delta > 5°C from first |
| Usage (NVIDIA) | CLI `nvidia-smi --query-gpu=utilization.gpu` | |
| Usage (fallback) | WMI `Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine` | Max utilization across all engines |
| Display enumeration | P/Invoke `EnumDisplayDevices` + `EnumDisplaySettings` | Physical monitors with resolution/refresh rate |
| Driver age | `DriverDate` from WMI, parsed as `yyyyMMdd` | **Flagged** if > 365 days old |
| **Flagged** | Temp > 85°C, usage > 90%, driver > 1 year | |

### Phase 7: Battery Health
| Check | Source | Details |
|-------|--------|---------|
| Battery presence | WMI `Win32_Battery` | |
| Health % | `FullChargeCapacity / DesignCapacity × 100` | |
| Power plan | CLI `powercfg /getactivescheme` | Parses plan name from parentheses |
| **Flagged** | Health < 50% | |

### Phase 8: Startup Programs
| Check | Source | Details |
|-------|--------|---------|
| Run keys | `HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run` | |
| Run keys | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run` | |
| StartupApproved | `HKCU\...\Explorer\StartupApproved\Run` | Binary data: bit 0 of byte 0 = disabled |
| StartupApproved | `HKLM\...\Explorer\StartupApproved\Run` | Same format |
| **Flagged** | > 10 enabled startup items | |

### Phase 9: Visual & Display Settings
| Check | Source | Details |
|-------|--------|---------|
| Visual effects setting | `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects\VisualFXSetting` | 0=Auto, 1=Appearance, 2=Performance, 3=Custom |
| Transparency | `HKCU\...\Themes\Personalize\EnableTransparency` | |
| Animations | `HKCU\Control Panel\Desktop\WindowMetrics\MinAnimate` | "0" = disabled |

### Phase 10: Antivirus & Security
| Check | Source | Details |
|-------|--------|---------|
| AV product | WMI `root\SecurityCenter2\AntiVirusProduct.displayName` | |
| Defender scan running | Registry `HKLM\SOFTWARE\Microsoft\Windows Defender\Scan\ScanRunning` | Fallback: check `MpCmdRun` process |
| BitLocker | CLI `manage-bde -status C:` | 10s timeout, output parsing |
| **Flagged** | No AV detected (Critical), BitLocker not encrypted (Warning) | |

### Phase 11: Windows Update
| Check | Source | Details |
|-------|--------|---------|
| Pending reboot | Registry `HKLM\...\WindowsUpdate\Auto Update\RebootRequired` | Key existence = reboot needed |
| Service status | `ServiceController("wuauserv")` | Atomic snapshot of Status + StartType |
| Cache size | `%WINDIR%\SoftwareDistribution\Download` | |
| **Flagged** | Service disabled, pending reboot | |

### Phase 12: Network Diagnostics
| Check | Source | Details |
|-------|--------|---------|
| Connection type | WMI `Win32_NetworkAdapter` (NetConnectionStatus=2) | WiFi vs Ethernet |
| WiFi details | CLI `netsh wlan show interfaces` | Signal, band, link speed |
| DNS response | `Dns.GetHostEntry()` with unique GUID subdomains | 3-sample median (avoids caching) |
| VPN detection | WMI network adapters with VPN/TAP/Cisco/GlobalProtect/Fortinet in name | |
| Ping latency | CLI `ping -n 3 8.8.8.8` | Parses "Average = Xms" |
| **Flagged** | DNS > 150ms (poor), DNS 50-150ms (fair) | |

### Phase 13: Internet Speed Test
| Check | Source | Details |
|-------|--------|---------|
| Download | `https://speed.cloudflare.com/__down?bytes=26214400` (25 MB) | `ResponseHeadersRead` for pure throughput timing |
| Upload | `https://speed.cloudflare.com/__up` (5 MB POST) | |
| Warm-up | 1 KB download to establish DNS + TLS | |
| **Flagged** | Download < 10 Mbps | |

### Phase 14: Network Drives
| Check | Source | Details |
|-------|--------|---------|
| Mapped drives | `DriveInfo.GetDrives()` where `DriveType == Network` | |
| UNC path | WMI `Win32_LogicalDisk.ProviderName` | |
| Accessibility + latency | `drive.IsReady` + `Directory.GetDirectories()` timed | |
| Offline Files | `%WINDIR%\CSC` directory existence | WMI `Win32_OfflineFilesCache.Status` |
| **Flagged** | Latency > 200ms, usage > 90%, unreachable | |

### Phase 15: Outlook / Email
| Check | Source | Details |
|-------|--------|---------|
| Installed | Registry `HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\ProductReleaseIds` | Fallback: `HKLM\SOFTWARE\Microsoft\Office` |
| Data files | `%LOCALAPPDATA%\Microsoft\Outlook\*.ost` + `*.pst` | Recursive enumeration |
| Add-ins | `HKCU\SOFTWARE\Microsoft\Office\Outlook\Addins` subkey count | |
| **Flagged** | Data file > 5 GB | |

### Phase 16: Browser Check
| Check | Source | Details |
|-------|--------|---------|
| Tab count (Chrome/Edge) | UI Automation `TreeWalker.ControlViewWalker` | Counts `ControlType.TabItem` at depth ≤ 8, doesn't recurse into tab content |
| Tab count (Firefox) | Process count / 2 (rough estimate) | Fission makes process count unreliable |
| Cache size | File system scan of `Default\Cache` | Chrome/Edge: `%LOCALAPPDATA%\Google\Chrome\User Data`, Edge equivalent |
| Extensions | Directory count in `Default\Extensions` | |
| **Flagged** | > 30 tabs, cache > 200 MB (plan threshold), cache > 1 GB (issue threshold) | |

### Phase 17: User Profile
| Check | Source | Details |
|-------|--------|---------|
| Profile size | `%USERPROFILE%` (depth 2) | |
| Profile age | `NTUSER.DAT` creation time | Fallback: folder creation time |
| Desktop items | `Environment.SpecialFolder.Desktop` file system entries | |
| Corruption | Profile path contains ".bak" or "TEMP" | Temp profile indicator |
| **Flagged** | > 50 desktop items, corruption detected | |

### Phase 18: Office Diagnostics
| Check | Source | Details |
|-------|--------|---------|
| Installed (C2R) | Registry `HKLM\...\Office\ClickToRun\Configuration` | |
| Installed (MSI) | Registry `HKLM\...\Office\{16.0,15.0,14.0}\Common\InstallRoot` | |
| Repair needed | C2R exists but `UpdateChannel` is empty while `VersionToReport` exists | |

### Phase 19: Installed Software
| Check | Source | Details |
|-------|--------|---------|
| Enumeration | Registry `HKLM\...\Uninstall` + `WOW6432Node\...\Uninstall` | Deduped by name |
| EOL detection | Pattern match: Flash, Silverlight, Java 6/7, Python 2, IE, Shockwave, QuickTime, RealPlayer | |
| Bloatware detection | Pattern match: Toolbar, Conduit, Babylon, WildTangent, Candy Crush, etc. | |
| VC++ Runtimes | Matches "Visual C++" + "Redistributable" | |

### Phase 20: Event Log Analysis
| Check | Source | Details |
|-------|--------|---------|
| BSODs | System log, provider `Microsoft-Windows-WER-SystemErrorReporting`, EventID 1001 | Last 30 days, max 10 entries |
| Unexpected shutdowns | System log, provider `EventLog`, EventID 6008 | |
| Disk errors | System log, providers `disk` / `Ntfs` / `volmgr`, Level 1-3 | |
| App crashes | Application log, providers `Application Error` / `Application Hang`, EventID 1000/1002 | |
| Remediation cutoff | Per-category timestamps from `%LOCALAPPDATA%\DLack\event-log-remediated.json` | Events before cutoff are considered already addressed |

### Health Score Calculation

Starting from 100, deductions:

| Category | Condition | Deduction |
|----------|-----------|-----------|
| CPU | Load > 50% | -8 |
| CPU | Load > 15% | -4 |
| RAM | Usage > 95% | -10 |
| RAM | Usage > 85% | -5 |
| RAM | Insufficient (< 8 GB) | -2 |
| Disk | Health not Healthy/Unknown | -8 per drive |
| Disk | > 95% used | -6 per drive |
| Disk | > 85% used | -3 per drive |
| Disk | Activity flagged | -2 |
| Thermal | Throttling | -10 |
| Thermal | Temp flagged | -5 |
| GPU | Temp flagged | -5 |
| GPU | Usage flagged | -2 |
| GPU | Driver outdated | -3 |
| Battery | Health flagged | -5 |
| Security | No AV | -10 |
| Security | No BitLocker | **-7** |
| Uptime | > 14 days | -5 |
| Uptime | > 7 days | -2 |
| Network | Ping > 300ms | -4 |
| Network | Ping > 100ms | -2 |
| Network | Speed flagged (< 5 Mbps) | -4 |
| Network | Speed flagged (< 10 Mbps) | -2 |
| Network Drive | Unreachable | -4 per drive |
| Network Drive | High latency | -2 per drive |
| Network Drive | Space flagged | -2 per drive |
| Windows Update | Pending reboot | **-5** |
| Windows Update | Service disabled | -3 |
| Startup | Too many (> 10) | -3 |
| Software | EOL apps | -2 per app (max -4) |
| Software | Bloatware | -1 |
| User Profile | Corruption | -5 |
| User Profile | Desktop items flagged | -1 |
| Outlook | Large data file | -2 |
| Event Log | BSODs | **-4 per BSOD (max -10)** |
| Event Log | Disk errors ≥ 5 | -4 |
| Event Log | Disk errors 1-4 | -2 |
| Event Log | Unexpected shutdowns ≥ 3 | **-5** |
| Event Log | Unexpected shutdowns 1-2 | **-2** |

**Final score:** `Math.Clamp(score, 0, 100)`

---

## 5. Optimizer Module — Detailed Audit

### Plan Building (Thresholds & Conditions)

| Action Key | Trigger Condition | Risk | Auto-selected |
|------------|-------------------|------|---------------|
| `ClearTempFiles` | Temp > 500 MB AND directories have files | Safe | Yes |
| `ClearPrefetch` | Prefetch > 128 MB AND directory has files | Safe | Yes |
| `EmptyRecycleBin` | Recycle Bin > 100 MB (live P/Invoke check) | Safe | Yes |
| `ClearUpdateCache` | SW Distribution > 500 MB AND has files | Moderate | No |
| `LaunchDiskCleanup` | Windows.old exists AND size > 0 | Moderate | No |
| `CleanUpgradeLogs` | Upgrade logs > 200 MB | Safe | Yes |
| `ClearBrowserCache:{name}` | Browser cache > 200 MB | Safe | Yes |
| `OpenResourceMonitor` | CPU flagged | Safe (Shortcut) | Yes |
| `SetPowerPlanBalanced` | Power plan contains "Power saver" | Safe | Yes |
| `SwitchToPerformanceVisuals` | Transparency OR animations enabled | Safe | No |
| `KillProcess:{pid}:{name}` | Top RAM process > 500 MB | Moderate | No |
| `OpenTaskManagerStartup` | Too many startup items | Safe (Shortcut) | Yes |
| `StartWuauserv` | Windows Update service disabled | Safe | Yes |
| `ScheduleRestart` | Uptime flagged OR pending reboot | RequiresReboot | No |
| `OpenGpuDriverPage:{name}` | GPU driver outdated | Safe (Shortcut) | Yes |
| `OpenAppsSettings` | Office repair needed | Moderate (Shortcut) | No |
| `OpenBitLocker` | BitLocker not encrypted | Moderate (Shortcut) | No |
| `RunSfc` | BSODs detected (or ≥ 5 app crashes without BSODs), **not within 7-day cooldown** | Safe | Yes |
| `RunDism` | BSODs detected, **not within 7-day cooldown** | Safe | Yes |
| `InstallDriverUpdates` | BSODs detected, **not within 7-day cooldown** | Moderate | Yes |
| `ScheduleMemoryDiagnostic` | BSODs detected, **not within 7-day cooldown** | RequiresReboot | No |
| `ClearCrashDumps` | BSODs detected | Safe | Yes |
| `ScheduleChkdsk` | Disk errors detected, **not within 7-day cooldown** | RequiresReboot | No |
| `ReRegisterComponents` | ≥ 5 app crashes, **not within 7-day cooldown** | Safe | Yes |
| `ClearWerReports` | ≥ 5 app crashes | Safe | Yes |
| `DisableFastStartup` | Unexpected shutdowns detected, **not within 7-day cooldown** | Safe | Yes |
| `RepairPowerConfig` | Unexpected shutdowns detected, **not within 7-day cooldown** | Safe | Yes |

### Action Execution Details

#### ClearTempFiles
- **What it does:** Deletes all files in `Path.GetTempPath()` AND `%WINDIR%\Temp` recursively, then removes empty directories bottom-up
- **Implementation:** `Parallel.ForEach` over files, sets `FileAttributes.Normal` before delete, counts deleted/skipped
- **Permissions:** Admin (for Windows Temp)
- **Success detection:** Returns deleted count, skipped count, freed bytes

#### ClearPrefetch
- **What it does:** Deletes all files in `%WINDIR%\Prefetch`
- **Permissions:** Admin required

#### EmptyRecycleBin
- **What it does:** P/Invoke `SHEmptyRecycleBin` with `NOCONFIRMATION | NOPROGRESSUI | NOSOUND`
- **Before/After:** Measures size before and after via `SHQueryRecycleBin`
- **Error:** Throws `InvalidOperationException` if HRESULT < 0

#### ClearUpdateCache
- **What it does:** Stops `wuauserv` → deletes `%WINDIR%\SoftwareDistribution\Download` contents → restarts `wuauserv`
- **Safety:** `finally` block ensures service restart even on failure
- **Permissions:** Admin for service control and file deletion

#### LaunchDiskCleanup (Windows.old removal)
- **What it does:** Sets registry flags for `cleanmgr` sageset profile 65 → runs `cleanmgr /d C: /sagerun:65` → cleans up registry flags
- **Registry:** `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations\StateFlags0065`
- **Timeout:** 15 minutes for cleanmgr
- **Verification:** Checks if `Windows.old` directory still exists after completion

#### ~~SwitchDnsToCloudflare~~ ✅ REMOVED
- **Status:** Removed from codebase. DNS switching and flushing actions were deleted to avoid interfering with corporate DNS infrastructure.

#### SwitchToPerformanceVisuals
- **What it does:** Sets 3 registry values:
  - `HKCU\...\Themes\Personalize\EnableTransparency` → 0
  - `HKCU\Control Panel\Desktop\WindowMetrics\MinAnimate` → "0"
  - `HKCU\...\Explorer\VisualEffects\VisualFXSetting` → 2 (best performance)
  - `HKCU\Control Panel\Desktop\UserPreferencesMask` → performance-oriented binary mask

#### RunSfc
- **What it does:** `cmd /C sfc /scannow > NUL 2>&1`
- **Timeout:** 20 minutes
- **Exit codes:** 0 = success, -1 = timeout, other = check CBS.log

#### RunDism
- **What it does:** `cmd /C DISM /Online /Cleanup-Image /RestoreHealth > NUL 2>&1`
- **Timeout:** 45 minutes

#### ScheduleChkdsk
- **What it does:** `cmd /C echo Y | chkdsk {drive} /F /R`
- **Pipes Y** to handle the confirmation prompt for system drives

#### DisableFastStartup
- **What it does:** Sets `HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power\HiberbootEnabled` → 0
- **Fallback:** `powercfg /hibernate off` if registry access fails

#### RepairPowerConfig
- **What it does:** 5 sequential operations:
  1. `powercfg /restoredefaultschemes` — reset to defaults
  2. Set Balanced plan active
  3. Disable wake timers (AC + DC)
  4. Disable USB selective suspend (AC + DC)
  5. Disable PCI Express power management (AC + DC)

#### ReRegisterComponents
- **What it does:** 4 operations:
  1. `DISM /Online /Cleanup-Image /StartComponentCleanup`
  2. Clear Windows Store cache directory
  3. Rebuild font cache (stop FontCache → delete cache dir → start FontCache)
  4. `regsvr32 /s` for 5 core DLLs (oleaut32, ole32, mshtml, urlmon, shell32)

#### InstallDriverUpdates
- **What it does:** `pnputil /scan-devices` → `UsoClient StartInstall` (fallback: `wuauclt /detectnow /updatenow`)

---

## 6. Verification Module

After optimization completes, `VerifyActionsAsync` re-checks each affected item:

| Action | Verification Method | Verified Threshold |
|--------|--------------------|--------------------|
| ClearTempFiles | Re-measure temp folder size | < 100 MB |
| ClearPrefetch | Re-measure prefetch size | < 50 MB |
| EmptyRecycleBin | `SHQueryRecycleBin` | < 10 MB |
| ClearUpdateCache | Re-measure SW Distribution | < 50 MB |
| CleanUpgradeLogs | Re-measure all upgrade log dirs | < 50 MB |
| LaunchDiskCleanup | Check if `Windows.old` exists | Directory gone |
| KillProcess | `Process.GetProcessById` + `HasExited` | Process gone |
| SetPowerPlanBalanced | `powercfg /getactivescheme` contains Balanced GUID | Balanced plan active |
| SwitchToPerformanceVisuals | Read `EnableTransparency` registry | Value is 0 |
| StartWuauserv | `ServiceController.Status` | Running |
| DisableFastStartup | Read `HiberbootEnabled` registry | Value is 0 |
| ClearCrashDumps | Check Minidump dir + MEMORY.DMP | Both empty/gone |
| ScheduleChkdsk/MemDiag/Restart | — | `RequiresReboot` |
| RunSfc/RunDism/ReRegister | Trust action exit code | Based on Status |

---

## 7. PDF Report Generator

- **Library:** QuestPDF 2025.12.4 (Community license)
- **Layout:** A4, 2cm margins
- **Branding:** Pillsbury logo (from `Assets/pillsbury-logo-color.png`), sienna accent (#9B4D2B)
- **Sections:** Executive Summary → System Overview → CPU → RAM → GPU → Disk → Battery → Startup → Visual Settings → Antivirus → Windows Update → Network → Network Drives → Outlook → Browsers → User Profile → Office → Installed Software → Event Log → Recommendations → Optimization Plan/Results
- **Output:** Saved to Desktop as `EDR_Diagnostic_{ComputerName}_{timestamp}.pdf`, auto-opens in default PDF viewer

---

## 8. Theme & UI System

- **Persistence:** `%LOCALAPPDATA%\DLack\theme.json` (`{"DarkMode": true/false}`)
- **XAML resources:** `Themes/ProfessionalLight.xaml`, `Themes/ProfessionalDark.xaml`
- **Static colors:** Set on `ThemeManager` static properties, consumed by code-behind UI construction
- **All brushes frozen** for thread safety
- **Dynamic resources:** Button styles (`BluePrimaryButton`, `NeutralButton`, `ChromeCloseButton`), background/border brushes
- **Custom chrome:** `WindowChrome` with `CaptionHeight=40`, rounded corners, maximize/restore handling with proper taskbar respect via `WM_GETMINMAXINFO` hook

---

## 9. Bugs, Risks & Inconsistencies

### Bug #1: ~~HighCpuProcess is always null~~ ✅ FIXED
**Status:** Dead code removed — `HighCpuProcess` property deleted from `CpuDiagnostics`, dead `BuildPlan()` branch removed, health score deduction removed, PDF reference removed.

### Bug #2: ~~Hardcoded C:\ paths in Scanner~~ ✅ FIXED
**Status:** `ScanDiskDiagnostics()` and `ScanWindowsUpdate()` now use `Environment.GetFolderPath(Environment.SpecialFolder.Windows)` and `Path.GetPathRoot()` consistently.

### Bug #3: Firefox tab count is unreliable
**File:** `Scanner.cs`, `ScanFirefox()`  
**Issue:** `processCount / 2` is a "rough lower-bound estimate." With Firefox Fission, process count bears no relation to tab count.  
**Impact:** Low — informational only, but the "> 30 tabs" action will rarely trigger for Firefox.  
**Status:** Accepted — documenting limitation. Full fix would require parsing `recovery.jsonlz4`.

### Bug #4: ~~Browser cache path mismatch between Scanner and Optimizer~~ ✅ FIXED
**Status:** Scanner now measures both `Default\Cache` AND `Default\Code Cache` for Chrome/Edge, matching what the Optimizer clears.

### Bug #5: ~~Firefox cache path mismatch~~ ✅ ALREADY CORRECT
**Status:** Verified — `Optimizer.cs` `ClearBrowserCache("Firefox")` already uses `Environment.SpecialFolder.ApplicationData` (Roaming). No change needed.

### Bug #6: DNS measurement uses NXDOMAIN responses
**File:** `Scanner.cs`, DNS response measurement  
**Issue:** Uses random GUID subdomains of `example.com` to avoid caching, but these will always return NXDOMAIN. This measures the DNS resolver's NXDOMAIN response time, which may differ from successful resolution time.  
**Impact:** Low — the measurement is consistent and still reflects DNS resolver performance, but may show slightly different times than real-world browsing.  
**Status:** Accepted — behavior is consistent and non-harmful.

### Risk #1: SHEmptyRecycleBin without user confirmation
**File:** `Optimizer.cs`  
**Detail:** The Recycle Bin is emptied with `SHERB_NOCONFIRMATION`. Users select this in the optimization plan, but there's no separate "are you sure?" for this specific action.  
**Mitigation:** The action is marked Safe and auto-selected, which is appropriate since users explicitly chose to run it.

### Risk #2: ~~Power plan reset is destructive~~ ✅ MITIGATED
**Status:** Action description now warns users that custom power plans will be deleted.

### Risk #3: UserPreferencesMask hardcoded binary
**File:** `Optimizer.cs`, `SwitchToPerformanceVisuals()`  
**Detail:** Writes a hardcoded 8-byte `UserPreferencesMask` binary value. This controls many UI behaviors beyond transparency/animations. Different Windows versions may interpret these bytes differently.  
**Impact:** Low-Medium — could cause unexpected UI behavior on some Windows versions.  
**Status:** Accepted — action is not auto-selected and is labeled as a visual optimization.

### Inconsistency #1: ~~`GetDirectorySizeMB` / `GetFolderSizeMB` duplicated~~ ✅ FIXED
**Status:** Both now delegate to `NativeHelpers.GetFolderSizeMB()`.

### Inconsistency #2: ~~`SHQueryRecycleBin` duplicated~~ ✅ FIXED
**Status:** Both Scanner and Optimizer now delegate to `NativeHelpers.GetRecycleBinSizeMB()`. P/Invoke declaration lives only in `NativeHelpers.cs`.

### Inconsistency #3: ~~Empty AI files~~ ✅ FIXED
**Status:** `AiHelper.cs`, `AiEngine.cs`, `AiChatWindow.xaml.cs` removed from project.

---

## 10. Concrete Fixes & Recommendations

### P0 — Must Fix Before Deployment

1. ~~**Fix Firefox cache clearing path**~~ ✅ Already correct — verified at code level.

2. ~~**Fix hardcoded `C:\` paths in Scanner**~~ ✅ FIXED — now uses `Environment.GetFolderPath` and `Path.GetPathRoot()`.

3. ~~**Remove or implement AI files**~~ ✅ FIXED — `AiHelper.cs`, `AiEngine.cs`, `AiChatWindow.xaml.cs` removed.

### P1 — Should Fix

4. ~~**Remove dead HighCpuProcess code**~~ ✅ FIXED — property removed from `CpuDiagnostics`, dead `BuildPlan()` branch removed, health score/PDF references removed.

5. ~~**Warn about custom power plan deletion**~~ ✅ FIXED — `RepairPowerConfig` action description now warns users.

6. ~~**Deduplicate shared code**~~ ✅ FIXED — `NativeHelpers.cs` created with shared `GetRecycleBinSizeMB()` and `GetFolderSizeMB()`. Both Scanner and Optimizer delegate to it.

7. ~~**Include Code Cache in Scanner measurement**~~ ✅ FIXED — Scanner now measures both `Cache` and `Code Cache` for Chrome/Edge.

8. ~~**Remove DNS optimization actions**~~ ✅ FIXED — `SwitchDnsCloudflare` and `FlushDns` removed from `BuildPlan()`, dispatch, verification, and execution methods. DNS is measured for diagnostics only.

9. ~~**Make Internet Speed Test skippable**~~ ✅ FIXED — `Scanner.SkipSpeedTest` defaults to `true`. The speed test is skipped unless explicitly enabled. The 5 MB upload resembles data exfiltration to DLP/EDR tools.

### P2 — Nice to Have (remaining)

10. **Improve Firefox tab count**: Parse `recovery.jsonlz4` or use Marionette/Remote Debugging Protocol for accurate counts.

11. **Add logging to file**: Currently all logs go to the UI `RichTextBox`. Consider also writing to a log file for support scenarios.

12. **Add error telemetry**: Silent `catch {}` blocks throughout the codebase make debugging difficult. Consider structured logging.

13. **Consider `CancellationToken` propagation** in optimization actions. Currently, `Task.Run(() => DispatchAction(...), ct)` only checks cancellation between actions, not within long-running ones like `RunSfc`.

---

## 11. Security Audit

### Positive Findings

✅ **Admin enforcement:** Both manifest (`requireAdministrator`) and runtime check (`WindowsPrincipal.IsInRole`)  
✅ **Command injection prevention:** `ValidateShellArgument()` rejects dangerous characters before shell interpolation  
✅ **No secrets or credentials** in codebase  
✅ **System-critical process protection:** Kill action excludes System, svchost, csrss, wininit, services, lsass, smss, dwm, explorer  
✅ **Service restart guarantee:** `ClearUpdateCache()` uses `finally` to always restart `wuauserv`  
✅ **P/Invoke signatures correct:** SHQUERYRBINFO uses proper `Pack=4`, MEMORYSTATUSEX initializes `dwLength`

### Concerns

⚠ **`RunElevatedCmd` uses `cmd /C` with string interpolation.** The `ValidateShellArgument()` helper is used for `driveLetter` in chkdsk. Other interpolated values (like `cleanmgr` paths, `regsvr32` DLL names) are either constants or validated by context, but this pattern is fragile.

⚠ **Broad file deletion:** `DeleteDirectoryContents()` with `RecurseSubdirectories = true` and `AttributesToSkip = 0` will delete hidden/system files. This is intentional for temp cleanup but could be dangerous if pointed at wrong directories.

⚠ **QuestPDF Community License:** Verify that your usage qualifies for the community (free) license. QuestPDF requires a paid license for companies with annual gross revenue exceeding $1M.

### Known Limitations

✅ **WMI timeout cascading — MITIGATED:** Per-phase timeouts are now tailored to WMI risk. Phases that query hang-prone namespaces get shorter timeouts: Thermal Diagnostics (`root\WMI`) = 15s, Disk Diagnostics (`root\microsoft\windows\storage`) = 30s, Internet Speed Test = 45s, all others = 60s. Worst-case scan duration reduced from 20 minutes to ~13 minutes. Implemented via `GetPhaseTimeoutMs()` in `Scanner.cs`.

✅ **Outlook PST/OST file locking — HARDENED:** Phase 15 now catches `IOException` and `UnauthorizedAccessException` explicitly and logs the affected filename. `FileInfo.Length` reads NTFS metadata (not the file stream), so it succeeds even when Outlook holds an exclusive lock — but if access fails for any reason, the error is now visible in the scan log instead of silently swallowed.

✅ **Multi-user scenarios — SURFACED:** `DiagnosticResult.ScannedUser` now records `DOMAIN\Username` at scan start. The scan log shows which user context was used. The PDF report includes the scanned user in the Executive Summary. This makes it explicit that results reflect only one user's HKCU/LOCALAPPDATA.

✅ **Health score weighting — REBALANCED:** Increased deductions for security-critical items: BitLocker not encrypted (-4 → **-7**), pending reboot (-3 → **-5**), BSODs (-3/event cap -6 → **-4/event cap -10**), unexpected shutdowns ≥3 (-3 → **-5**), 1-2 shutdowns (-1 → **-2**). The same test scenario (no BitLocker + stale driver + 11 startup items + pending reboot) now scores **78** instead of 87. Two BSODs deduct **-8** instead of -6.

📋 **Internet Speed Test and DLP/EDR risk:** The 5 MB upload to `speed.cloudflare.com` looks identical to data exfiltration from the perspective of corporate security tools (CrowdStrike, Zscaler, Netskope). SSL inspection proxies may also interfere with measurements or block the request entirely. The speed test is the least actionable diagnostic — IT manages the network. `Scanner.SkipSpeedTest` defaults to `true` (speed test is **off** by default). Set it to `false` only when actively troubleshooting a specific user's connectivity complaint.

📋 **DNS: measurement only, no changes.** The scanner measures DNS response time for diagnostic purposes. The optimizer does **not** modify DNS settings — DNS switching and flushing actions have been removed to avoid interfering with corporate DNS infrastructure (split-horizon DNS, internal domains, security policies). The DNS flagged issue recommendation was changed from "Switch to Cloudflare/Google DNS" to "contact IT if browsing feels sluggish" — consistent with the "measure only" policy.

✅ **Remediation cooldown — NO REPEAT FIXES:** Event-log remediation actions (SFC, DISM, chkdsk, ReRegisterComponents, DisableFastStartup, RepairPowerConfig) now have a 7-day cooldown. If the same category was remediated within 7 days and new events still appear, the automated fix is replaced with a manual escalation card (e.g., "⚠ Disk errors persist after chkdsk — the drive may be failing"). This prevents the "Groundhog Day" loop where the same ineffective fix runs every scan. Disk cleanup actions (ClearCrashDumps, ClearWerReports) are exempt — they always run since they reclaim space regardless.

✅ **Disk cleanup thresholds — RAISED:** Temp file threshold raised from 100 MB → **500 MB**, prefetch from 50 MB → **128 MB**. The previous thresholds triggered on every reboot since Windows regenerates ~200-400 MB of temp data during startup. The new thresholds only propose cleanup when accumulation is genuinely excessive.

---

- [x] Fix P0 bugs listed in Section 10
- [x] Remove empty AI files
- [x] Deduplicate shared helpers into `NativeHelpers.cs`
- [x] Remove DNS optimization actions (SwitchDnsCloudflare, FlushDns)
- [x] Make Internet Speed Test off by default (`Scanner.SkipSpeedTest = true`)
- [ ] Set `SkipSpeedTest = false` only for targeted connectivity troubleshooting
- [ ] Verify `Assets/pillsbury-logo-white.png` and `Assets/pillsbury-logo-color.png` are included as `Resource`
- [ ] Test on non-C: Windows installations
- [ ] Test with no internet connection (speed test should gracefully skip or fail)
- [ ] Test with no battery (battery section should be hidden)
- [ ] Test with no mapped network drives
- [ ] Test with no Outlook installed
- [ ] Test with no GPU (VM scenario)
- [ ] Verify QuestPDF license compliance
- [ ] Code-sign the executable (since it requires admin elevation)
- [ ] Test on Windows 10 and Windows 11
- [ ] Verify PDF opens correctly on machines without a PDF reader (graceful fallback)
- [ ] Test cancellation during scan (Esc key, Cancel button)
- [ ] Test cancellation during optimization
- [ ] Verify theme persistence across app restarts
- [ ] Test with multiple monitors / high DPI

---

*End of Audit Report*
