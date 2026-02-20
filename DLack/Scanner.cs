using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using Microsoft.Win32;

namespace DLack
{
    public class Scanner : IDisposable
    {
        // ── Events ───────────────────────────────────────────────────
        public event Action<ScanProgress> OnProgress;
        public event Action<string> OnLog;

        // ── State ────────────────────────────────────────────────────
        private CancellationTokenSource _cts;
        private bool _disposed;
        private List<ProcessInfo> _topProcesses;
        private double _cachedMaxClockMHz;
        private double _cachedCurrentClockMHz;
        private List<double> _cachedThermalZonesC;
        private CancellationToken _phaseToken;

        /// <summary>
        /// When true, the Internet Speed Test phase is skipped.
        /// Defaults to true — the 5 MB upload to speed.cloudflare.com resembles data exfiltration
        /// to DLP/EDR tools and may trigger security alerts in corporate environments.
        /// Set to false only when actively troubleshooting a specific user's connectivity.
        /// </summary>
        public bool SkipSpeedTest { get; set; } = true;

        // ═══════════════════════════════════════════════════════════════
        //  PUBLIC API
        // ═══════════════════════════════════════════════════════════════

        public async Task<DiagnosticResult> RunScan()
        {
            _cts?.Dispose();
            _cts = new CancellationTokenSource();

            var result = new DiagnosticResult();
            var stopwatch = Stopwatch.StartNew();

            var phaseList = new List<(string Name, Func<DiagnosticResult, Task> Action)>
            {
                ("System Overview",         r => Task.Run(() => ScanSystemOverview(r))),
                ("CPU Diagnostics",         r => Task.Run(() => ScanCpuDiagnostics(r))),
                ("RAM Diagnostics",         r => Task.Run(() => ScanRamDiagnostics(r))),
                ("Thermal Diagnostics",     r => Task.Run(() => ScanThermalDiagnostics(r))),
                ("Disk Diagnostics",        r => Task.Run(() => ScanDiskDiagnostics(r))),
                ("GPU Diagnostics",         r => Task.Run(() => ScanGpuDiagnostics(r))),
                ("Battery Health",          r => Task.Run(() => ScanBatteryHealth(r))),
                ("Startup Programs",        r => Task.Run(() => ScanStartupPrograms(r))),
                ("Visual & Display",        r => Task.Run(() => ScanVisualSettings(r))),
                ("Antivirus & Security",    r => Task.Run(() => ScanAntivirus(r))),
                ("Windows Update",          r => Task.Run(() => ScanWindowsUpdate(r))),
                ("Network Diagnostics",     r => Task.Run(() => ScanNetwork(r))),
            };

            if (!SkipSpeedTest)
                phaseList.Add(("Internet Speed Test", r => ScanSpeedTest(r)));

            phaseList.AddRange(new (string, Func<DiagnosticResult, Task>)[]
            {
                ("Network Drives",          r => Task.Run(() => ScanNetworkDrives(r))),
                ("Outlook / Email",         r => Task.Run(() => ScanOutlook(r))),
                ("Browser Check",           r => Task.Run(() => ScanBrowsers(r))),
                ("User Profile",            r => Task.Run(() => ScanUserProfile(r))),
                ("Office Diagnostics",      r => Task.Run(() => ScanOffice(r))),
                ("Installed Software",      r => Task.Run(() => ScanInstalledSoftware(r))),
                ("Event Log Analysis",      r => Task.Run(() => ScanEventLogs(r))),
            });

            var phases = phaseList.ToArray();

            // Record which user context this scan reflects
            result.ScannedUser = $@"{Environment.UserDomainName}\{Environment.UserName}";
            SafeLog($"Starting comprehensive system diagnostic scan as {result.ScannedUser}...");

            // Snapshot top processes once — used by both CPU and RAM phases
            _topProcesses = GetTopProcessesByMemory(10);

            bool cancelled = false;
            for (int i = 0; i < phases.Length; i++)
            {
                if (_cts.Token.IsCancellationRequested)
                {
                    SafeLog("Scan cancelled.");
                    cancelled = true;
                    break;
                }

                var phase = phases[i];

                SafeProgress(new ScanProgress
                {
                    PhaseIndex = i + 1,
                    Total = phases.Length,
                    CurrentPhase = phase.Name,
                    PhaseDescription = GetPhaseDescription(phase.Name),
                    EstimatedSeconds = GetPhaseEstimate(phase.Name)
                });

                SafeLog($"[{i + 1}/{phases.Length}] {phase.Name}...");

                // Per-phase cancellation: cancelled on timeout so the phase stops mutating result
                var phaseCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token);
                _phaseToken = phaseCts.Token;

                try
                {
                    var phaseTask = phase.Action(result);
                    int phaseTimeoutMs = GetPhaseTimeoutMs(phase.Name);
                    if (await Task.WhenAny(phaseTask, Task.Delay(phaseTimeoutMs, _cts.Token)) == phaseTask)
                    {
                        await phaseTask; // propagate exceptions
                        SafeLog($"  ✓ {phase.Name} complete");
                    }
                    else
                    {
                        phaseCts.Cancel(); // Stop the timed-out phase from further mutating result
                        SafeLog($"  ⚠ {phase.Name} timed out ({phaseTimeoutMs / 1000}s) — skipping");
                    }
                }
                catch (OperationCanceledException) { cancelled = true; break; }
                catch (Exception ex)
                {
                    SafeLog($"  ⚠ {phase.Name} error: {ex.Message}");
                }
            }

            if (cancelled)
                return null;

            // Build aggregated flagged issues
            BuildFlaggedIssues(result);
            CalculateHealthScore(result);

            stopwatch.Stop();
            result.ScanDurationSeconds = Math.Round(stopwatch.Elapsed.TotalSeconds, 1);

            SafeLog($"Scan complete in {result.ScanDurationSeconds}s. {result.FlaggedIssues.Count} issue(s) flagged.");
            return result;
        }

        public void CancelScan()
        {
            try { _cts?.Cancel(); }
            catch (ObjectDisposedException) { }
        }

        // ═══════════════════════════════════════════════════════════════
        //  1. SYSTEM OVERVIEW
        // ═══════════════════════════════════════════════════════════════

        private void ScanSystemOverview(DiagnosticResult result)
        {
            var so = result.SystemOverview;
            so.ComputerName = Environment.MachineName;

            try
            {
                using var cs = new ManagementObjectSearcher("SELECT Manufacturer, Model FROM Win32_ComputerSystem");
                cs.Options.Timeout = TimeSpan.FromSeconds(10);
                foreach (var obj in cs.Get())
                {
                    so.Manufacturer = obj["Manufacturer"]?.ToString() ?? "";
                    so.Model = obj["Model"]?.ToString() ?? "";
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ ComputerSystem: {ex.Message}"); }

            try
            {
                using var bios = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS");
                bios.Options.Timeout = TimeSpan.FromSeconds(10);
                foreach (var obj in bios.Get())
                    so.SerialNumber = obj["SerialNumber"]?.ToString() ?? "";
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ BIOS: {ex.Message}"); }

            try
            {
                using var os = new ManagementObjectSearcher("SELECT Caption, BuildNumber FROM Win32_OperatingSystem");
                os.Options.Timeout = TimeSpan.FromSeconds(10);
                foreach (var obj in os.Get())
                {
                    so.WindowsVersion = obj["Caption"]?.ToString() ?? "";
                    so.WindowsBuild = obj["BuildNumber"]?.ToString() ?? "";
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ OperatingSystem: {ex.Message}"); }

            try
            {
                using var cpu = new ManagementObjectSearcher("SELECT Name, MaxClockSpeed, CurrentClockSpeed FROM Win32_Processor");
                cpu.Options.Timeout = TimeSpan.FromSeconds(10);
                foreach (var obj in cpu.Get())
                {
                    so.CpuModel = obj["Name"]?.ToString()?.Trim() ?? "";
                    double maxClock = Convert.ToDouble(obj["MaxClockSpeed"] ?? 0);
                    double currentClock = Convert.ToDouble(obj["CurrentClockSpeed"] ?? 0);
                    so.CpuClockSpeed = $"{maxClock:F0} MHz";

                    // Cache for throttle detection in ScanThermalDiagnostics
                    _cachedMaxClockMHz = maxClock;
                    _cachedCurrentClockMHz = currentClock;
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ Processor: {ex.Message}"); }

            var mem = new MEMORYSTATUSEX();
            if (GlobalMemoryStatusEx(mem))
                so.TotalRamMB = (long)(mem.ullTotalPhys / 1024 / 1024);

            so.Uptime = TimeSpan.FromMilliseconds(Environment.TickCount64);
            so.UptimeFlagged = so.Uptime.TotalDays > 7;
        }

        // ═══════════════════════════════════════════════════════════════
        //  2. CPU DIAGNOSTICS
        // ═══════════════════════════════════════════════════════════════

        private void ScanCpuDiagnostics(DiagnosticResult result)
        {
            var ct = _phaseToken;
            var cpu = result.Cpu;

            // Multi-sample CPU load for stable readings (3 samples, 500ms apart)
            try
            {
                var samples = new List<double>();
                for (int i = 0; i < 3; i++)
                {
                    if (ct.IsCancellationRequested) break;
                    if (i > 0) Thread.Sleep(500);
                    using var searcher = new ManagementObjectSearcher(
                        "SELECT LoadPercentage FROM Win32_Processor");
                    searcher.Options.Timeout = TimeSpan.FromSeconds(5);
                    foreach (var obj in searcher.Get())
                    {
                        double load = Convert.ToDouble(obj["LoadPercentage"] ?? 0);
                        samples.Add(load);
                        break;
                    }
                }
                cpu.CpuLoadPercent = samples.Count > 0
                    ? Math.Round(samples.Average(), 1) : 0;
            }
            catch (Exception ex) { cpu.CpuLoadPercent = 0; OnLog?.Invoke($"  ⚡ CPU load: {ex.Message}"); }

            cpu.CpuLoadFlagged = cpu.CpuLoadPercent > 30;

            // Top 10 processes by memory (reliable proxy)
            cpu.TopCpuProcesses = _topProcesses;
        }

        // ═══════════════════════════════════════════════════════════════
        //  3. RAM DIAGNOSTICS
        // ═══════════════════════════════════════════════════════════════

        private void ScanRamDiagnostics(DiagnosticResult result)
        {
            var ram = result.Ram;

            var mem = new MEMORYSTATUSEX();
            if (GlobalMemoryStatusEx(mem))
            {
                ram.TotalMB = (long)(mem.ullTotalPhys / 1024 / 1024);
                ram.AvailableMB = (long)(mem.ullAvailPhys / 1024 / 1024);
                ram.UsedMB = ram.TotalMB - ram.AvailableMB;
                ram.PercentUsed = ram.TotalMB > 0
                    ? Math.Round((double)ram.UsedMB / ram.TotalMB * 100, 1)
                    : 0;
            }

            ram.UsageFlagged = ram.PercentUsed > 85;
            ram.InsufficientRam = ram.TotalMB < 8192;

            ram.TopRamProcesses = _topProcesses;

            // Check Windows Memory Diagnostic status
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsMemoryDiagnostic");
                ram.MemoryDiagnosticStatus = key?.GetValue("LastResult")?.ToString() ?? "No results available";
            }
            catch { ram.MemoryDiagnosticStatus = "Unable to check"; }
        }

        // ═══════════════════════════════════════════════════════════════
        //  4. DISK DIAGNOSTICS
        // ═══════════════════════════════════════════════════════════════

        private void ScanDiskDiagnostics(DiagnosticResult result)
        {
            var ct = _phaseToken;
            var disk = result.Disk;

            // List all drives
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (ct.IsCancellationRequested) return;
                if (!drive.IsReady) continue;
                try
                {
                    long totalMB = drive.TotalSize / 1024 / 1024;
                    long freeMB = drive.AvailableFreeSpace / 1024 / 1024;
                    long usedMB = totalMB - freeMB;
                    double pctUsed = totalMB > 0 ? Math.Round((double)usedMB / totalMB * 100, 1) : 0;

                    var di = new DriveInfoResult
                    {
                        DriveLetter = drive.Name.TrimEnd('\\'),
                        TotalMB = totalMB,
                        UsedMB = usedMB,
                        FreeMB = freeMB,
                        PercentUsed = pctUsed,
                        UsageFlagged = pctUsed > 85,
                        DriveType = GetDriveMediaType(drive.Name)
                    };

                    // Health status via WMI
                    di.HealthStatus = GetDriveHealthStatus(drive.Name);

                    disk.Drives.Add(di);
                }
                catch (Exception ex) { OnLog?.Invoke($"  ⚡ Drive {drive.Name}: {ex.Message}"); }
            }

            // Disk activity via WMI
            try
            {
                using var diskSearch = new ManagementObjectSearcher(
                    "SELECT PercentDiskTime FROM Win32_PerfFormattedData_PerfDisk_PhysicalDisk WHERE Name='_Total'");
                diskSearch.Options.Timeout = TimeSpan.FromSeconds(5);
                foreach (var obj in diskSearch.Get())
                {
                    disk.DiskActivityPercent = Convert.ToDouble(obj["PercentDiskTime"] ?? 0);
                    break;
                }
                disk.DiskActivityFlagged = disk.DiskActivityPercent > 50;
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ Disk activity: {ex.Message}"); }

            if (ct.IsCancellationRequested) return;

            // Folder sizes (limit depth for potentially huge folders)
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string sysRoot = Path.GetPathRoot(winDir) ?? @"C:\";

            disk.WindowsTempMB = GetFolderSizeMB(Path.Combine(winDir, "Temp"), 3);
            disk.UserTempMB = GetFolderSizeMB(Path.GetTempPath(), 3);
            disk.SoftwareDistributionMB = GetFolderSizeMB(Path.Combine(winDir, "SoftwareDistribution"), 4);
            disk.PrefetchMB = GetFolderSizeMB(Path.Combine(winDir, "Prefetch"), 1);

            // Windows.old
            string winOld = Path.Combine(sysRoot, "Windows.old");
            disk.WindowsOldExists = Directory.Exists(winOld);
            if (disk.WindowsOldExists)
                disk.WindowsOldMB = GetFolderSizeMB(winOld, 3);

            // Windows upgrade logs (Panther, upgrade logs, etc.)
            long upgradeLogs = 0;
            upgradeLogs += GetFolderSizeMB(Path.Combine(winDir, "Logs", "WindowsUpdate"), 2);
            upgradeLogs += GetFolderSizeMB(Path.Combine(winDir, "Panther"), 2);
            upgradeLogs += GetFolderSizeMB(Path.Combine(sysRoot, "$Windows.~BT"), 3);
            upgradeLogs += GetFolderSizeMB(Path.Combine(sysRoot, "$Windows.~WS"), 3);
            disk.UpgradeLogsMB = upgradeLogs;

            // Recycle Bin size
            disk.RecycleBinMB = GetRecycleBinSizeMB();
        }

        // ═══════════════════════════════════════════════════════════════
        //  5. THERMAL DIAGNOSTICS (merged into CPU)
        // ═══════════════════════════════════════════════════════════════

        private void ScanThermalDiagnostics(DiagnosticResult result)
        {
            var cpu = result.Cpu;

            // Query all thermal zones once and cache for GPU fallback
            try
            {
                _cachedThermalZonesC = new List<double>();
                using var searcher = new ManagementObjectSearcher(
                    @"root\WMI", "SELECT CurrentTemperature FROM MSAcpi_ThermalZoneTemperature");
                searcher.Options.Timeout = TimeSpan.FromSeconds(5);
                foreach (var obj in searcher.Get())
                {
                    double kelvinTenths = Convert.ToDouble(obj["CurrentTemperature"]);
                    double celsius = Math.Round((kelvinTenths / 10.0) - 273.15, 1);
                    if (celsius > 0 && celsius < 120)
                        _cachedThermalZonesC.Add(celsius);
                }
                // First zone is typically CPU
                if (_cachedThermalZonesC.Count > 0)
                    cpu.CpuTemperatureC = _cachedThermalZonesC[0];
            }
            catch (Exception ex) { cpu.CpuTemperatureC = 0; OnLog?.Invoke($"  ⚡ Thermal: {ex.Message}"); }

            cpu.TemperatureFlagged = cpu.CpuTemperatureC > 80;

            // Throttle detection — only flag if clock is significantly reduced AND temperature is elevated
            // Modern CPUs dynamically lower clock speed when idle — that is power management, not throttling
            if (_cachedMaxClockMHz > 0)
            {
                cpu.IsThrottling = _cachedCurrentClockMHz < _cachedMaxClockMHz * 0.5 && cpu.CpuTemperatureC > 70;
            }
            else
            {
                // Fallback: query only if SystemOverview didn't run first
                try
                {
                    using var searcher = new ManagementObjectSearcher("SELECT CurrentClockSpeed, MaxClockSpeed FROM Win32_Processor");
                    searcher.Options.Timeout = TimeSpan.FromSeconds(10);
                    foreach (var obj in searcher.Get())
                    {
                        double current = Convert.ToDouble(obj["CurrentClockSpeed"]);
                        double max = Convert.ToDouble(obj["MaxClockSpeed"]);
                        cpu.IsThrottling = max > 0 && current < max * 0.5 && cpu.CpuTemperatureC > 70;
                    }
                }
                catch (Exception ex) { OnLog?.Invoke($"  ⚡ Throttle check: {ex.Message}"); }
            }

            // Fan status
            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Fan");
                searcher.Options.Timeout = TimeSpan.FromSeconds(5);
                var fans = searcher.Get();
                cpu.FanStatus = fans.Count > 0 ? $"{fans.Count} fan(s) detected" : "No fan data available";
            }
            catch { cpu.FanStatus = "Unable to query"; }
        }

        // ═══════════════════════════════════════════════════════════════
        //  GPU DIAGNOSTICS
        // ═══════════════════════════════════════════════════════════════

        private void ScanGpuDiagnostics(DiagnosticResult result)
        {
            var ct = _phaseToken;
            var gpu = result.Gpu;
            var allGpus = new List<GpuInfo>();

            // ── Enumerate ALL GPUs from WMI ──
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    "SELECT Name, DriverVersion, DriverDate, AdapterRAM, Status, " +
                    "CurrentHorizontalResolution, CurrentVerticalResolution, CurrentRefreshRate " +
                    "FROM Win32_VideoController");
                searcher.Options.Timeout = TimeSpan.FromSeconds(10);

                foreach (var obj in searcher.Get())
                {
                    var info = new GpuInfo();
                    info.Name = obj["Name"]?.ToString() ?? "Unknown";
                    info.DriverVersion = obj["DriverVersion"]?.ToString() ?? "Unknown";
                    info.AdapterStatus = obj["Status"]?.ToString() ?? "Unknown";

                    // Parse driver date
                    string rawDate = obj["DriverDate"]?.ToString() ?? "";
                    if (rawDate.Length >= 8)
                    {
                        try
                        {
                            var dt = DateTime.ParseExact(rawDate.Substring(0, 8), "yyyyMMdd",
                                System.Globalization.CultureInfo.InvariantCulture);
                            info.DriverDate = dt.ToString("yyyy-MM-dd");
                            info.DriverOutdated = (DateTime.Now - dt).TotalDays > 365;
                        }
                        catch { info.DriverDate = rawDate.Substring(0, 8); }
                    }

                    // VRAM from WMI (UInt32 — may overflow for >4 GB)
                    long adapterRam = Convert.ToInt64(obj["AdapterRAM"] ?? 0);
                    if (adapterRam > 0 && adapterRam < 0xFFF00000L)
                        info.DedicatedVideoMemoryMB = adapterRam / (1024 * 1024);
                    else if (adapterRam >= 0xFFF00000L)
                        info.DedicatedVideoMemoryMB = 4096; // UInt32 overflow — at least 4 GB

                    // Resolution & refresh rate (only the active display adapter has these)
                    int hRes = Convert.ToInt32(obj["CurrentHorizontalResolution"] ?? 0);
                    int vRes = Convert.ToInt32(obj["CurrentVerticalResolution"] ?? 0);
                    if (hRes > 0 && vRes > 0)
                        info.Resolution = $"{hRes} x {vRes}";
                    int refreshRate = Convert.ToInt32(obj["CurrentRefreshRate"] ?? 0);
                    if (refreshRate > 0)
                        info.RefreshRate = $"{refreshRate} Hz";

                    allGpus.Add(info);
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ GPU enumeration: {ex.Message}"); }

            if (allGpus.Count == 0)
                return;

            if (ct.IsCancellationRequested) return;

            // ── Get accurate VRAM from registry for each GPU ──
            try
            {
                var registryVram = GetAllGpuVramFromRegistry();
                foreach (var gpuInfo in allGpus)
                {
                    foreach (var rv in registryVram)
                    {
                        if (gpuInfo.Name.Contains(rv.Name, StringComparison.OrdinalIgnoreCase) ||
                            rv.Name.Contains(gpuInfo.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            if (rv.VramMB > gpuInfo.DedicatedVideoMemoryMB)
                                gpuInfo.DedicatedVideoMemoryMB = rv.VramMB;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ GPU VRAM registry: {ex.Message}"); }

            // ── Primary GPU: prefer dedicated
            GpuInfo primary = null;
            foreach (var g in allGpus)
            {
                bool isDedicated = g.Name.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase) ||
                                   g.Name.Contains("GeForce", StringComparison.OrdinalIgnoreCase) ||
                                   g.Name.Contains("Radeon", StringComparison.OrdinalIgnoreCase) ||
                                   g.Name.Contains("AMD", StringComparison.OrdinalIgnoreCase) ||
                                   g.Name.Contains("Quadro", StringComparison.OrdinalIgnoreCase) ||
                                   g.Name.Contains("Arc A", StringComparison.OrdinalIgnoreCase);
                if (isDedicated) { primary = g; break; }
            }
            primary ??= allGpus.OrderByDescending(g => g.DedicatedVideoMemoryMB).First();
            primary.IsPrimary = true;

            // If primary has no resolution (NVIDIA Optimus: iGPU drives display), grab from the one that does
            if (primary.Resolution == "Unknown")
            {
                var displayGpu = allGpus.FirstOrDefault(g => g.Resolution != "Unknown");
                if (displayGpu != null)
                {
                    primary.Resolution = displayGpu.Resolution;
                    primary.RefreshRate = displayGpu.RefreshRate;
                }
            }

            // ── Populate top-level fields ──
            gpu.GpuName = primary.Name;
            gpu.DriverVersion = primary.DriverVersion;
            gpu.DriverDate = primary.DriverDate;
            gpu.DriverOutdated = primary.DriverOutdated;
            gpu.AdapterStatus = primary.AdapterStatus;
            gpu.DedicatedVideoMemoryMB = primary.DedicatedVideoMemoryMB;
            gpu.Resolution = primary.Resolution;
            gpu.RefreshRate = primary.RefreshRate;
            gpu.AllGpus = allGpus;

            gpu.AdditionalAdapters.Clear();
            foreach (var g in allGpus)
            {
                if (g != primary)
                    gpu.AdditionalAdapters.Add(g.Name);
            }

            // ── Temperature: nvidia-smi for NVIDIA (most reliable) ──
            bool isNvidia = primary.Name.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase) ||
                            primary.Name.Contains("GeForce", StringComparison.OrdinalIgnoreCase);
            if (isNvidia)
            {
                try
                {
                    var nvTemp = RunCliTool("nvidia-smi",
                        "--query-gpu=temperature.gpu --format=csv,noheader,nounits");
                    if (double.TryParse(nvTemp?.Trim(), System.Globalization.CultureInfo.InvariantCulture, out double t) && t > 0 && t < 120)
                        gpu.GpuTemperatureC = t;

                    var nvUsage = RunCliTool("nvidia-smi",
                        "--query-gpu=utilization.gpu --format=csv,noheader,nounits");
                    if (nvUsage != null)
                    {
                        string digits = new string(nvUsage.Where(c => char.IsDigit(c) || c == '.').ToArray());
                        if (double.TryParse(digits, System.Globalization.CultureInfo.InvariantCulture, out double u))
                            gpu.GpuUsagePercent = Math.Round(u, 1);
                    }
                }
                catch { }
            }

            // ── Fallback temperature: use cached thermal zones from ScanThermalDiagnostics ──
            if (gpu.GpuTemperatureC <= 0 && _cachedThermalZonesC is { Count: >= 2 })
            {
                // If there are 2+ zones with >5°C difference, second zone is likely GPU
                if (Math.Abs(_cachedThermalZonesC[0] - _cachedThermalZonesC[1]) > 5)
                    gpu.GpuTemperatureC = _cachedThermalZonesC[1];
            }

            gpu.TemperatureFlagged = gpu.GpuTemperatureC > 85;

            // ── Fallback usage: WMI GPU perf counters ──
            if (gpu.GpuUsagePercent <= 0)
            {
                try
                {
                    using var searcher = new ManagementObjectSearcher(
                        "SELECT UtilizationPercentage FROM Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine");
                    searcher.Options.Timeout = TimeSpan.FromSeconds(8);
                    double maxUsage = 0;
                    foreach (var obj in searcher.Get())
                    {
                        double usage = Convert.ToDouble(obj["UtilizationPercentage"] ?? 0);
                        if (usage > maxUsage) maxUsage = usage;
                    }
                    gpu.GpuUsagePercent = Math.Round(maxUsage, 1);
                }
                catch { }
            }

            gpu.UsageFlagged = gpu.GpuUsagePercent > 90;

            // Propagate temp/usage to primary GpuInfo
            primary.TemperatureC = gpu.GpuTemperatureC;
            primary.TemperatureFlagged = gpu.TemperatureFlagged;
            primary.UsagePercent = gpu.GpuUsagePercent;
            primary.UsageFlagged = gpu.UsageFlagged;

            // ── Enumerate physical displays ──
            try
            {
                gpu.Displays = EnumerateDisplays();
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ Display enumeration: {ex.Message}"); }
        }

        /// <summary>
        /// Runs a process and captures text output with a hard timeout.
        /// Reads stdout and stderr asynchronously to prevent buffer deadlock.
        /// Returns false if launch fails or the process does not exit in time.
        /// </summary>
        private static bool TryRunProcessForOutput(
            string fileName, string arguments, int timeoutMs, out string output)
        {
            output = "";
            try
            {
                using var proc = Process.Start(new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
                if (proc == null) return false;

                Task<string> stdoutTask = proc.StandardOutput.ReadToEndAsync();
                Task<string> stderrTask = proc.StandardError.ReadToEndAsync();

                if (!proc.WaitForExit(timeoutMs))
                {
                    try { proc.Kill(entireProcessTree: true); } catch { }
                    return false;
                }

                Task.WaitAll(new Task[] { stdoutTask, stderrTask }, 2000);

                string stdout = stdoutTask.Status == TaskStatus.RanToCompletion ? stdoutTask.Result : "";
                string stderr = stderrTask.Status == TaskStatus.RanToCompletion ? stderrTask.Result : "";
                output = !string.IsNullOrWhiteSpace(stdout) ? stdout.Trim() : stderr.Trim();
                return true;
            }
            catch
            {
                output = "";
                return false;
            }
        }

        /// <summary>
        /// Runs a CLI tool and returns stdout (up to 5 s). Returns null on failure.
        /// </summary>
        private static string RunCliTool(string fileName, string arguments)
        {
            return TryRunProcessForOutput(fileName, arguments, 5000, out string output)
                && !string.IsNullOrWhiteSpace(output)
                ? output
                : null;
        }

        /// <summary>
        /// Reads VRAM for all GPU entries from the display adapter registry keys.
        /// Returns accurate 64-bit values that WMI (UInt32) cannot.
        /// </summary>
        private static List<(string Name, long VramMB)> GetAllGpuVramFromRegistry()
        {
            var results = new List<(string Name, long VramMB)>();
            try
            {
                const string basePath = @"SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}";
                using var baseKey = Registry.LocalMachine.OpenSubKey(basePath);
                if (baseKey == null) return results;

                foreach (string subKeyName in baseKey.GetSubKeyNames())
                {
                    if (!subKeyName.All(char.IsDigit)) continue;
                    using var subKey = baseKey.OpenSubKey(subKeyName);
                    if (subKey == null) continue;

                    string driverDesc = subKey.GetValue("DriverDesc")?.ToString();
                    if (string.IsNullOrEmpty(driverDesc)) continue;

                    long vramMB = 0;

                    // qwMemorySize is a QWORD (64-bit)
                    object qwMem = subKey.GetValue("HardwareInformation.qwMemorySize");
                    if (qwMem is long qwVal && qwVal > 0)
                        vramMB = qwVal / (1024 * 1024);
                    else if (qwMem is byte[] qwBytes && qwBytes.Length >= 8)
                        vramMB = BitConverter.ToInt64(qwBytes, 0) / (1024 * 1024);

                    // Fallback: MemorySize DWORD
                    if (vramMB <= 0)
                    {
                        object dwMem = subKey.GetValue("HardwareInformation.MemorySize");
                        if (dwMem is int dwVal && dwVal > 0)
                            vramMB = (uint)dwVal / (1024 * 1024);
                        else if (dwMem is byte[] dwBytes && dwBytes.Length >= 4)
                            vramMB = BitConverter.ToUInt32(dwBytes, 0) / (1024 * 1024);
                    }

                    if (vramMB > 0)
                        results.Add((driverDesc, vramMB));
                }
            }
            catch { }
            return results;
        }

        // ═══════════════════════════════════════════════════════════════
        //  6. BATTERY HEALTH
        // ═══════════════════════════════════════════════════════════════

        private void ScanBatteryHealth(DiagnosticResult result)
        {
            var batt = result.Battery;

            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Battery");
                searcher.Options.Timeout = TimeSpan.FromSeconds(10);
                var batteries = searcher.Get();
                batt.HasBattery = batteries.Count > 0;

                if (!batt.HasBattery) return;

                foreach (var obj in batteries)
                {
                    batt.DesignCapacityMWh = Convert.ToInt64(obj["DesignCapacity"] ?? 0);
                    batt.FullChargeCapacityMWh = Convert.ToInt64(obj["FullChargeCapacity"] ?? 0);
                    int status = Convert.ToInt32(obj["BatteryStatus"] ?? 0);
                    batt.PowerSource = status == 2 ? "Plugged In (Charging)" : "On Battery";
                }

                if (batt.DesignCapacityMWh > 0)
                {
                    batt.HealthPercent = Math.Round(
                        (double)batt.FullChargeCapacityMWh / batt.DesignCapacityMWh * 100, 1);
                    batt.HealthFlagged = batt.HealthPercent < 50;
                }
            }
            catch { batt.HasBattery = false; }

            // Power plan
            batt.PowerPlan = GetPowerPlan();
        }

        // ═══════════════════════════════════════════════════════════════
        //  7. STARTUP PROGRAMS
        // ═══════════════════════════════════════════════════════════════

        private void ScanStartupPrograms(DiagnosticResult result)
        {
            var startup = result.Startup;

            // HKCU Run
            AddStartupEntries(startup, Registry.CurrentUser,
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);

            // HKLM Run
            AddStartupEntries(startup, Registry.LocalMachine,
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);

            // Override enabled/disabled status from StartupApproved registry
            ApplyStartupApprovedStatus(startup, Registry.CurrentUser,
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run");
            ApplyStartupApprovedStatus(startup, Registry.LocalMachine,
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run");

            startup.EnabledCount = startup.Entries.Count(e => e.Enabled);
            startup.TooManyFlagged = startup.EnabledCount > 10;
        }

        private static void AddStartupEntries(StartupDiagnostics startup, RegistryKey root,
            string subKey, bool enabled)
        {
            try
            {
                using var key = root.OpenSubKey(subKey);
                if (key == null) return;
                foreach (var name in key.GetValueNames())
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (startup.Entries.Any(e => e.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                        continue;
                    startup.Entries.Add(new StartupEntry
                    {
                        Name = name,
                        Publisher = "",
                        Enabled = enabled
                    });
                }
            }
            catch { }
        }

        private static void ApplyStartupApprovedStatus(StartupDiagnostics startup, RegistryKey root,
            string subKey)
        {
            try
            {
                using var key = root.OpenSubKey(subKey);
                if (key == null) return;
                foreach (var name in key.GetValueNames())
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var existing = startup.Entries.FirstOrDefault(
                        e => e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                    if (existing == null)
                    {
                        existing = new StartupEntry { Name = name, Publisher = "" };
                        startup.Entries.Add(existing);
                    }
                    // First byte: 02/06 = enabled, 03/07 = disabled
                    if (key.GetValue(name) is byte[] data && data.Length >= 1)
                        existing.Enabled = (data[0] & 0x01) == 0;
                }
            }
            catch { }
        }

        // ═══════════════════════════════════════════════════════════════
        //  8. VISUAL & DISPLAY SETTINGS
        // ═══════════════════════════════════════════════════════════════

        private void ScanVisualSettings(DiagnosticResult result)
        {
            var vs = result.VisualSettings;

            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects");
                int setting = (int)(key?.GetValue("VisualFXSetting") ?? 0);
                vs.VisualEffectsSetting = setting switch
                {
                    0 => "Let Windows choose (default)",
                    1 => "Best Appearance",
                    2 => "Best Performance",
                    3 => "Custom",
                    _ => "Unknown"
                };
            }
            catch { }

            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                vs.TransparencyEnabled = (int)(key?.GetValue("EnableTransparency") ?? 1) == 1;
            }
            catch { }

            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop\WindowMetrics");
                string minAnimate = key?.GetValue("MinAnimate")?.ToString() ?? "1";
                vs.AnimationsEnabled = minAnimate != "0";
            }
            catch { }
        }

        // ═══════════════════════════════════════════════════════════════
        //  9. ANTIVIRUS & SECURITY
        // ═══════════════════════════════════════════════════════════════

        private void ScanAntivirus(DiagnosticResult result)
        {
            var av = result.Antivirus;

            try
            {
                using var searcher = new ManagementObjectSearcher(
                    @"root\SecurityCenter2", "SELECT displayName FROM AntiVirusProduct");
                searcher.Options.Timeout = TimeSpan.FromSeconds(5);
                foreach (var obj in searcher.Get())
                {
                    av.AntivirusName = obj["displayName"]?.ToString() ?? "Unknown";
                    break;
                }
            }
            catch (Exception ex) { av.AntivirusName = "Unable to detect"; OnLog?.Invoke($"  ⚡ Antivirus: {ex.Message}"); }

            // Check if a full Defender scan is running via WMI AMNamespace
            // This is far more reliable than process memory heuristics
            try
            {
                // Method 1: Check if MsMpEng process has high CPU (sampled)
                var avProcesses = Process.GetProcessesByName("MsMpEng");
                bool defenderRunning = avProcesses.Length > 0;
                foreach (var p in avProcesses) { try { p.Dispose(); } catch { } }

                if (defenderRunning)
                {
                    // Method 2: Check Defender scan status via registry
                    // When a scan is active, ScanRunning is set to 1
                    try
                    {
                        using var key = Registry.LocalMachine.OpenSubKey(
                            @"SOFTWARE\Microsoft\Windows Defender\Scan");
                        if (key != null)
                        {
                            int scanRunning = Convert.ToInt32(key.GetValue("ScanRunning") ?? 0);
                            av.FullScanRunning = scanRunning == 1;
                        }
                    }
                    catch
                    {
                        // Fallback: check MpCmdRun process which runs during active scans
                        var cmdRun = Process.GetProcessesByName("MpCmdRun");
                        av.FullScanRunning = cmdRun.Length > 0;
                        foreach (var p in cmdRun) { try { p.Dispose(); } catch { } }
                    }
                }
            }
            catch { }

            // BitLocker status
            try
            {
                if (TryRunProcessForOutput("manage-bde", "-status C:", 10_000, out string output))
                {
                    if (output.Contains("Fully Encrypted", StringComparison.OrdinalIgnoreCase))
                        av.BitLockerStatus = "Fully Encrypted";
                    else if (output.Contains("Fully Decrypted", StringComparison.OrdinalIgnoreCase))
                        av.BitLockerStatus = "Not Encrypted";
                    else if (output.Contains("Encryption in Progress", StringComparison.OrdinalIgnoreCase))
                        av.BitLockerStatus = "Encryption In Progress";
                    else
                        av.BitLockerStatus = "Unknown";
                }
                else
                {
                    av.BitLockerStatus = "Check timed out";
                }
            }
            catch { av.BitLockerStatus = "Unable to check"; }
        }

        // ═══════════════════════════════════════════════════════════════
        //  10. WINDOWS UPDATE
        // ═══════════════════════════════════════════════════════════════

        private void ScanWindowsUpdate(DiagnosticResult result)
        {
            var wu = result.WindowsUpdate;

            // Check pending updates via registry
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired");
                wu.PendingUpdates = key != null ? "Pending reboot required" : "No pending updates detected";
            }
            catch { wu.PendingUpdates = "Unable to check"; }

            // Update service status — atomic snapshot of both status and start type
            (wu.UpdateServiceStatus, wu.ServiceStartType) = QueryServiceState("wuauserv");

            // Update cache size
            wu.UpdateCacheMB = GetFolderSizeMB(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "SoftwareDistribution", "Download"));
        }

        // ═══════════════════════════════════════════════════════════════
        //  11. NETWORK DIAGNOSTICS
        // ═══════════════════════════════════════════════════════════════

        private void ScanNetwork(DiagnosticResult result)
        {
            var ct = _phaseToken;
            var net = result.Network;

            // Connection type and adapter info
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    "SELECT NetConnectionID, Speed, NetConnectionStatus FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2");
                searcher.Options.Timeout = TimeSpan.FromSeconds(10);
                foreach (var obj in searcher.Get())
                {
                    string name = obj["NetConnectionID"]?.ToString() ?? "";
                    long speed = Convert.ToInt64(obj["Speed"] ?? 0);
                    net.AdapterSpeed = speed > 0 ? $"{speed / 1_000_000} Mbps" : "Unknown";

                    if (name.Contains("Wi-Fi", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("Wireless", StringComparison.OrdinalIgnoreCase))
                    {
                        net.ConnectionType = "WiFi";
                    }
                    else if (name.Contains("Ethernet", StringComparison.OrdinalIgnoreCase))
                    {
                        net.ConnectionType = "Ethernet";
                    }
                    else
                    {
                        net.ConnectionType = name;
                    }
                    break;
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ Network adapter: {ex.Message}"); }

            // WiFi details via netsh
            if (net.ConnectionType == "WiFi")
            {
                try
                {
                    if (TryRunProcessForOutput("netsh", "wlan show interfaces", 5000, out string output))
                    {
                        foreach (var line in output.Split('\n'))
                        {
                            var trimmed = line.Trim();
                            if (trimmed.StartsWith("Signal", StringComparison.OrdinalIgnoreCase))
                                net.WifiSignalStrength = trimmed.Split(':').LastOrDefault()?.Trim() ?? "Unknown";
                            if (trimmed.StartsWith("Receive rate", StringComparison.OrdinalIgnoreCase) ||
                                trimmed.StartsWith("Transmit rate", StringComparison.OrdinalIgnoreCase))
                                net.LinkSpeed = trimmed.Split(':').LastOrDefault()?.Trim() ?? "Unknown";
                            if (trimmed.StartsWith("Radio type", StringComparison.OrdinalIgnoreCase) ||
                                trimmed.StartsWith("Band", StringComparison.OrdinalIgnoreCase))
                            {
                                string val = trimmed.Split(':').LastOrDefault()?.Trim() ?? "";
                                if (val.Contains("5") || val.Contains("802.11a") || val.Contains("802.11ac") || val.Contains("802.11ax"))
                                    net.WifiBand = "5 GHz";
                                else if (val.Contains("2.4") || val.Contains("802.11b") || val.Contains("802.11g") || val.Contains("802.11n"))
                                    net.WifiBand = "2.4 GHz";
                                else
                                    net.WifiBand = val;
                            }
                        }
                    }
                }
                catch { }
            }

            if (ct.IsCancellationRequested) return;

            // DNS response time — 3-sample median for accuracy
            try
            {
                // Warm up DNS resolver + OS cache with unrelated domain
                try { System.Net.Dns.GetHostEntry("dns.google"); } catch { }

                // Use unique subdomains to bypass OS/resolver caching
                string[] targets =
                [
                    $"{Guid.NewGuid():N}.example.com",
                    $"{Guid.NewGuid():N}.example.com",
                    $"{Guid.NewGuid():N}.example.com"
                ];

                var samples = new List<long>(3);
                foreach (string target in targets)
                {
                    var sw = Stopwatch.StartNew();
                    try { System.Net.Dns.GetHostEntry(target); }
                    catch (System.Net.Sockets.SocketException) { /* NXDOMAIN is expected — we only measure resolver round-trip */ }
                    sw.Stop();
                    samples.Add(sw.ElapsedMilliseconds);
                }

                samples.Sort();
                long median = samples[1]; // middle of 3
                net.DnsResponseTime = $"{median} ms";
            }
            catch { net.DnsResponseTime = "Failed"; }

            if (ct.IsCancellationRequested) return;

            // VPN check
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    "SELECT Name, NetConnectionStatus FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2");
                searcher.Options.Timeout = TimeSpan.FromSeconds(10);
                foreach (var obj in searcher.Get())
                {
                    string name = obj["Name"]?.ToString() ?? "";
                    if (name.Contains("VPN", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("TAP", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("Cisco", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("GlobalProtect", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("Fortinet", StringComparison.OrdinalIgnoreCase))
                    {
                        net.VpnActive = true;
                        net.VpnClient = name;
                        break;
                    }
                }
            }
            catch { }

            if (ct.IsCancellationRequested) return;

            // Ping latency
            try
            {
                if (TryRunProcessForOutput("ping", "-n 3 8.8.8.8", 15_000, out string output))
                {
                    // Parse "Average = Xms"
                    var avgLine = output.Split('\n')
                        .FirstOrDefault(l => l.Contains("Average", StringComparison.OrdinalIgnoreCase) ||
                                             l.Contains("Moyenne", StringComparison.OrdinalIgnoreCase));
                    if (avgLine != null)
                    {
                        var parts = avgLine.Split('=');
                        net.PingLatency = parts.LastOrDefault()?.Trim() ?? "Unknown";
                    }
                    else if (output.Contains("Request timed out", StringComparison.OrdinalIgnoreCase))
                    {
                        net.PingLatency = "Timed out";
                    }
                }
                else
                {
                    net.PingLatency = "Failed";
                }
            }
            catch { net.PingLatency = "Failed"; }
        }

        // ═══════════════════════════════════════════════════════════════
        //  11b. INTERNET SPEED TEST
        // ═══════════════════════════════════════════════════════════════

        private async Task ScanSpeedTest(DiagnosticResult result)
        {
            var net = result.Network;

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            http.DefaultRequestHeaders.Add("User-Agent", "EDR-Diag/2.0");

            // Warm up: establish connection (DNS + TLS) so it doesn't skew measurements
            try
            {
                using var warmup = await http.GetAsync("https://speed.cloudflare.com/__down?bytes=1024",
                    HttpCompletionOption.ResponseHeadersRead, _cts.Token);
                warmup.EnsureSuccessStatusCode();
                // Drain the small response
                using var ws = await warmup.Content.ReadAsStreamAsync();
                var wBuf = new byte[2048];
                while (await ws.ReadAsync(wBuf.AsMemory(), _cts.Token) > 0) { }
            }
            catch (OperationCanceledException) { throw; }
            catch { /* warm-up failure is not critical */ }

            // ── Download speed ──
            // Cloudflare speed test: download 25 MB payload
            try
            {
                const int downloadBytes = 25 * 1024 * 1024; // 25 MB for more accurate measurement
                string url = $"https://speed.cloudflare.com/__down?bytes={downloadBytes}";

                // Use ResponseHeadersRead — start timing only after headers arrive
                var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, _cts.Token);
                response.EnsureSuccessStatusCode();

                long totalRead = 0;
                var buffer = new byte[81920];
                using var stream = await response.Content.ReadAsStreamAsync();

                // Start timing AFTER connection is established — measures pure throughput
                var sw = Stopwatch.StartNew();
                int read;
                while ((read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), _cts.Token)) > 0)
                {
                    totalRead += read;
                }
                sw.Stop();

                if (totalRead > 0 && sw.ElapsedMilliseconds > 0)
                {
                    double seconds = sw.ElapsedMilliseconds / 1000.0;
                    double bitsPerSecond = (totalRead * 8.0) / seconds;
                    net.DownloadSpeedMbps = Math.Round(bitsPerSecond / 1_000_000.0, 2);
                    OnLog?.Invoke($"  Download speed: {net.DownloadSpeedMbps} Mbps ({totalRead / 1024.0 / 1024.0:F1} MB in {seconds:F1}s)");
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                net.SpeedTestError = $"Download: {ex.Message}";
                OnLog?.Invoke($"  Speed test download failed: {ex.Message}");
            }

            // ── Upload speed ──
            // Cloudflare speed test: upload 5 MB payload
            try
            {
                const int uploadBytes = 5 * 1024 * 1024; // 5 MB
                string url = "https://speed.cloudflare.com/__up";
                var payload = new byte[uploadBytes]; // zeros are fine

                using var content = new ByteArrayContent(payload);
                // Start timing AFTER request is sent — measure server response round-trip
                var sw = Stopwatch.StartNew();
                var response = await http.PostAsync(url, content, _cts.Token);
                response.EnsureSuccessStatusCode();
                // Read full response to ensure upload is complete
                await response.Content.ReadAsStringAsync();
                sw.Stop();

                if (sw.ElapsedMilliseconds > 0)
                {
                    double seconds = sw.ElapsedMilliseconds / 1000.0;
                    double bitsPerSecond = (uploadBytes * 8.0) / seconds;
                    net.UploadSpeedMbps = Math.Round(bitsPerSecond / 1_000_000.0, 2);
                    OnLog?.Invoke($"  Upload speed: {net.UploadSpeedMbps} Mbps ({uploadBytes / 1024.0 / 1024.0:F1} MB in {seconds:F1}s)");
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(net.SpeedTestError))
                    net.SpeedTestError = $"Upload: {ex.Message}";
                else
                    net.SpeedTestError += $"; Upload: {ex.Message}";
                OnLog?.Invoke($"  Speed test upload failed: {ex.Message}");
            }

            // Flag if download speed is under 10 Mbps
            net.SpeedTestFlagged = net.DownloadSpeedMbps > 0 && net.DownloadSpeedMbps < 10;
        }

        // ═══════════════════════════════════════════════════════════════
        //  11b. NETWORK DRIVE HEALTH
        // ═══════════════════════════════════════════════════════════════

        private void ScanNetworkDrives(DiagnosticResult result)
        {
            var nd = result.NetworkDrives;

            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    if (drive.DriveType != DriveType.Network)
                        continue;

                    var info = new NetworkDriveInfo
                    {
                        DriveLetter = drive.Name.TrimEnd('\\'),
                        DriveLabel = ""
                    };

                    // Get UNC path via WMI
                    try
                    {
                        using var searcher = new ManagementObjectSearcher(
                            $"SELECT ProviderName FROM Win32_LogicalDisk WHERE DeviceID='{info.DriveLetter}'");
                        foreach (ManagementObject disk in searcher.Get())
                        {
                            info.UncPath = disk["ProviderName"]?.ToString() ?? "";
                            break;
                        }
                    }
                    catch { }

                    // Measure accessibility and latency
                    var sw = Stopwatch.StartNew();
                    try
                    {
                        if (drive.IsReady)
                        {
                            info.IsAccessible = true;
                            info.DriveLabel = drive.VolumeLabel;
                            info.TotalMB = drive.TotalSize / (1024 * 1024);
                            info.FreeMB = drive.AvailableFreeSpace / (1024 * 1024);
                            info.PercentUsed = info.TotalMB > 0
                                ? (int)((info.TotalMB - info.FreeMB) * 100 / info.TotalMB)
                                : 0;
                            info.SpaceFlagged = info.PercentUsed > 90;

                            // Measure actual file access latency
                            try
                            {
                                Directory.GetDirectories(drive.RootDirectory.FullName);
                            }
                            catch { }
                        }
                    }
                    catch
                    {
                        info.IsAccessible = false;
                    }
                    sw.Stop();
                    info.LatencyMs = sw.ElapsedMilliseconds;
                    info.LatencyFlagged = info.LatencyMs > 200;

                    // Check Offline Files / CSC sync status
                    try
                    {
                        string cscPath = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                            "CSC");
                        info.OfflineFilesEnabled = Directory.Exists(cscPath);

                        if (info.OfflineFilesEnabled)
                        {
                            using var searcher = new ManagementObjectSearcher(
                                "SELECT Status FROM Win32_OfflineFilesCache");
                            foreach (ManagementObject obj in searcher.Get())
                            {
                                info.SyncStatus = obj["Status"]?.ToString() ?? "Unknown";
                                break;
                            }
                        }
                    }
                    catch
                    {
                        info.SyncStatus = info.OfflineFilesEnabled ? "Unknown" : "N/A";
                    }

                    nd.Drives.Add(info);
                    OnLog?.Invoke($"  Network drive {info.DriveLetter}: " +
                        (info.IsAccessible
                            ? $"OK ({info.LatencyMs}ms, {info.PercentUsed}% used)"
                            : "Unreachable"));
                }
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"  Network drive scan error: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  12. OUTLOOK / EMAIL
        // ═══════════════════════════════════════════════════════════════

        private void ScanOutlook(DiagnosticResult result)
        {
            var outlook = result.Outlook;

            // Check if Outlook is installed
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
                string products = key?.GetValue("ProductReleaseIds")?.ToString() ?? "";
                outlook.OutlookInstalled = products.Contains("Outlook", StringComparison.OrdinalIgnoreCase);

                if (!outlook.OutlookInstalled)
                {
                    using var key2 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office");
                    outlook.OutlookInstalled = key2 != null;
                }
            }
            catch { }

            if (!outlook.OutlookInstalled) return;

            // Find OST/PST files
            string outlookDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft", "Outlook");

            if (Directory.Exists(outlookDir))
            {
                try
                {
                    foreach (var file in Directory.EnumerateFiles(outlookDir, "*.ost", SearchOption.AllDirectories)
                        .Concat(Directory.EnumerateFiles(outlookDir, "*.pst", SearchOption.AllDirectories)))
                    {
                        try
                        {
                            // FileInfo.Length reads NTFS metadata, not the file stream —
                            // succeeds even when Outlook holds an exclusive lock on the .ost
                            var fi = new FileInfo(file);
                            long sizeMB = fi.Length / 1024 / 1024;
                            outlook.DataFiles.Add(new OutlookDataFile
                            {
                                Path = file,
                                SizeMB = sizeMB,
                                SizeFlagged = sizeMB > 5120
                            });
                        }
                        catch (IOException ex)
                        {
                            OnLog?.Invoke($"  ⚡ Outlook file locked: {Path.GetFileName(file)} — {ex.Message}");
                        }
                        catch (UnauthorizedAccessException ex)
                        {
                            OnLog?.Invoke($"  ⚡ Outlook file access denied: {Path.GetFileName(file)} — {ex.Message}");
                        }
                    }
                }
                catch { }
            }

            // Outlook add-ins
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Office\Outlook\Addins");
                outlook.AddInCount = key?.GetSubKeyNames().Length ?? 0;
            }
            catch { }
        }

        // ═══════════════════════════════════════════════════════════════
        //  13. BROWSER CHECK
        // ═══════════════════════════════════════════════════════════════

        private void ScanBrowsers(DiagnosticResult result)
        {
            var browsers = result.Browser;

            // Chrome
            ScanBrowser(browsers, "Chrome", "chrome",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Google", "Chrome", "User Data"),
                "Default\\Extensions");

            // Edge
            ScanBrowser(browsers, "Microsoft Edge", "msedge",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Edge", "User Data"),
                "Default\\Extensions");

            // Firefox
            ScanFirefox(browsers);
        }

        private void ScanBrowser(BrowserDiagnostics browsers, string name, string processName,
            string userDataPath, string extensionSubPath)
        {
            var info = new BrowserInfo { Name = name };

            try
            {
                var walker = TreeWalker.ControlViewWalker;
                var procs = Process.GetProcessesByName(processName);
                foreach (var p in procs)
                {
                    IntPtr hwnd = IntPtr.Zero;
                    try { hwnd = p.MainWindowHandle; } catch { }
                    p.Dispose();
                    if (hwnd == IntPtr.Zero) continue;

                    try
                    {
                        var window = AutomationElement.FromHandle(hwnd);
                        int count = CountTabItemsShallow(walker, window, 8);
                        if (count > info.OpenTabs)
                            info.OpenTabs = count;
                        if (count > 1) break; // found the real browser window
                    }
                    catch { }
                }
            }
            catch { }

            // Cache size (include Code Cache to match optimizer cleanup)
            if (Directory.Exists(userDataPath))
            {
                string cachePath = Path.Combine(userDataPath, "Default", "Cache");
                string codeCachePath = Path.Combine(userDataPath, "Default", "Code Cache");
                long cacheSize = 0;
                if (Directory.Exists(cachePath))
                    cacheSize += GetFolderSizeMB(cachePath);
                if (Directory.Exists(codeCachePath))
                    cacheSize += GetFolderSizeMB(codeCachePath);
                info.CacheSizeMB = cacheSize;

                // Extension count
                string extPath = Path.Combine(userDataPath, extensionSubPath);
                if (Directory.Exists(extPath))
                {
                    try
                    {
                        info.ExtensionCount = Directory.GetDirectories(extPath).Length;
                    }
                    catch { }
                }
            }

            if (info.OpenTabs > 0 || Directory.Exists(userDataPath))
                browsers.Browsers.Add(info);
        }

        /// <summary>
        /// Counts ControlType.TabItem elements by walking the UI Automation
        /// control-view tree to a limited depth.  When a TabItem is found it
        /// is counted but NOT recursed into — this prevents web-page ARIA
        /// role="tab" elements (which live deep inside TabItem content) from
        /// inflating the count.  Does not depend on the tab strip container
        /// being any specific control type.
        /// </summary>
        private static int CountTabItemsShallow(
            TreeWalker walker, AutomationElement element, int maxDepth)
        {
            if (maxDepth <= 0) return 0;

            int count = 0;
            AutomationElement child;
            try { child = walker.GetFirstChild(element); }
            catch { return 0; }

            while (child != null)
            {
                try
                {
                    if (child.Current.ControlType == ControlType.TabItem)
                        count++;                      // count it, don't recurse into it
                    else
                        count += CountTabItemsShallow(walker, child, maxDepth - 1);
                }
                catch { }

                try { child = walker.GetNextSibling(child); }
                catch { child = null; }
            }

            return count;
        }

        private void ScanFirefox(BrowserDiagnostics browsers)
        {
            string profilesDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Mozilla", "Firefox", "Profiles");
            if (!Directory.Exists(profilesDir)) return;

            var info = new BrowserInfo { Name = "Firefox" };

            // Count open tabs via recovery files
            try
            {
                var procs = Process.GetProcessesByName("firefox");
                int processCount = procs.Length;
                foreach (var p in procs) p.Dispose();
                // Firefox uses multi-process (Fission) — process count is not a reliable tab indicator
                if (processCount > 1)
                    info.OpenTabs = Math.Clamp(processCount / 2, 1, 10); // rough lower-bound estimate
            }
            catch { }

            // Cache + extensions from first profile
            try
            {
                var profileDir = Directory.GetDirectories(profilesDir)
                    .FirstOrDefault(d => d.EndsWith(".default-release") || d.Contains("default"));
                if (profileDir != null)
                {
                    string cachePath = Path.Combine(profileDir, "cache2");
                    if (Directory.Exists(cachePath))
                        info.CacheSizeMB = GetFolderSizeMB(cachePath);

                    string extPath = Path.Combine(profileDir, "extensions");
                    if (Directory.Exists(extPath))
                        info.ExtensionCount = Directory.GetFileSystemEntries(extPath).Length;
                }
            }
            catch { }

            if (info.OpenTabs > 0 || Directory.Exists(profilesDir))
                browsers.Browsers.Add(info);
        }

        // ═══════════════════════════════════════════════════════════════
        //  14. USER PROFILE
        // ═══════════════════════════════════════════════════════════════

        private void ScanUserProfile(DiagnosticResult result)
        {
            var profile = result.UserProfile;
            string userDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            // Profile size (top-level folders only for speed)
            profile.ProfileSizeMB = GetFolderSizeMB(userDir, maxDepth: 2);

            // Profile age — use NTUSER.DAT creation time (more reliable than folder creation time
            // which can get reset during OS upgrades)
            try
            {
                string ntUserDat = Path.Combine(userDir, "NTUSER.DAT");
                DateTime created;
                if (File.Exists(ntUserDat))
                    created = File.GetCreationTime(ntUserDat);
                else
                    created = Directory.GetCreationTime(userDir);

                var age = DateTime.Now - created;
                if (age.TotalDays >= 365)
                    profile.ProfileAge = $"{(int)(age.TotalDays / 365)} year(s)";
                else if (age.TotalDays >= 30)
                    profile.ProfileAge = $"{(int)(age.TotalDays / 30)} month(s)";
                else if (age.TotalDays >= 1)
                    profile.ProfileAge = $"{(int)age.TotalDays} day(s)";
                else
                    profile.ProfileAge = "Less than a day";
            }
            catch { }

            // Desktop item count
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                profile.DesktopItemCount = Directory.GetFileSystemEntries(desktop).Length;
                profile.DesktopItemsFlagged = profile.DesktopItemCount > 50;
            }
            catch { }

            // Corruption check — look for temp profile indicators
            try
            {
                string profilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                profile.CorruptionDetected = profilePath.Contains(".bak", StringComparison.OrdinalIgnoreCase) ||
                                              profilePath.Contains("TEMP", StringComparison.OrdinalIgnoreCase);
            }
            catch { }
        }

        // ═══════════════════════════════════════════════════════════════
        //  15. GENERAL OFFICE
        // ═══════════════════════════════════════════════════════════════

        private void ScanOffice(DiagnosticResult result)
        {
            var office = result.Office;

            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
                if (key != null)
                {
                    office.OfficeInstalled = true;
                    office.OfficeVersion = key.GetValue("VersionToReport")?.ToString() ?? "Unknown";
                }
                else
                {
                    // Check MSI-based installations
                    string[] versions = { "16.0", "15.0", "14.0" };
                    foreach (var v in versions)
                    {
                        using var vKey = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Microsoft\Office\{v}\Common\InstallRoot");
                        if (vKey != null)
                        {
                            office.OfficeInstalled = true;
                            office.OfficeVersion = $"Office {v}";
                            break;
                        }
                    }
                }
            }
            catch { }

            // Check for repair needed — only flag if Click-to-Run is present but misconfigured
            if (office.OfficeInstalled)
            {
                try
                {
                    using var key = Registry.LocalMachine.OpenSubKey(
                        @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
                    if (key != null)
                    {
                        // Only flag if C2R exists but has no update channel (genuinely broken)
                        string updateChannel = key.GetValue("UpdateChannel")?.ToString() ?? "";
                        string versionToReport = key.GetValue("VersionToReport")?.ToString() ?? "";
                        office.RepairNeeded = string.IsNullOrEmpty(updateChannel) && !string.IsNullOrEmpty(versionToReport);
                    }
                }
                catch { }
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  16. INSTALLED SOFTWARE
        // ═══════════════════════════════════════════════════════════════

        private void ScanInstalledSoftware(DiagnosticResult result)
        {
            var ct = _phaseToken;
            var sw = result.InstalledSoftware;
            var apps = new Dictionary<string, InstalledApp>(StringComparer.OrdinalIgnoreCase);

            string[] regPaths =
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (var regPath in regPaths)
            {
                try
                {
                    using var key = Registry.LocalMachine.OpenSubKey(regPath);
                    if (key == null) continue;

                    foreach (var subKeyName in key.GetSubKeyNames())
                    {
                        if (ct.IsCancellationRequested) return;
                        try
                        {
                            using var subKey = key.OpenSubKey(subKeyName);
                            string name = subKey?.GetValue("DisplayName")?.ToString();
                            if (string.IsNullOrWhiteSpace(name)) continue;
                            if (apps.ContainsKey(name)) continue;

                            apps[name] = new InstalledApp
                            {
                                Name = name,
                                Version = subKey.GetValue("DisplayVersion")?.ToString() ?? "",
                                Publisher = subKey.GetValue("Publisher")?.ToString() ?? ""
                            };
                        }
                        catch { }
                    }
                }
                catch (Exception ex) { OnLog?.Invoke($"  ⚡ Software registry: {ex.Message}"); }
            }

            sw.Applications = apps.Values.OrderBy(a => a.Name).ToList();
            sw.TotalCount = sw.Applications.Count;

            // Detect end-of-life / unsupported software
            string[] eolPatterns = {
                "Adobe Flash Player", "Silverlight", "Java 6", "Java 7",
                "Python 2", "Internet Explorer", "Adobe Shockwave",
                "QuickTime", "RealPlayer", "Windows Media Player"
            };
            sw.EOLApps = sw.Applications.Where(a =>
                eolPatterns.Any(p => a.Name.Contains(p, StringComparison.OrdinalIgnoreCase))).ToList();

            // Detect known bloatware / PUPs
            string[] bloatPatterns = {
                "Toolbar", "Ask ", "Conduit", "Babylon", "MyWebSearch",
                "WildTangent", "Candy Crush", "Bubble Witch",
                "McAfee WebAdvisor", "Norton Power Eraser", "Bonjour",
                "CyberLink", "WinZip Driver", "Driver Booster",
                "IObit", "Avast Cleanup", "CCleaner Browser"
            };
            sw.BloatwareApps = sw.Applications.Where(a =>
                bloatPatterns.Any(p => a.Name.Contains(p, StringComparison.OrdinalIgnoreCase))).ToList();

            // Collect Visual C++ Redistributable versions
            sw.RuntimeApps = sw.Applications.Where(a =>
                a.Name.Contains("Visual C++", StringComparison.OrdinalIgnoreCase) &&
                a.Name.Contains("Redistributable", StringComparison.OrdinalIgnoreCase)).ToList();
            sw.RuntimeCount = sw.RuntimeApps.Count;
        }

        // ═══════════════════════════════════════════════════════════════
        //  17. EVENT LOG ANALYSIS
        // ═══════════════════════════════════════════════════════════════

        private void ScanEventLogs(DiagnosticResult result)
        {
            var ct = _phaseToken;
            var el = result.EventLog;
            var baseCutoffUtc = DateTime.UtcNow.AddDays(-30);

            // Per-category cutoffs: only flag events that occurred AFTER the
            // last successful remediation for that category.
            DateTime CutoffFor(string category)
            {
                DateTime? remediatedAt = Optimizer.GetRemediationTimestamp(category);
                return remediatedAt.HasValue && remediatedAt.Value > baseCutoffUtc
                    ? remediatedAt.Value
                    : baseCutoffUtc;
            }

            DateTime bsodCutoff = CutoffFor("BSODs");
            DateTime shutdownCutoff = CutoffFor("UnexpectedShutdowns");
            DateTime diskCutoff = CutoffFor("DiskErrors");
            DateTime appCrashCutoff = CutoffFor("AppCrashes");

            // BSODs — BugCheck events from Windows Error Reporting
            try
            {
                using var logSystem = new System.Diagnostics.Eventing.Reader.EventLogReader(
                    new System.Diagnostics.Eventing.Reader.EventLogQuery(
                        "System", System.Diagnostics.Eventing.Reader.PathType.LogName,
                        "*[System[Provider[@Name='Microsoft-Windows-WER-SystemErrorReporting'] and (EventID=1001) and TimeCreated[timediff(@SystemTime) <= 2592000000]]]"));
                ReadEntries(logSystem, el.BSODs, 10, bsodCutoff);
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ BSOD log: {ex.Message}"); }

            if (ct.IsCancellationRequested) return;

            // Unexpected shutdowns — EventLog source, ID 6008
            try
            {
                using var logShutdown = new System.Diagnostics.Eventing.Reader.EventLogReader(
                    new System.Diagnostics.Eventing.Reader.EventLogQuery(
                        "System", System.Diagnostics.Eventing.Reader.PathType.LogName,
                        "*[System[Provider[@Name='EventLog'] and (EventID=6008) and TimeCreated[timediff(@SystemTime) <= 2592000000]]]"));
                ReadEntries(logShutdown, el.UnexpectedShutdowns, 10, shutdownCutoff);
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ Shutdown log: {ex.Message}"); }

            if (ct.IsCancellationRequested) return;

            // Disk errors — disk / Ntfs / volmgr
            try
            {
                using var logDisk = new System.Diagnostics.Eventing.Reader.EventLogReader(
                    new System.Diagnostics.Eventing.Reader.EventLogQuery(
                        "System", System.Diagnostics.Eventing.Reader.PathType.LogName,
                        "*[System[(Provider[@Name='disk'] or Provider[@Name='Ntfs'] or Provider[@Name='volmgr']) and (Level=1 or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= 2592000000]]]"));
                ReadEntries(logDisk, el.DiskErrors, 10, diskCutoff);
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ Disk error log: {ex.Message}"); }

            if (ct.IsCancellationRequested) return;

            // App crashes — Application Error (1000) and Application Hang (1002)
            try
            {
                using var logApp = new System.Diagnostics.Eventing.Reader.EventLogReader(
                    new System.Diagnostics.Eventing.Reader.EventLogQuery(
                        "Application", System.Diagnostics.Eventing.Reader.PathType.LogName,
                        "*[System[Provider[@Name='Application Error' or @Name='Application Hang'] and (EventID=1000 or EventID=1002) and TimeCreated[timediff(@SystemTime) <= 2592000000]]]"));
                ReadEntries(logApp, el.AppCrashes, 10, appCrashCutoff);
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚡ App crash log: {ex.Message}"); }

            OnLog?.Invoke($"  Found {el.TotalEventCount} event(s): {el.BSODs.Count} BSODs, {el.AppCrashes.Count} app crashes, {el.DiskErrors.Count} disk errors, {el.UnexpectedShutdowns.Count} unexpected shutdowns");

            // Log remediation context for categories with no new events
            DateTime? anyRemediation = Optimizer.GetRemediationTimestamp();
            if (anyRemediation.HasValue && el.TotalEventCount == 0)
                OnLog?.Invoke($"  (older events were remediated on {anyRemediation.Value.ToLocalTime():g})");
        }

        private static void ReadEntries(
            System.Diagnostics.Eventing.Reader.EventLogReader reader,
            List<EventLogEntry> target, int maxEntries, DateTime cutoffUtc)
        {
            System.Diagnostics.Eventing.Reader.EventRecord record;
            while ((record = reader.ReadEvent()) != null && target.Count < maxEntries)
            {
                using (record)
                {
                    // Skip events that occurred before the cutoff (already remediated).
                    // Normalize to UTC to avoid timezone mismatches — EventRecord.TimeCreated
                    // may return UTC while the cutoff may originate from local time.
                    if (record.TimeCreated.HasValue &&
                        record.TimeCreated.Value.ToUniversalTime() < cutoffUtc)
                        continue;

                    string message;
                    try { message = record.FormatDescription() ?? ""; }
                    catch { message = ""; }

                    // Truncate long messages
                    if (message.Length > 200)
                        message = message[..200] + "...";

                    target.Add(new EventLogEntry
                    {
                        Source = record.ProviderName ?? "",
                        EventId = record.Id,
                        Level = record.LevelDisplayName ?? (record.Level?.ToString() ?? ""),
                        Timestamp = record.TimeCreated ?? DateTime.MinValue,
                        Message = message
                    });
                }
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  FLAGGED ISSUES BUILDER
        // ═══════════════════════════════════════════════════════════════

        private static void BuildFlaggedIssues(DiagnosticResult r)
        {
            var issues = r.FlaggedIssues;

            // System Overview
            if (r.SystemOverview.UptimeFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = r.SystemOverview.Uptime.TotalDays > 14 ? Severity.Critical : Severity.Warning,
                    Category = "System Overview",
                    Description = $"System uptime is {(int)r.SystemOverview.Uptime.TotalDays} days",
                    Recommendation = "Restart the system to clear memory leaks and apply pending updates"
                });

            // CPU
            if (r.Cpu.CpuLoadFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = r.Cpu.CpuLoadPercent > 50 ? Severity.Critical : Severity.Warning,
                    Category = "CPU",
                    Description = $"CPU load is {r.Cpu.CpuLoadPercent}% at idle (above 30% threshold)",
                    Recommendation = "Identify and close resource-heavy background processes"
                });

            // RAM
            if (r.Ram.UsageFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = r.Ram.PercentUsed > 95 ? Severity.Critical : Severity.Warning,
                    Category = "RAM",
                    Description = $"RAM usage is {r.Ram.PercentUsed}% ({DiagnosticFormatters.FormatMB(r.Ram.UsedMB)} / {DiagnosticFormatters.FormatMB(r.Ram.TotalMB)})",
                    Recommendation = "Close unused applications to free memory"
                });

            if (r.Ram.InsufficientRam)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "RAM",
                    Description = $"Total RAM is only {DiagnosticFormatters.FormatMB(r.Ram.TotalMB)} — potentially insufficient",
                    Recommendation = "Consider upgrading to at least 16 GB RAM"
                });

            // Disk
            foreach (var drive in r.Disk.Drives.Where(d => d.UsageFlagged))
                issues.Add(new FlaggedIssue
                {
                    Severity = drive.PercentUsed > 95 ? Severity.Critical : Severity.Warning,
                    Category = "Disk",
                    Description = $"Drive {drive.DriveLetter} is {drive.PercentUsed}% full ({DiagnosticFormatters.FormatMB(drive.FreeMB)} free)",
                    Recommendation = "Delete unnecessary files or move data to external storage"
                });

            foreach (var drive in r.Disk.Drives.Where(d =>
                !string.IsNullOrEmpty(d.HealthStatus) &&
                d.HealthStatus != "Healthy" &&
                d.HealthStatus != "Unknown"))
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Critical,
                    Category = "Disk",
                    Description = $"Drive {drive.DriveLetter} health is '{drive.HealthStatus}'",
                    Recommendation = "Back up data immediately and consider replacing the drive"
                });

            if (r.Disk.DiskActivityFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Disk",
                    Description = $"Disk activity at {r.Disk.DiskActivityPercent}% while idle (above 50%)",
                    Recommendation = "Check for background processes causing high disk I/O"
                });

            if (r.Disk.WindowsOldExists)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Disk",
                    Description = $"Windows.old folder exists ({DiagnosticFormatters.FormatMB(r.Disk.WindowsOldMB)})",
                    Recommendation = "Use Disk Cleanup to remove old Windows installation files"
                });

            // Thermal (CPU)
            if (r.Cpu.TemperatureFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Critical,
                    Category = "CPU",
                    Description = $"CPU temperature is {r.Cpu.CpuTemperatureC}°C (above 80°C)",
                    Recommendation = "Check cooling system, clean vents, and reduce CPU-intensive tasks"
                });

            if (r.Cpu.IsThrottling)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Critical,
                    Category = "CPU",
                    Description = "CPU is thermal throttling — performance is being reduced",
                    Recommendation = "Improve cooling immediately to prevent hardware damage"
                });

            // GPU
            if (r.Gpu.TemperatureFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Critical,
                    Category = "GPU",
                    Description = $"GPU temperature is {r.Gpu.GpuTemperatureC}°C (above 85°C)",
                    Recommendation = "Check GPU cooling, clean dust from fans and heatsink"
                });

            if (r.Gpu.UsageFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "GPU",
                    Description = $"GPU usage is {r.Gpu.GpuUsagePercent}% (above 90%)",
                    Recommendation = "Close GPU-intensive applications or check for crypto mining processes"
                });

            if (r.Gpu.DriverOutdated)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "GPU",
                    Description = $"GPU driver is over 1 year old (dated {r.Gpu.DriverDate})",
                    Recommendation = "Update GPU driver from the manufacturer's website"
                });

            if (r.Gpu.AdapterStatus != "OK" && r.Gpu.AdapterStatus != "Unknown")
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "GPU",
                    Description = $"GPU adapter status is '{r.Gpu.AdapterStatus}'",
                    Recommendation = "Check Device Manager for GPU errors and reinstall drivers if needed"
                });

            // Battery
            if (r.Battery.HasBattery && r.Battery.HealthFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Battery",
                    Description = $"Battery health is {r.Battery.HealthPercent}% (below 50%)",
                    Recommendation = "Consider replacing the battery"
                });

            // Startup
            if (r.Startup.TooManyFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Startup",
                    Description = $"{r.Startup.EnabledCount} startup programs enabled (above 10)",
                    Recommendation = "Disable unnecessary startup programs via Task Manager"
                });

            // Outlook
            foreach (var df in r.Outlook.DataFiles.Where(f => f.SizeFlagged))
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Outlook",
                    Description = $"Outlook data file is {DiagnosticFormatters.FormatMB(df.SizeMB)} (above 5 GB)",
                    Recommendation = "Archive old emails to reduce data file size"
                });

            // User Profile
            if (r.UserProfile.DesktopItemsFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "User Profile",
                    Description = $"{r.UserProfile.DesktopItemCount} items on desktop",
                    Recommendation = "Reduce desktop items to improve Explorer performance"
                });

            if (r.UserProfile.CorruptionDetected)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Critical,
                    Category = "User Profile",
                    Description = "User profile shows signs of corruption (temp profile detected)",
                    Recommendation = "Contact IT support to repair or recreate the user profile"
                });

            // Office
            if (r.Office.OfficeInstalled && r.Office.RepairNeeded)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Office",
                    Description = "Office may need repair",
                    Recommendation = "Run Office Online Repair from Settings > Apps"
                });

            // Antivirus
            if (r.Antivirus.AntivirusName == "Unable to detect" || r.Antivirus.AntivirusName == "Unknown")
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Critical,
                    Category = "Antivirus",
                    Description = "No antivirus product detected",
                    Recommendation = "Install and enable Windows Defender or a third-party antivirus"
                });

            if (r.Antivirus.BitLockerStatus == "Not Encrypted")
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Security",
                    Description = "C: drive is not encrypted with BitLocker",
                    Recommendation = "Enable BitLocker to protect data at rest"
                });

            // Windows Update
            if (r.WindowsUpdate.PendingUpdates.Contains("reboot", StringComparison.OrdinalIgnoreCase))
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "WindowsUpdate",
                    Description = "System requires a reboot to complete pending updates",
                    Recommendation = "Restart the computer to finish installing updates"
                });

            if (r.WindowsUpdate.ServiceStartType == "Disabled")
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "WindowsUpdate",
                    Description = "Windows Update service is disabled",
                    Recommendation = "Enable the Windows Update service to receive security patches"
                });

            // Network
            if (r.Network.ConnectionType == "WiFi" && r.Network.WifiBand == "2.4 GHz")
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Network",
                    Description = "WiFi is connected on 2.4 GHz band",
                    Recommendation = "Connect to a 5 GHz network for better speed if available"
                });

            // DNS response
            int dnsRating = DiagnosticFormatters.RateDns(r.Network.DnsResponseTime);
            if (dnsRating > 0)
                issues.Add(new FlaggedIssue
                {
                    Severity = dnsRating == 2 ? Severity.Warning : Severity.Info,
                    Category = "Network",
                    Description = $"DNS response time is {(dnsRating == 2 ? "poor" : "fair")} ({r.Network.DnsResponseTime})",
                    Recommendation = "Slow DNS may indicate network congestion or resolver issues — contact IT if browsing feels sluggish"
                });

            if (r.Network.PingLatency != "Unknown" && r.Network.PingLatency != "Failed")
            {
                string digits = new string(r.Network.PingLatency.Where(char.IsDigit).ToArray());
                if (int.TryParse(digits, out int latency) && latency > 100)
                    issues.Add(new FlaggedIssue
                    {
                        Severity = latency > 300 ? Severity.Warning : Severity.Info,
                        Category = "Network",
                        Description = $"Ping latency is {r.Network.PingLatency} (above 100 ms)",
                        Recommendation = "High latency may cause slow network performance. Check connection quality."
                    });
            }

            if (r.Network.SpeedTestFlagged)
                issues.Add(new FlaggedIssue
                {
                    Severity = r.Network.DownloadSpeedMbps < 5 ? Severity.Critical : Severity.Warning,
                    Category = "Network",
                    Description = $"Download speed is only {r.Network.DownloadSpeedMbps:F1} Mbps",
                    Recommendation = "Slow internet detected. Check your connection, try a wired connection, or contact your ISP."
                });

            if (r.Network.UploadSpeedMbps > 0 && r.Network.UploadSpeedMbps < 5)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Network",
                    Description = $"Upload speed is only {r.Network.UploadSpeedMbps:F1} Mbps",
                    Recommendation = "Slow upload may affect video calls, file uploads, and cloud sync."
                });

            // Browser
            foreach (var b in r.Browser.Browsers)
            {
                if (b.OpenTabs > 30)
                    issues.Add(new FlaggedIssue
                    {
                        Severity = Severity.Info,
                        Category = "Browser",
                        Description = $"{b.Name} has {b.OpenTabs} tabs open",
                        Recommendation = "Close unused tabs to free memory"
                    });
                if (b.CacheSizeMB > 1024)
                    issues.Add(new FlaggedIssue
                    {
                        Severity = Severity.Info,
                        Category = "Browser",
                        Description = $"{b.Name} cache is {DiagnosticFormatters.FormatMB(b.CacheSizeMB)}",
                        Recommendation = "Clear browser cache to free disk space"
                    });
            }

            // Temp folder sizes
            long totalTempMB = r.Disk.WindowsTempMB + r.Disk.UserTempMB;
            if (totalTempMB > 1024)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Disk",
                    Description = $"Temp folders using {DiagnosticFormatters.FormatMB(totalTempMB)}",
                    Recommendation = "Use Disk Cleanup to clear temporary files"
                });

            if (r.Disk.RecycleBinMB > 1024)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Disk",
                    Description = $"Recycle Bin is {DiagnosticFormatters.FormatMB(r.Disk.RecycleBinMB)}",
                    Recommendation = "Empty the Recycle Bin to free disk space"
                });

            if (r.Disk.SoftwareDistributionMB > 2048)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Disk",
                    Description = $"Windows Update cache is {DiagnosticFormatters.FormatMB(r.Disk.SoftwareDistributionMB)}",
                    Recommendation = "Run Disk Cleanup to clear Windows Update cache"
                });

            if (r.Disk.UpgradeLogsMB > 200)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Disk",
                    Description = $"Windows upgrade logs using {DiagnosticFormatters.FormatMB(r.Disk.UpgradeLogsMB)}",
                    Recommendation = "Run optimization to clear upgrade log files"
                });

            // Visual Settings
            if (r.VisualSettings.TransparencyEnabled && r.VisualSettings.AnimationsEnabled)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Visual Settings",
                    Description = "Transparency and animations are enabled — consuming extra CPU/GPU resources",
                    Recommendation = "Switch to performance visuals to reduce overhead"
                });

            // Antivirus scan impact
            if (r.Antivirus.FullScanRunning)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Antivirus",
                    Description = "A full antivirus scan is currently running — this may slow down the system",
                    Recommendation = "Wait for the scan to complete or schedule scans during idle hours"
                });

            // Browser extensions
            foreach (var b in r.Browser.Browsers)
            {
                if (b.ExtensionCount > 15)
                    issues.Add(new FlaggedIssue
                    {
                        Severity = Severity.Info,
                        Category = "Browser",
                        Description = $"{b.Name} has {b.ExtensionCount} extensions installed",
                        Recommendation = "Remove unused extensions to improve browser performance and security"
                    });
            }

            // Installed Software
            foreach (var eol in r.InstalledSoftware.EOLApps)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Software",
                    Description = $"'{eol.Name}' is end-of-life / unsupported",
                    Recommendation = "Uninstall EOL software to reduce security vulnerabilities"
                });

            foreach (var bloat in r.InstalledSoftware.BloatwareApps)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Software",
                    Description = $"'{bloat.Name}' is potential bloatware",
                    Recommendation = "Consider uninstalling to improve system performance"
                });

            if (r.InstalledSoftware.RuntimeCount > 6)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Info,
                    Category = "Software",
                    Description = $"{r.InstalledSoftware.RuntimeCount} Visual C++ Redistributables installed",
                    Recommendation = "Older versions may be safely removed if no legacy apps depend on them"
                });

            // Network Drives
            foreach (var nd in r.NetworkDrives.Drives)
            {
                if (!nd.IsAccessible)
                    issues.Add(new FlaggedIssue
                    {
                        Severity = Severity.Critical,
                        Category = "Network Drive",
                        Description = $"Network drive {nd.DriveLetter} ({nd.UncPath}) is unreachable",
                        Recommendation = "Check network connectivity and ensure the file server is online"
                    });
                else if (nd.LatencyFlagged)
                    issues.Add(new FlaggedIssue
                    {
                        Severity = Severity.Warning,
                        Category = "Network Drive",
                        Description = $"Network drive {nd.DriveLetter} has high latency ({nd.LatencyMs}ms)",
                        Recommendation = "Check network congestion or switch to a wired connection"
                    });
                if (nd.SpaceFlagged)
                    issues.Add(new FlaggedIssue
                    {
                        Severity = Severity.Warning,
                        Category = "Network Drive",
                        Description = $"Network drive {nd.DriveLetter} is {nd.PercentUsed}% full ({DiagnosticFormatters.FormatMB(nd.FreeMB)} free)",
                        Recommendation = "Contact IT to request additional storage or archive old files"
                    });
            }

            // Event Log
            if (r.EventLog.BSODs.Count > 0)
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Critical,
                    Category = "Event Log",
                    Description = $"{r.EventLog.BSODs.Count} BSOD(s) in the last 30 days (latest: {r.EventLog.BSODs[0].Timestamp:g})",
                    Recommendation = "Investigate blue screen crash dumps — may indicate driver or hardware failure"
                });

            if (r.EventLog.UnexpectedShutdowns.Count > 0)
                issues.Add(new FlaggedIssue
                {
                    Severity = r.EventLog.UnexpectedShutdowns.Count >= 3 ? Severity.Critical : Severity.Warning,
                    Category = "Event Log",
                    Description = $"{r.EventLog.UnexpectedShutdowns.Count} unexpected shutdown(s) in the last 30 days (latest: {r.EventLog.UnexpectedShutdowns[0].Timestamp:g})",
                    Recommendation = "Check power supply, overheating, or forced shutdowns"
                });

            if (r.EventLog.DiskErrors.Count > 0)
                issues.Add(new FlaggedIssue
                {
                    Severity = r.EventLog.DiskErrors.Count >= 5 ? Severity.Critical : Severity.Warning,
                    Category = "Event Log",
                    Description = $"{r.EventLog.DiskErrors.Count} disk error(s) in the last 30 days (latest: {r.EventLog.DiskErrors[0].Timestamp:g})",
                    Recommendation = "Back up data and run chkdsk — disk may be failing"
                });

            if (r.EventLog.AppCrashes.Count >= 5)
            {
                var topCrashers = r.EventLog.GetTopCrashingApps(3);
                string crasherDetail = topCrashers.Count > 0
                    ? $" (top: {string.Join(", ", topCrashers.Select(c => $"{c.AppName} {c.Count}×"))})"
                    : "";
                issues.Add(new FlaggedIssue
                {
                    Severity = Severity.Warning,
                    Category = "Event Log",
                    Description = $"{r.EventLog.AppCrashes.Count} application crash(es) in the last 30 days{crasherDetail}",
                    Recommendation = topCrashers.Count > 0
                        ? $"Consider reinstalling or updating {topCrashers[0].AppName}"
                        : "Review crashing applications — consider reinstalling or updating them"
                });
            }

            // Sort: Critical first, then Warning, then Info
            r.FlaggedIssues = issues.OrderBy(i => i.Severity switch
            {
                Severity.Critical => 0,
                Severity.Warning => 1,
                _ => 2
            }).ToList();
        }

        // ═══════════════════════════════════════════════════════════════
        //  HEALTH SCORE CALCULATOR
        // ═══════════════════════════════════════════════════════════════

        private static void CalculateHealthScore(DiagnosticResult r)
        {
            int score = 100;

            // CPU (max -8)
            if (r.Cpu.CpuLoadPercent > 50) score -= 8;
            else if (r.Cpu.CpuLoadPercent > 15) score -= 4;

            // RAM (max -12)
            if (r.Ram.PercentUsed > 95) score -= 10;
            else if (r.Ram.PercentUsed > 85) score -= 5;
            if (r.Ram.InsufficientRam) score -= 2;

            // Disk (max ~18)
            foreach (var d in r.Disk.Drives)
            {
                if (!string.IsNullOrEmpty(d.HealthStatus) &&
                    d.HealthStatus != "Healthy" && d.HealthStatus != "Unknown")
                    score -= 8;
                if (d.PercentUsed > 95) score -= 6;
                else if (d.PercentUsed > 85) score -= 3;
            }
            if (r.Disk.DiskActivityFlagged) score -= 2;

            // Thermal (max -15)
            if (r.Cpu.IsThrottling) score -= 10;
            else if (r.Cpu.TemperatureFlagged) score -= 5;

            // GPU (max -10)
            if (r.Gpu.TemperatureFlagged) score -= 5;
            if (r.Gpu.UsageFlagged) score -= 2;
            if (r.Gpu.DriverOutdated) score -= 3;

            // Battery (max -5)
            if (r.Battery.HasBattery && r.Battery.HealthFlagged) score -= 5;

            // Security (max -17)
            if (r.Antivirus.AntivirusName == "Unable to detect" ||
                r.Antivirus.AntivirusName == "Unknown")
                score -= 10;
            if (r.Antivirus.BitLockerStatus == "Not Encrypted") score -= 7;

            // Uptime (max -5)
            if (r.SystemOverview.Uptime.TotalDays > 14) score -= 5;
            else if (r.SystemOverview.UptimeFlagged) score -= 2;

            // Network (max -8)
            if (r.Network.PingLatency != "Unknown" && r.Network.PingLatency != "Failed")
            {
                string digits = new string(r.Network.PingLatency.Where(char.IsDigit).ToArray());
                if (int.TryParse(digits, out int latency))
                {
                    if (latency > 300) score -= 4;
                    else if (latency > 100) score -= 2;
                }
            }
            if (r.Network.SpeedTestFlagged)
                score -= r.Network.DownloadSpeedMbps < 5 ? 4 : 2;

            // Network Drives (max -8)
            foreach (var nd in r.NetworkDrives.Drives)
            {
                if (!nd.IsAccessible) score -= 4;
                else if (nd.LatencyFlagged) score -= 2;
                if (nd.SpaceFlagged) score -= 2;
            }

            // Windows Update (max -8)
            if (r.WindowsUpdate.PendingUpdates.Contains("reboot", StringComparison.OrdinalIgnoreCase))
                score -= 5;
            // Only penalize if the service is truly disabled — wuauserv is normally
            // Stopped when idle (demand-start) and that's expected behavior
            if (r.WindowsUpdate.ServiceStartType == "Disabled") score -= 3;

            // Startup (max -3)
            if (r.Startup.TooManyFlagged) score -= 3;

            // Software (max -5)
            score -= Math.Min(r.InstalledSoftware.EOLApps.Count * 2, 4);
            if (r.InstalledSoftware.BloatwareApps.Count > 0) score -= 1;

            // User Profile (max -6)
            if (r.UserProfile.CorruptionDetected) score -= 5;
            if (r.UserProfile.DesktopItemsFlagged) score -= 1;

            // Outlook (max -2)
            if (r.Outlook.DataFiles.Any(f => f.SizeFlagged)) score -= 2;

            // Event Log (max -15)
            if (r.EventLog.BSODs.Count > 0) score -= Math.Min(r.EventLog.BSODs.Count * 4, 10);
            if (r.EventLog.DiskErrors.Count >= 5) score -= 4;
            else if (r.EventLog.DiskErrors.Count > 0) score -= 2;
            if (r.EventLog.UnexpectedShutdowns.Count >= 3) score -= 5;
            else if (r.EventLog.UnexpectedShutdowns.Count > 0) score -= 2;

            r.HealthScore = Math.Clamp(score, 0, 100);
        }

        // ═══════════════════════════════════════════════════════════════
        //  HELPER METHODS
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Invokes OnLog inside a try/catch so a broken subscriber cannot crash the scan loop.
        /// </summary>
        private void SafeLog(string message)
        {
            try { OnLog?.Invoke(message); }
            catch { }
        }

        /// <summary>
        /// Invokes OnProgress inside a try/catch so a broken subscriber cannot crash the scan loop.
        /// </summary>
        private void SafeProgress(ScanProgress progress)
        {
            try { OnProgress?.Invoke(progress); }
            catch { }
        }

        private static List<ProcessInfo> GetTopProcessesByMemory(int count)
        {
            var list = new List<ProcessInfo>();
            var procs = Process.GetProcesses();
            try
            {
                var top = procs
                    .Where(p => { try { return !string.IsNullOrEmpty(p.ProcessName) && p.WorkingSet64 > 0; } catch { return false; } })
                    .OrderByDescending(p => { try { return p.WorkingSet64; } catch { return 0L; } })
                    .Take(count);

                foreach (var p in top)
                {
                    try
                    {
                        list.Add(new ProcessInfo
                        {
                            Name = p.ProcessName,
                            Id = p.Id,
                            MemoryMB = p.WorkingSet64 / (1024 * 1024)
                        });
                    }
                    catch { }
                }
            }
            catch { }
            finally { foreach (var p in procs) { try { p.Dispose(); } catch { } } }
            return list;
        }

        private static long GetFolderSizeMB(string path, int maxDepth = int.MaxValue)
            => NativeHelpers.GetFolderSizeMB(path, maxDepth);

        private static long GetRecycleBinSizeMB() => NativeHelpers.GetRecycleBinSizeMB();

        /// <summary>
        /// Queries a Windows service's runtime status and start type in a single atomic snapshot.
        /// Calls Refresh() to ensure the state is current, not cached from a previous read.
        /// Returns consistent (status, startType) tuple — never mixes stale and fresh values.
        /// </summary>
        private static (string Status, string StartType) QueryServiceState(string serviceName)
        {
            try
            {
                using var sc = new System.ServiceProcess.ServiceController(serviceName);
                sc.Refresh();
                string status = sc.Status.ToString();
                string startType = sc.StartType.ToString();
                return (status, startType);
            }
            catch (InvalidOperationException)
            {
                return ("Not Installed", "Unknown");
            }
            catch
            {
                return ("Unable to check", "Unknown");
            }
        }

        private static string GetDriveMediaType(string driveName)
        {
            try
            {
                string letter = driveName.TrimEnd('\\').TrimEnd(':');

                // Map drive letter → partition → physical disk via WMI associations
                try
                {
                    using var partSearch = new ManagementObjectSearcher(
                        $"ASSOCIATORS OF {{Win32_LogicalDisk.DeviceID='{letter}:'}} WHERE AssocClass=Win32_LogicalDiskToPartition");
                    partSearch.Options.Timeout = TimeSpan.FromSeconds(5);
                    foreach (var partition in partSearch.Get())
                    {
                        string partId = partition["DeviceID"]?.ToString();
                        if (string.IsNullOrEmpty(partId)) continue;

                        using var diskSearch = new ManagementObjectSearcher(
                            $"ASSOCIATORS OF {{Win32_DiskPartition.DeviceID='{partId}'}} WHERE AssocClass=Win32_DiskDriveToDiskPartition");
                        diskSearch.Options.Timeout = TimeSpan.FromSeconds(5);
                        foreach (var disk in diskSearch.Get())
                        {
                            int diskIndex = Convert.ToInt32(disk["Index"] ?? -1);
                            if (diskIndex < 0) continue;

                            // Query MSFT_PhysicalDisk with matching DeviceId
                            try
                            {
                                using var storageSearch = new ManagementObjectSearcher(
                                    @"\\.\root\microsoft\windows\storage",
                                    $"SELECT MediaType FROM MSFT_PhysicalDisk WHERE DeviceId='{diskIndex}'");
                                storageSearch.Options.Timeout = TimeSpan.FromSeconds(5);
                                foreach (var obj in storageSearch.Get())
                                {
                                    int mediaType = Convert.ToInt32(obj["MediaType"]);
                                    return mediaType switch
                                    {
                                        3 => "HDD",
                                        4 => "SSD",
                                        5 => "SCM",
                                        _ => "Unknown"
                                    };
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }

                // Fallback: if only one physical disk, return its type
                try
                {
                    using var storageSearch = new ManagementObjectSearcher(
                        @"\\.\root\microsoft\windows\storage",
                        "SELECT MediaType FROM MSFT_PhysicalDisk");
                    storageSearch.Options.Timeout = TimeSpan.FromSeconds(5);
                    var disks = storageSearch.Get();
                    if (disks.Count == 1)
                    {
                        foreach (var obj in disks)
                        {
                            int mediaType = Convert.ToInt32(obj["MediaType"]);
                            return mediaType switch
                            {
                                3 => "HDD",
                                4 => "SSD",
                                5 => "SCM",
                                _ => "Unknown"
                            };
                        }
                    }
                }
                catch { }
            }
            catch { }
            return "Unknown";
        }

        private static string GetDriveHealthStatus(string driveName)
        {
            try
            {
                string letter = driveName.TrimEnd('\\').TrimEnd(':');

                // Map drive letter → partition → physical disk to get the correct disk's status
                try
                {
                    using var partSearch = new ManagementObjectSearcher(
                        $"ASSOCIATORS OF {{Win32_LogicalDisk.DeviceID='{letter}:'}} WHERE AssocClass=Win32_LogicalDiskToPartition");
                    partSearch.Options.Timeout = TimeSpan.FromSeconds(5);
                    foreach (var partition in partSearch.Get())
                    {
                        string partId = partition["DeviceID"]?.ToString();
                        if (string.IsNullOrEmpty(partId)) continue;

                        using var diskSearch = new ManagementObjectSearcher(
                            $"ASSOCIATORS OF {{Win32_DiskPartition.DeviceID='{partId}'}} WHERE AssocClass=Win32_DiskDriveToDiskPartition");
                        diskSearch.Options.Timeout = TimeSpan.FromSeconds(5);
                        foreach (var disk in diskSearch.Get())
                        {
                            string status = disk["Status"]?.ToString() ?? "Unknown";
                            return status == "OK" ? "Healthy" : status;
                        }
                    }
                }
                catch { }

                // Fallback: if only one physical disk, return its status
                using var searcher = new ManagementObjectSearcher("SELECT Status FROM Win32_DiskDrive");
                searcher.Options.Timeout = TimeSpan.FromSeconds(5);
                var disks = searcher.Get();
                if (disks.Count == 1)
                {
                    foreach (var obj in disks)
                    {
                        string status = obj["Status"]?.ToString() ?? "Unknown";
                        return status == "OK" ? "Healthy" : status;
                    }
                }
            }
            catch { }
            return "Unknown";
        }

        private static string GetPowerPlan()
        {
            try
            {
                if (!TryRunProcessForOutput("powercfg.exe", "/getactivescheme", 3000, out string output))
                    return "Unknown";

                int start = output.LastIndexOf('(');
                int end = output.LastIndexOf(')');
                if (start >= 0 && end > start)
                    return output[(start + 1)..end];
                return "Unknown";
            }
            catch { return "Unknown"; }
        }

        // ═══════════════════════════════════════════════════════════════
        //  P/INVOKE
        // ═══════════════════════════════════════════════════════════════

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private class MEMORYSTATUSEX
        {
            public uint dwLength;
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
            public MEMORYSTATUSEX() { dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX)); }
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX lpBuffer);

        // ── Display enumeration P/Invoke ──

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool EnumDisplayDevices(string lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool EnumDisplaySettings(string lpszDeviceName, int iModeNum, ref DEVMODE lpDevMode);

        private const int ENUM_CURRENT_SETTINGS = -1;
        private const int DISPLAY_DEVICE_ACTIVE = 0x1;
        private const int DISPLAY_DEVICE_PRIMARY_DEVICE = 0x4;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct DISPLAY_DEVICE
        {
            public int cb;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string DeviceName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string DeviceString;
            public int StateFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string DeviceID;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string DeviceKey;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct DEVMODE
        {
            private const int CCHDEVICENAME = 32;
            private const int CCHFORMNAME = 32;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHDEVICENAME)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public int dmPositionX;
            public int dmPositionY;
            public int dmDisplayOrientation;
            public int dmDisplayFixedOutput;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHFORMNAME)]
            public string dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmReserved1;
            public int dmReserved2;
            public int dmPanningWidth;
            public int dmPanningHeight;
        }

        /// <summary>Enumerates all active physical displays using Win32 EnumDisplayDevices + EnumDisplaySettings.</summary>
        private static List<DisplayInfo> EnumerateDisplays()
        {
            var displays = new List<DisplayInfo>();
            var device = new DISPLAY_DEVICE();
            device.cb = Marshal.SizeOf(device);

            for (uint i = 0; EnumDisplayDevices(null, i, ref device, 0); i++)
            {
                if ((device.StateFlags & DISPLAY_DEVICE_ACTIVE) == 0)
                {
                    device.cb = Marshal.SizeOf(device);
                    continue;
                }

                bool isPrimary = (device.StateFlags & DISPLAY_DEVICE_PRIMARY_DEVICE) != 0;
                string adapterName = device.DeviceName;

                // Get the monitor name from the second-level device
                var monitor = new DISPLAY_DEVICE();
                monitor.cb = Marshal.SizeOf(monitor);
                string monitorName = "Unknown Monitor";
                if (EnumDisplayDevices(adapterName, 0, ref monitor, 0))
                    monitorName = string.IsNullOrWhiteSpace(monitor.DeviceString) ? "Unknown Monitor" : monitor.DeviceString;

                // Get current display mode
                var devMode = new DEVMODE();
                devMode.dmSize = (short)Marshal.SizeOf(devMode);
                if (EnumDisplaySettings(adapterName, ENUM_CURRENT_SETTINGS, ref devMode))
                {
                    displays.Add(new DisplayInfo
                    {
                        DeviceName = adapterName,
                        MonitorName = monitorName,
                        Width = devMode.dmPelsWidth,
                        Height = devMode.dmPelsHeight,
                        RefreshRate = devMode.dmDisplayFrequency,
                        IsPrimary = isPrimary
                    });
                }

                device.cb = Marshal.SizeOf(device);
            }

            return displays;
        }

        // ═══════════════════════════════════════════════════════════════
        //  DISPOSAL
        // ═══════════════════════════════════════════════════════════════

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing)
            {
                _cts?.Cancel();
                _cts?.Dispose();
            }
            _disposed = true;
        }
        private static string GetPhaseDescription(string phase) => phase switch
        {
            "System Overview" => "Reading hardware info, OS version, and uptime",
            "CPU Diagnostics" => "Measuring CPU load, temperature, and throttling",
            "RAM Diagnostics" => "Checking memory usage and top consumers",
            "GPU Diagnostics" => "Detecting GPUs, drivers, and display configuration",
            "Disk Diagnostics" => "Scanning drives, temp files, and disk health",
            "Thermal Diagnostics" => "Reading system temperatures and fan status",
            "Battery Health" => "Checking battery capacity and power plan",
            "Startup Programs" => "Enumerating startup items and their impact",
            "Visual & Display" => "Checking visual effects and transparency settings",
            "Antivirus & Security" => "Detecting antivirus and BitLocker status",
            "Windows Update" => "Checking update service and pending patches",
            "Network Diagnostics" => "Testing DNS, latency, and adapter speed",
            "Internet Speed Test" => "Measuring download and upload throughput",
            "Network Drives" => "Checking mapped drive accessibility and latency",
            "Outlook / Email" => "Scanning Outlook data files and add-ins",
            "Browser Check" => "Counting tabs, cache, and extensions",
            "User Profile" => "Measuring profile size and desktop items",
            "Office Diagnostics" => "Checking Office version and repair status",
            "Installed Software" => "Scanning for bloatware, EOL, and runtimes",
            "Event Log Analysis" => "Searching for BSODs, crashes, and disk errors",
            _ => ""
        };

        private static int GetPhaseEstimate(string phase) => phase switch
        {
            "CPU Diagnostics" => 4,
            "Internet Speed Test" => 8,
            "Network Diagnostics" => 5,
            "Disk Diagnostics" => 3,
            "GPU Diagnostics" => 2,
            "Event Log Analysis" => 2,
            _ => 1
        };

        /// <summary>
        /// Per-phase timeout in milliseconds. Phases that query hang-prone WMI
        /// namespaces (root\WMI, root\microsoft\windows\storage) get shorter
        /// timeouts so a single hung provider doesn't cascade into a 20-minute scan.
        /// </summary>
        private static int GetPhaseTimeoutMs(string phase) => phase switch
        {
            "Thermal Diagnostics" => 15_000,  // root\WMI hangs on some hardware
            "Disk Diagnostics"    => 30_000,  // root\microsoft\windows\storage hangs on some hardware
            "Internet Speed Test" => 45_000,  // network-bound, but has its own 30s HttpClient timeout
            _                     => 60_000
        };
    }

    // ═══════════════════════════════════════════════════════════════
    //  SCAN PROGRESS (kept for UI compatibility)
    // ═══════════════════════════════════════════════════════════════

    public class ScanProgress
    {
        public int PhaseIndex { get; set; }
        public int Total { get; set; }
        public string CurrentPhase { get; set; } = "Initializing...";
        public string PhaseDescription { get; set; } = "";
        public int EstimatedSeconds { get; set; }
    }
}
