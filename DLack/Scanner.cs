using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace DLack
{
    public class Scanner : IDisposable
    {
        // ── Events ───────────────────────────────────────────────────
        public event Action<ScanProgress> OnProgress;
        public event Action<string> OnLog;

        // ── State ────────────────────────────────────────────────────
        private CancellationTokenSource _cts;
        private PerformanceCounter _cpuCounter;
        private volatile bool _cpuCounterReady;
        private bool _disposed;
        private string _cachedTopCpuProcess = "N/A";
        private string _cachedTopRamProcess = "N/A";
        private int _processPollCounter;
        private readonly Dictionary<int, TimeSpan> _lastProcessCpuTimes = new();
        private DateTime _lastProcessSnapshotUtc = DateTime.MinValue;

        // ═══════════════════════════════════════════════════════════════
        //  PUBLIC API
        // ═══════════════════════════════════════════════════════════════

        public async Task<ScanResult> RunScan(int durationSeconds = 60)
        {
            if (durationSeconds <= 0)
                durationSeconds = 1;

            _cts?.Dispose();
            _cts = new CancellationTokenSource();

            var samples = new List<SystemMetrics>(durationSeconds);
            bool wasCancelled = false;

            OnLog?.Invoke($"Starting {durationSeconds}-second system scan...");

            // Initialize CPU counter on a background thread — never blocks the scan
            if (!_cpuCounterReady)
                _ = Task.Run(InitCpuCounterAsync);

            for (int i = 0; i < durationSeconds; i++)
            {
                if (_cts.Token.IsCancellationRequested)
                {
                    OnLog?.Invoke($"Scan cancelled at {i} seconds.");
                    wasCancelled = true;
                    break;
                }

                var metrics = CaptureMetrics();
                samples.Add(metrics);

                OnProgress?.Invoke(new ScanProgress
                {
                    Elapsed = i + 1,
                    Total = durationSeconds,
                    CurrentCpu = metrics.CpuPercent,
                    CurrentRam = metrics.RamPercent,
                    TopCpuProcess = metrics.TopCpuProcess,
                    TopRamProcess = metrics.TopRamProcess
                });

                if ((i + 1) % 10 == 0)
                    OnLog?.Invoke($"Sample {i + 1}/{durationSeconds} — CPU {metrics.CpuPercent}% | RAM {metrics.RamPercent}%");

                await SafeDelay(1000);
            }

            if (_cts.Token.IsCancellationRequested)
                wasCancelled = true;

            if (wasCancelled)
            {
                OnLog?.Invoke("Scan was cancelled. Discarding partial samples.");
                return new ScanResult();
            }

            OnLog?.Invoke("Scan complete. Analyzing results...");
            return AnalyzeSamples(samples);
        }

        public void CancelScan()
        {
            try { _cts?.Cancel(); }
            catch (ObjectDisposedException) { }
        }

        // ═══════════════════════════════════════════════════════════════
        //  METRICS CAPTURE
        // ═══════════════════════════════════════════════════════════════

        private void InitCpuCounterAsync()
        {
            if (_cpuCounterReady) return;

            try
            {
                var counter = new PerformanceCounter("Processor", "% Processor Time", "_Total", true);
                counter.NextValue(); // Prime it (first read is always 0)
                System.Threading.Thread.Sleep(500);
                counter.NextValue(); // Second read gives real data

                _cpuCounter = counter;
                _cpuCounterReady = true;
            }
            catch
            {
                // Counter unavailable — fallback will be used
                _cpuCounterReady = false;
            }
        }

        private SystemMetrics CaptureMetrics()
        {
            var metrics = new SystemMetrics();

            try
            {
                // ── CPU ──
                metrics.CpuPercent = GetCpuUsage();

                // ── RAM (native call — fast and reliable) ──
                var memStatus = new MEMORYSTATUSEX();
                if (GlobalMemoryStatusEx(memStatus))
                {
                    metrics.RamTotalMB = (long)(memStatus.ullTotalPhys / 1024 / 1024);
                    long availableMB = (long)(memStatus.ullAvailPhys / 1024 / 1024);
                    metrics.RamUsedMB = metrics.RamTotalMB - availableMB;
                    metrics.RamPercent = metrics.RamTotalMB > 0
                        ? Math.Round((double)metrics.RamUsedMB / metrics.RamTotalMB * 100, 1)
                        : 0;
                }

                // ── Disk ──
                var drive = new DriveInfo("C:");
                if (drive.IsReady)
                {
                    metrics.DiskFreeMB = drive.AvailableFreeSpace / 1024 / 1024;
                    metrics.DiskFreePercent = Math.Round(
                        (double)drive.AvailableFreeSpace / drive.TotalSize * 100, 1);
                }

                // ── Top processes (poll every 5s to reduce overhead) ──
                _processPollCounter++;
                if (_processPollCounter >= 5 || _cachedTopCpuProcess == "N/A" || _cachedTopRamProcess == "N/A")
                {
                    var (topCpu, topRam) = GetTopProcesses();
                    _cachedTopCpuProcess = topCpu;
                    _cachedTopRamProcess = topRam;
                    metrics.TopCpuProcess = topCpu;
                    metrics.TopRamProcess = topRam;
                    _processPollCounter = 0;
                }
                else
                {
                    metrics.TopCpuProcess = _cachedTopCpuProcess;
                    metrics.TopRamProcess = _cachedTopRamProcess;
                }
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"⚠ Metrics capture error: {ex.Message}");
            }

            return metrics;
        }

        private double GetCpuUsage()
        {
            try
            {
                if (_cpuCounterReady && _cpuCounter != null)
                {
                    double value = _cpuCounter.NextValue();
                    return Math.Round(Math.Min(value, 100), 1);
                }
            }
            catch { }

            // Fallback until counter is ready (or if unavailable)
            return GetCpuFallback();
        }

        private static double GetCpuFallback()
        {
            try
            {
                var mem = new MEMORYSTATUSEX();
                if (GlobalMemoryStatusEx(mem))
                    return Math.Round(mem.dwMemoryLoad * 0.7, 1); // rough proxy

                return 0;
            }
            catch { return 0; }
        }

        private (string topCpu, string topRam) GetTopProcesses()
        {
            string topCpu = "N/A";
            string topRam = "N/A";
            var snapshotUtc = DateTime.UtcNow;
            double elapsedMs = _lastProcessSnapshotUtc == DateTime.MinValue
                ? 0
                : (snapshotUtc - _lastProcessSnapshotUtc).TotalMilliseconds;
            var activeProcessIds = new HashSet<int>();

            try
            {
                var procs = Process.GetProcesses();
                try
                {
                    string ramName = "N/A";
                    long ramBytes = -1;
                    string cpuName = "N/A";
                    double cpuPercent = -1;

                    foreach (var p in procs)
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(p.ProcessName))
                                continue;

                            activeProcessIds.Add(p.Id);

                            long workingSet = p.WorkingSet64;
                            if (workingSet > ramBytes)
                            {
                                ramBytes = workingSet;
                                ramName = p.ProcessName;
                            }

                            TimeSpan totalCpu = p.TotalProcessorTime;
                            if (elapsedMs > 0 && _lastProcessCpuTimes.TryGetValue(p.Id, out var previousCpu))
                            {
                                double deltaMs = (totalCpu - previousCpu).TotalMilliseconds;
                                double currentCpu = deltaMs / (elapsedMs * Environment.ProcessorCount) * 100;
                                if (currentCpu > cpuPercent)
                                {
                                    cpuPercent = currentCpu;
                                    cpuName = p.ProcessName;
                                }
                            }

                            _lastProcessCpuTimes[p.Id] = totalCpu;
                        }
                        catch { }
                    }

                    if (ramBytes >= 0)
                        topRam = $"{ramName} ({ramBytes / (1024 * 1024)} MB)";

                    if (cpuPercent >= 0)
                        topCpu = $"{cpuName} ({cpuPercent:F1}% CPU)";
                    else if (topRam != "N/A")
                        topCpu = topRam;
                }
                finally
                {
                    // Dispose process objects to prevent handle leaks
                    foreach (var p in procs)
                    {
                        try { p.Dispose(); }
                        catch { }
                    }
                }
            }
            catch { }

            if (_lastProcessCpuTimes.Count > 0)
            {
                var staleIds = _lastProcessCpuTimes.Keys.Where(id => !activeProcessIds.Contains(id)).ToList();
                foreach (var staleId in staleIds)
                    _lastProcessCpuTimes.Remove(staleId);
            }

            _lastProcessSnapshotUtc = snapshotUtc;
            return (topCpu, topRam);
        }

        // ═══════════════════════════════════════════════════════════════
        //  ANALYSIS
        // ═══════════════════════════════════════════════════════════════

        private ScanResult AnalyzeSamples(List<SystemMetrics> samples)
        {
            if (samples.Count == 0)
                return new ScanResult();

            var last = samples[^1];

            var result = new ScanResult
            {
                SampleCount = samples.Count,
                DurationSeconds = samples.Count,

                // CPU
                AvgCpu = Math.Round(samples.Average(s => s.CpuPercent), 1),
                PeakCpu = Math.Round(samples.Max(s => s.CpuPercent), 1),

                // RAM
                AvgRam = Math.Round(samples.Average(s => s.RamPercent), 1),
                PeakRam = Math.Round(samples.Max(s => s.RamPercent), 1),
                RamTotalMB = last.RamTotalMB,
                RamUsedMB = last.RamUsedMB,

                // Disk
                DiskFreePercent = last.DiskFreePercent,
                DiskFreeMB = last.DiskFreeMB,

                // System info
                Uptime = TimeSpan.FromMilliseconds(Environment.TickCount64),
                PowerPlan = GetPowerPlan(),
                TempFolderSizeMB = GetTempSize(),
                BrowserProcessCount = GetBrowserCount(),

                // Top processes
                TopCpuProcesses = GetTopNProcessesByCpu(10),
                TopRamProcesses = GetTopNProcessesByMemory(10)
            };

            result.Recommendations = GenerateRecommendations(result);
            return result;
        }

        // ═══════════════════════════════════════════════════════════════
        //  SYSTEM INFO HELPERS
        // ═══════════════════════════════════════════════════════════════

        private static string GetPowerPlan()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powercfg.exe",
                    Arguments = "/getactivescheme",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };

                using var process = Process.Start(psi);
                if (process == null) return "Unknown";

                if (!process.WaitForExit(3000))
                {
                    try { process.Kill(entireProcessTree: true); }
                    catch { }
                    return "Unknown";
                }

                string output = process.StandardOutput.ReadToEnd();

                // Output format: "Power Scheme GUID: xxx  (Plan Name)"
                // Extract the name between parentheses
                int start = output.LastIndexOf('(');
                int end = output.LastIndexOf(')');
                if (start >= 0 && end > start)
                    return output[(start + 1)..end];

                return "Unknown";
            }
            catch { return "Unknown"; }
        }

        private static long GetTempSize()
        {
            try
            {
                string temp = Path.GetTempPath();
                if (!Directory.Exists(temp)) return 0;

                long totalBytes = 0;

                // Use EnumerateFiles for better perf, skip inaccessible dirs
                foreach (var file in Directory.EnumerateFiles(temp, "*", new EnumerationOptions
                {
                    IgnoreInaccessible = true,
                    RecurseSubdirectories = true
                }))
                {
                    try { totalBytes += new FileInfo(file).Length; }
                    catch { }
                }

                return totalBytes / (1024 * 1024);
            }
            catch { return 0; }
        }

        private static int GetBrowserCount()
        {
            try
            {
                string[] browsers = { "chrome", "msedge", "firefox", "brave", "opera" };
                var procs = Process.GetProcesses();

                int count = procs.Count(p =>
                {
                    try
                    {
                        return browsers.Any(b =>
                            p.ProcessName.Contains(b, StringComparison.OrdinalIgnoreCase));
                    }
                    catch { return false; }
                });

                foreach (var p in procs)
                {
                    try { p.Dispose(); }
                    catch { }
                }

                return count;
            }
            catch { return 0; }
        }

        private static List<ProcessInfo> GetTopNProcessesByMemory(int count)
        {
            var list = new List<ProcessInfo>();

            try
            {
                var procs = Process.GetProcesses();

                var top = procs
                    .Where(p =>
                    {
                        try { return !string.IsNullOrEmpty(p.ProcessName) && p.WorkingSet64 > 0; }
                        catch { return false; }
                    })
                    .OrderByDescending(p =>
                    {
                        try { return p.WorkingSet64; }
                        catch { return 0L; }
                    })
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

                foreach (var p in procs)
                {
                    try { p.Dispose(); }
                    catch { }
                }
            }
            catch { }

            return list;
        }

        private static List<ProcessInfo> GetTopNProcessesByCpu(int count)
        {
            var list = new List<ProcessInfo>();

            try
            {
                var procs = Process.GetProcesses();
                try
                {
                    var cpuSamples = new List<ProcessInfo>();

                    foreach (var p in procs)
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(p.ProcessName))
                                continue;

                            double uptimeSeconds = (DateTime.Now - p.StartTime).TotalSeconds;
                            if (uptimeSeconds <= 0.5)
                                continue;

                            double cpuPercent =
                                p.TotalProcessorTime.TotalSeconds /
                                (uptimeSeconds * Environment.ProcessorCount) * 100.0;

                            cpuSamples.Add(new ProcessInfo
                            {
                                Name = p.ProcessName,
                                Id = p.Id,
                                MemoryMB = p.WorkingSet64 / (1024 * 1024),
                                CpuPercent = Math.Round(Math.Max(cpuPercent, 0), 1)
                            });
                        }
                        catch { }
                    }

                    list = cpuSamples
                        .OrderByDescending(p => p.CpuPercent)
                        .ThenByDescending(p => p.MemoryMB)
                        .Take(count)
                        .ToList();
                }
                finally
                {
                    foreach (var p in procs)
                    {
                        try { p.Dispose(); }
                        catch { }
                    }
                }
            }
            catch { }

            return list;
        }

        // ═══════════════════════════════════════════════════════════════
        //  RECOMMENDATIONS
        // ═══════════════════════════════════════════════════════════════

        private static List<string> GenerateRecommendations(ScanResult result)
        {
            var recs = new List<string>();

            if (!string.Equals(result.PowerPlan, "High performance", StringComparison.OrdinalIgnoreCase))
                recs.Add("Set power plan to High Performance");

            if (result.AvgCpu >= 80)
                recs.Add("High CPU usage — check for runaway processes");
            else if (result.AvgCpu >= 60)
                recs.Add("Moderate CPU usage — consider closing background apps");

            if (result.AvgRam >= 90)
                recs.Add("Critical memory pressure — restart recommended");
            else if (result.AvgRam >= 80)
                recs.Add("High memory pressure — close unused applications");
            else if (result.AvgRam >= 70)
                recs.Add("Moderate memory usage detected");

            if (result.DiskFreePercent < 10)
                recs.Add("Critical: disk space below 10% — immediate cleanup needed");
            else if (result.DiskFreePercent < 20)
                recs.Add("Low disk space — cleanup recommended");

            if (result.Uptime.TotalDays >= 14)
                recs.Add("System uptime exceeds 14 days — reboot strongly recommended");
            else if (result.Uptime.TotalDays >= 7)
                recs.Add("System uptime exceeds 7 days — consider rebooting");

            if (result.BrowserProcessCount > 20)
                recs.Add($"{result.BrowserProcessCount} browser processes — consider closing tabs");

            if (result.TempFolderSizeMB > 500)
                recs.Add($"Temp folder is {result.TempFolderSizeMB} MB — clear temp files");
            else
                recs.Add("Clear temp files for minor cleanup");

            recs.Add("Apply Performance Visual Mode");

            return recs;
        }

        // ═══════════════════════════════════════════════════════════════
        //  UTILITY
        // ═══════════════════════════════════════════════════════════════

        /// <summary>Task.Delay that doesn't throw on cancellation.</summary>
        private async Task SafeDelay(int ms)
        {
            try
            {
                await Task.Delay(ms, _cts.Token);
            }
            catch (TaskCanceledException) { }
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

            public MEMORYSTATUSEX()
            {
                dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
            }
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX lpBuffer);

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
                _cpuCounter?.Dispose();
            }

            _disposed = true;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  DATA MODELS
    // ═══════════════════════════════════════════════════════════════

    public class SystemMetrics
    {
        public double CpuPercent { get; set; }
        public double RamPercent { get; set; }
        public long RamUsedMB { get; set; }
        public long RamTotalMB { get; set; }
        public double DiskFreePercent { get; set; }
        public long DiskFreeMB { get; set; }
        public string TopCpuProcess { get; set; } = "N/A";
        public string TopRamProcess { get; set; } = "N/A";
    }

    public class ScanProgress
    {
        public int Elapsed { get; set; }
        public int Total { get; set; }
        public double CurrentCpu { get; set; }
        public double CurrentRam { get; set; }
        public string TopCpuProcess { get; set; }
        public string TopRamProcess { get; set; }
    }

    public class ScanResult
    {
        public int SampleCount { get; set; }
        public int DurationSeconds { get; set; }
        public double AvgCpu { get; set; }
        public double PeakCpu { get; set; }
        public double AvgRam { get; set; }
        public double PeakRam { get; set; }
        public long RamTotalMB { get; set; }
        public long RamUsedMB { get; set; }
        public double DiskFreePercent { get; set; }
        public long DiskFreeMB { get; set; }
        public TimeSpan Uptime { get; set; }
        public string PowerPlan { get; set; }
        public long TempFolderSizeMB { get; set; }
        public int BrowserProcessCount { get; set; }
        public List<ProcessInfo> TopCpuProcesses { get; set; } = new();
        public List<ProcessInfo> TopRamProcesses { get; set; } = new();
        public List<string> Recommendations { get; set; } = new();
    }

    public class ProcessInfo
    {
        public string Name { get; set; }
        public int Id { get; set; }
        public long MemoryMB { get; set; }
        public double CpuPercent { get; set; }
    }
}