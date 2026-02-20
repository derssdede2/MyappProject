using System;
using System.Collections.Generic;
using System.Linq;

namespace DLack
{
    // ═══════════════════════════════════════════════════════════════
    //  SEVERITY LEVELS
    // ═══════════════════════════════════════════════════════════════

    public enum Severity
    {
        Info,
        Warning,
        Critical
    }

    // ═══════════════════════════════════════════════════════════════
    //  FLAGGED ISSUE
    // ═══════════════════════════════════════════════════════════════

    public class FlaggedIssue
    {
        public Severity Severity { get; set; }
        public string Category { get; set; } = "";
        public string Description { get; set; } = "";
        public string Recommendation { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    //  TOP-LEVEL DIAGNOSTIC RESULT
    // ═══════════════════════════════════════════════════════════════

    public class DiagnosticResult
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public double ScanDurationSeconds { get; set; }
        public int HealthScore { get; set; } = 100;

        /// <summary>
        /// The Windows user account under which the scan was performed.
        /// Results for HKCU, %LOCALAPPDATA%, and process enumeration reflect this user only.
        /// </summary>
        public string ScannedUser { get; set; } = "";

        // Categories
        public SystemOverview SystemOverview { get; set; } = new();
        public CpuDiagnostics Cpu { get; set; } = new();
        public RamDiagnostics Ram { get; set; } = new();
        public DiskDiagnostics Disk { get; set; } = new();
        public BatteryDiagnostics Battery { get; set; } = new();
        public StartupDiagnostics Startup { get; set; } = new();
        public VisualSettingsDiagnostics VisualSettings { get; set; } = new();
        public AntivirusDiagnostics Antivirus { get; set; } = new();
        public WindowsUpdateDiagnostics WindowsUpdate { get; set; } = new();
        public NetworkDiagnostics Network { get; set; } = new();
        public OutlookDiagnostics Outlook { get; set; } = new();
        public BrowserDiagnostics Browser { get; set; } = new();
        public UserProfileDiagnostics UserProfile { get; set; } = new();
        public OfficeDiagnostics Office { get; set; } = new();
        public InstalledSoftwareDiagnostics InstalledSoftware { get; set; } = new();
        public GpuDiagnostics Gpu { get; set; } = new();
        public NetworkDriveDiagnostics NetworkDrives { get; set; } = new();
        public EventLogDiagnostics EventLog { get; set; } = new();

        // Aggregated flags
        public List<FlaggedIssue> FlaggedIssues { get; set; } = new();

        // Optimization results — populated only after the optimizer has been run
        public List<OptimizationAction> OptimizationActions { get; set; }
        public OptimizationSummary OptimizationSummary { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  1. SYSTEM OVERVIEW
    // ═══════════════════════════════════════════════════════════════

    public class SystemOverview
    {
        public string ComputerName { get; set; } = "";
        public string Manufacturer { get; set; } = "";
        public string Model { get; set; } = "";
        public string SerialNumber { get; set; } = "";
        public string WindowsVersion { get; set; } = "";
        public string WindowsBuild { get; set; } = "";
        public string CpuModel { get; set; } = "";
        public string CpuClockSpeed { get; set; } = "";
        public long TotalRamMB { get; set; }
        public TimeSpan Uptime { get; set; }
        public bool UptimeFlagged { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  2. CPU DIAGNOSTICS
    // ═══════════════════════════════════════════════════════════════

    public class CpuDiagnostics
    {
        public double CpuLoadPercent { get; set; }
        public bool CpuLoadFlagged { get; set; }
        public List<ProcessInfo> TopCpuProcesses { get; set; } = new();

        // Thermal data (merged from former Thermal section)
        public double CpuTemperatureC { get; set; }
        public bool TemperatureFlagged { get; set; }
        public bool IsThrottling { get; set; }
        public string FanStatus { get; set; } = "Unknown";
    }

    // ═══════════════════════════════════════════════════════════════
    //  3. RAM DIAGNOSTICS
    // ═══════════════════════════════════════════════════════════════

    public class RamDiagnostics
    {
        public long TotalMB { get; set; }
        public long UsedMB { get; set; }
        public long AvailableMB { get; set; }
        public double PercentUsed { get; set; }
        public bool UsageFlagged { get; set; }
        public bool InsufficientRam { get; set; }
        public List<ProcessInfo> TopRamProcesses { get; set; } = new();
        public string MemoryDiagnosticStatus { get; set; } = "Unknown";
    }

    // ═══════════════════════════════════════════════════════════════
    //  4. DISK DIAGNOSTICS
    // ═══════════════════════════════════════════════════════════════

    public class DiskDiagnostics
    {
        public List<DriveInfoResult> Drives { get; set; } = new();
        public double DiskActivityPercent { get; set; }
        public bool DiskActivityFlagged { get; set; }
        public long WindowsTempMB { get; set; }
        public long UserTempMB { get; set; }
        public long SoftwareDistributionMB { get; set; }
        public bool WindowsOldExists { get; set; }
        public long WindowsOldMB { get; set; }
        public long UpgradeLogsMB { get; set; }
        public long RecycleBinMB { get; set; }
        public long PrefetchMB { get; set; }
    }

    public class DriveInfoResult
    {
        public string DriveLetter { get; set; } = "";
        public long TotalMB { get; set; }
        public long UsedMB { get; set; }
        public long FreeMB { get; set; }
        public double PercentUsed { get; set; }
        public bool UsageFlagged { get; set; }
        public string HealthStatus { get; set; } = "Unknown";
        public string DriveType { get; set; } = "Unknown";
    }

    // ═══════════════════════════════════════════════════════════════
    //  GPU DIAGNOSTICS
    // ═══════════════════════════════════════════════════════════════

    public class GpuInfo
    {
        public string Name { get; set; } = "Unknown";
        public string DriverVersion { get; set; } = "Unknown";
        public string DriverDate { get; set; } = "Unknown";
        public bool DriverOutdated { get; set; }
        public long DedicatedVideoMemoryMB { get; set; }
        public string Resolution { get; set; } = "Unknown";
        public string RefreshRate { get; set; } = "Unknown";
        public string AdapterStatus { get; set; } = "Unknown";
        public double TemperatureC { get; set; }
        public bool TemperatureFlagged { get; set; }
        public double UsagePercent { get; set; }
        public bool UsageFlagged { get; set; }
        public bool IsPrimary { get; set; }
    }

    public class GpuDiagnostics
    {
        public string GpuName { get; set; } = "Unknown";
        public string DriverVersion { get; set; } = "Unknown";
        public string DriverDate { get; set; } = "Unknown";
        public long DedicatedVideoMemoryMB { get; set; }
        public long SharedSystemMemoryMB { get; set; }
        public string Resolution { get; set; } = "Unknown";
        public string RefreshRate { get; set; } = "Unknown";
        public double GpuTemperatureC { get; set; }
        public bool TemperatureFlagged { get; set; }
        public double GpuUsagePercent { get; set; }
        public bool UsageFlagged { get; set; }
        public bool DriverOutdated { get; set; }
        public string AdapterStatus { get; set; } = "Unknown";
        public List<string> AdditionalAdapters { get; set; } = new();
        public List<GpuInfo> AllGpus { get; set; } = new();
        public List<DisplayInfo> Displays { get; set; } = new();
    }

    /// <summary>Represents a physical display/monitor detected via EnumDisplaySettings.</summary>
    public class DisplayInfo
    {
        public string DeviceName { get; set; } = "";
        public string MonitorName { get; set; } = "";
        public int Width { get; set; }
        public int Height { get; set; }
        public int RefreshRate { get; set; }
        public bool IsPrimary { get; set; }

        public string Resolution => Width > 0 && Height > 0 ? $"{Width} x {Height}" : "Unknown";
        public string RefreshRateText => RefreshRate > 0 ? $"{RefreshRate} Hz" : "Unknown";
    }

    // ═══════════════════════════════════════════════════════════════
    //  6. BATTERY HEALTH
    // ═══════════════════════════════════════════════════════════════

    public class BatteryDiagnostics
    {
        public bool HasBattery { get; set; }
        public long DesignCapacityMWh { get; set; }
        public long FullChargeCapacityMWh { get; set; }
        public double HealthPercent { get; set; }
        public bool HealthFlagged { get; set; }
        public string PowerSource { get; set; } = "Unknown";
        public string PowerPlan { get; set; } = "Unknown";
    }

    // ═══════════════════════════════════════════════════════════════
    //  7. STARTUP PROGRAMS
    // ═══════════════════════════════════════════════════════════════

    public class StartupDiagnostics
    {
        public List<StartupEntry> Entries { get; set; } = new();
        public int EnabledCount { get; set; }
        public bool TooManyFlagged { get; set; }
    }

    public class StartupEntry
    {
        public string Name { get; set; } = "";
        public string Publisher { get; set; } = "";
        public bool Enabled { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  8. VISUAL & DISPLAY SETTINGS
    // ═══════════════════════════════════════════════════════════════

    public class VisualSettingsDiagnostics
    {
        public string VisualEffectsSetting { get; set; } = "Unknown";
        public bool TransparencyEnabled { get; set; }
        public bool AnimationsEnabled { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  9. ANTIVIRUS & SECURITY
    // ═══════════════════════════════════════════════════════════════

    public class AntivirusDiagnostics
    {
        public string AntivirusName { get; set; } = "Unknown";
        public bool FullScanRunning { get; set; }
        public string BitLockerStatus { get; set; } = "Unknown";
    }

    // ═══════════════════════════════════════════════════════════════
    //  10. WINDOWS UPDATE
    // ═══════════════════════════════════════════════════════════════

    public class WindowsUpdateDiagnostics
    {
        public string PendingUpdates { get; set; } = "Unknown";
        public string UpdateServiceStatus { get; set; } = "Unknown";
        public string ServiceStartType { get; set; } = "Unknown";
        public long UpdateCacheMB { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  11. NETWORK DIAGNOSTICS
    // ═══════════════════════════════════════════════════════════════

    public class NetworkDiagnostics
    {
        public string ConnectionType { get; set; } = "Unknown";
        public string WifiBand { get; set; } = "N/A";
        public string WifiSignalStrength { get; set; } = "N/A";
        public string LinkSpeed { get; set; } = "Unknown";
        public string AdapterSpeed { get; set; } = "Unknown";
        public string DnsResponseTime { get; set; } = "Unknown";
        public bool VpnActive { get; set; }
        public string VpnClient { get; set; } = "N/A";
        public string PingLatency { get; set; } = "Unknown";
        public double DownloadSpeedMbps { get; set; }
        public double UploadSpeedMbps { get; set; }
        public bool SpeedTestFlagged { get; set; }
        public string SpeedTestError { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    //  11b. NETWORK DRIVE HEALTH
    // ═══════════════════════════════════════════════════════════════

    public class NetworkDriveDiagnostics
    {
        public List<NetworkDriveInfo> Drives { get; set; } = new();
    }

    public class NetworkDriveInfo
    {
        public string DriveLetter { get; set; } = "";
        public string UncPath { get; set; } = "";
        public string DriveLabel { get; set; } = "";
        public bool IsAccessible { get; set; }
        public long LatencyMs { get; set; }
        public long TotalMB { get; set; }
        public long FreeMB { get; set; }
        public int PercentUsed { get; set; }
        public bool OfflineFilesEnabled { get; set; }
        public string SyncStatus { get; set; } = "N/A";
        public bool LatencyFlagged { get; set; }
        public bool SpaceFlagged { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  12. OUTLOOK / EMAIL
    // ═══════════════════════════════════════════════════════════════

    public class OutlookDiagnostics
    {
        public bool OutlookInstalled { get; set; }
        public List<OutlookDataFile> DataFiles { get; set; } = new();
        public int AddInCount { get; set; }
    }

    public class OutlookDataFile
    {
        public string Path { get; set; } = "";
        public long SizeMB { get; set; }
        public bool SizeFlagged { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  13. BROWSER CHECK
    // ═══════════════════════════════════════════════════════════════

    public class BrowserDiagnostics
    {
        public List<BrowserInfo> Browsers { get; set; } = new();
    }

    public class BrowserInfo
    {
        public string Name { get; set; } = "";
        public int OpenTabs { get; set; }
        public long CacheSizeMB { get; set; }
        public int ExtensionCount { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  14. USER PROFILE
    // ═══════════════════════════════════════════════════════════════

    public class UserProfileDiagnostics
    {
        public long ProfileSizeMB { get; set; }
        public string ProfileAge { get; set; } = "Unknown";
        public int DesktopItemCount { get; set; }
        public bool DesktopItemsFlagged { get; set; }
        public bool CorruptionDetected { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  15. GENERAL OFFICE
    // ═══════════════════════════════════════════════════════════════

    public class OfficeDiagnostics
    {
        public bool OfficeInstalled { get; set; }
        public string OfficeVersion { get; set; } = "Unknown";
        public bool RepairNeeded { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  16. INSTALLED SOFTWARE
    // ═══════════════════════════════════════════════════════════════

    public class InstalledSoftwareDiagnostics
    {
        public List<InstalledApp> Applications { get; set; } = new();
        public int TotalCount { get; set; }
        public List<InstalledApp> EOLApps { get; set; } = new();
        public List<InstalledApp> BloatwareApps { get; set; } = new();
        public List<InstalledApp> RuntimeApps { get; set; } = new();
        public int RuntimeCount { get; set; }
    }

    public class InstalledApp
    {
        public string Name { get; set; } = "";
        public string Version { get; set; } = "";
        public string Publisher { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    //  EVENT LOG DIAGNOSTICS
    // ═══════════════════════════════════════════════════════════════

    public class EventLogDiagnostics
    {
        public List<EventLogEntry> BSODs { get; set; } = new();
        public List<EventLogEntry> AppCrashes { get; set; } = new();
        public List<EventLogEntry> DiskErrors { get; set; } = new();
        public List<EventLogEntry> UnexpectedShutdowns { get; set; } = new();
        public int TotalEventCount => BSODs.Count + AppCrashes.Count + DiskErrors.Count + UnexpectedShutdowns.Count;

        /// <summary>
        /// Extracts faulting application names from AppCrash event messages,
        /// grouped by count descending. Event ID 1000 messages start with
        /// "Faulting application name: foo.exe, ..." — we parse that prefix.
        /// </summary>
        public List<(string AppName, int Count)> GetTopCrashingApps(int max = 5)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in AppCrashes)
            {
                string app = ExtractFaultingApp(entry.Message);
                if (!string.IsNullOrEmpty(app))
                {
                    counts.TryGetValue(app, out int c);
                    counts[app] = c + 1;
                }
            }
            return counts
                .OrderByDescending(kv => kv.Value)
                .Take(max)
                .Select(kv => (kv.Key, kv.Value))
                .ToList();
        }

        private static string ExtractFaultingApp(string message)
        {
            if (string.IsNullOrEmpty(message)) return null;

            // "Faulting application name: outlook.exe, version: ..."
            const string marker = "application name:";
            int idx = message.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                int start = idx + marker.Length;
                int comma = message.IndexOf(',', start);
                if (comma < 0) comma = message.Length;
                string name = message[start..comma].Trim();
                if (!string.IsNullOrEmpty(name)) return name;
            }

            // Fallback for "Application Hang" (EventID 1002): "The program X stopped interacting..."
            const string hangMarker = "program ";
            idx = message.IndexOf(hangMarker, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                int start = idx + hangMarker.Length;
                int space = message.IndexOf(' ', start);
                if (space < 0) space = message.Length;
                string name = message[start..space].Trim();
                if (!string.IsNullOrEmpty(name)) return name;
            }

            return null;
        }
    }

    public class EventLogEntry
    {
        public string Source { get; set; } = "";
        public long EventId { get; set; }
        public string Level { get; set; } = "";
        public DateTime Timestamp { get; set; }
        public string Message { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    //  BEFORE/AFTER SNAPSHOT (for post-optimization comparison)
    // ═══════════════════════════════════════════════════════════════

    public record BeforeAfterSnapshot(
        int HealthScore,
        long DiskFreeMB,
        double RamUsedPercent,
        int StartupCount,
        int IssueCount);

    // ═══════════════════════════════════════════════════════════════
    //  SHARED DATA MODELS (kept for compatibility)
    // ═══════════════════════════════════════════════════════════════

    public class ProcessInfo
    {
        public string Name { get; set; } = "";
        public int Id { get; set; }
        public long MemoryMB { get; set; }
    }
}
