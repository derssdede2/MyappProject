using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace DLack
{
    /// <summary>
    /// Lightweight snapshot of key metrics from a single scan, stored for trend tracking.
    /// </summary>
    public record ScanSnapshot(
        DateTime Timestamp,
        int HealthScore,
        double RamUsedPercent,
        long DiskFreeMB,
        int CriticalCount,
        int WarningCount,
        int CrashCount,
        int StartupEnabled,
        double CpuLoadPercent);

    /// <summary>
    /// Persists scan snapshots to disk and provides trend data.
    /// Stores up to <see cref="MaxEntries"/> snapshots in a JSON file
    /// at <c>%LOCALAPPDATA%\DLack\scan-history.json</c>.
    /// </summary>
    internal static class ScanHistoryManager
    {
        private const int MaxEntries = 50;

        private static readonly string HistoryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DLack", "scan-history.json");

        /// <summary>Creates a snapshot from a completed scan result.</summary>
        public static ScanSnapshot CreateSnapshot(DiagnosticResult result)
        {
            ArgumentNullException.ThrowIfNull(result);
            return new ScanSnapshot(
                Timestamp: result.Timestamp,
                HealthScore: result.HealthScore,
                RamUsedPercent: result.Ram.PercentUsed,
                DiskFreeMB: result.Disk.Drives.FirstOrDefault()?.FreeMB ?? 0,
                CriticalCount: result.FlaggedIssues.Count(i => i.Severity == Severity.Critical),
                WarningCount: result.FlaggedIssues.Count(i => i.Severity == Severity.Warning),
                CrashCount: result.EventLog.AppCrashes.Count + result.EventLog.BSODs.Count,
                StartupEnabled: result.Startup.EnabledCount,
                CpuLoadPercent: result.Cpu.CpuLoadPercent);
        }

        /// <summary>Saves a snapshot, appending to the persisted history.</summary>
        public static void Save(ScanSnapshot snapshot)
        {
            ArgumentNullException.ThrowIfNull(snapshot);
            var history = Load();
            history.Add(snapshot);

            // Trim oldest entries beyond the cap
            if (history.Count > MaxEntries)
                history.RemoveRange(0, history.Count - MaxEntries);

            try
            {
                string dir = Path.GetDirectoryName(HistoryPath)!;
                Directory.CreateDirectory(dir);

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(history, options);
                File.WriteAllText(HistoryPath, json);
            }
            catch { /* best-effort persistence */ }
        }

        /// <summary>Loads all saved snapshots, ordered oldest to newest.</summary>
        public static List<ScanSnapshot> Load()
        {
            try
            {
                if (!File.Exists(HistoryPath))
                    return new List<ScanSnapshot>();

                string json = File.ReadAllText(HistoryPath);
                return JsonSerializer.Deserialize<List<ScanSnapshot>>(json)
                    ?? new List<ScanSnapshot>();
            }
            catch
            {
                return new List<ScanSnapshot>();
            }
        }

        /// <summary>
        /// Returns the most recent snapshot before the current scan, or null if no history exists.
        /// </summary>
        public static ScanSnapshot GetPrevious()
        {
            var history = Load();
            return history.Count >= 2 ? history[^2] : history.FirstOrDefault();
        }

        /// <summary>
        /// Compares two snapshots and returns human-readable delta descriptions.
        /// </summary>
        public static List<(string Metric, string Before, string After, int Direction)> CompareTo(
            ScanSnapshot previous, ScanSnapshot current)
        {
            ArgumentNullException.ThrowIfNull(previous);
            ArgumentNullException.ThrowIfNull(current);

            var deltas = new List<(string, string, string, int)>();

            AddDelta(deltas, "Health Score",
                $"{previous.HealthScore}/100", $"{current.HealthScore}/100",
                current.HealthScore - previous.HealthScore);

            AddDelta(deltas, "Disk Free",
                DiagnosticFormatters.FormatMB(previous.DiskFreeMB),
                DiagnosticFormatters.FormatMB(current.DiskFreeMB),
                Math.Sign(current.DiskFreeMB - previous.DiskFreeMB));

            AddDelta(deltas, "RAM Usage",
                $"{previous.RamUsedPercent:F0}%", $"{current.RamUsedPercent:F0}%",
                -Math.Sign((int)(current.RamUsedPercent - previous.RamUsedPercent))); // lower is better

            AddDelta(deltas, "Critical Issues",
                $"{previous.CriticalCount}", $"{current.CriticalCount}",
                -Math.Sign(current.CriticalCount - previous.CriticalCount)); // fewer is better

            AddDelta(deltas, "Crashes",
                $"{previous.CrashCount}", $"{current.CrashCount}",
                -Math.Sign(current.CrashCount - previous.CrashCount));

            AddDelta(deltas, "Startup Items",
                $"{previous.StartupEnabled}", $"{current.StartupEnabled}",
                -Math.Sign(current.StartupEnabled - previous.StartupEnabled));

            return deltas;
        }

        private static void AddDelta(
            List<(string, string, string, int)> list,
            string metric, string before, string after, int direction)
        {
            // direction: +1 = improved, -1 = degraded, 0 = unchanged
            list.Add((metric, before, after, Math.Sign(direction)));
        }
    }
}
