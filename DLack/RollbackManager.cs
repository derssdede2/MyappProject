using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Microsoft.Win32;

namespace DLack
{
    /// <summary>
    /// Captures registry values before optimization and supports restoring them.
    /// Rollback data is persisted at <c>%LOCALAPPDATA%\DLack\rollback.json</c>.
    /// </summary>
    internal sealed class RollbackManager
    {
        private static readonly string RollbackPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DLack", "rollback.json");

        private readonly List<RollbackEntry> _entries = new();

        /// <summary>Whether any rollback entries have been captured.</summary>
        public bool HasEntries => _entries.Count > 0;

        /// <summary>
        /// Reads the current value of a registry key/value and stores it for potential rollback,
        /// then returns the current value so the caller can decide whether to proceed.
        /// </summary>
        public object CaptureRegistryValue(RegistryKey root, string subKeyPath, string valueName)
        {
            try
            {
                using var key = root.OpenSubKey(subKeyPath);
                object currentValue = key?.GetValue(valueName);
                var kind = key?.GetValueKind(valueName) ?? RegistryValueKind.None;

                _entries.Add(new RollbackEntry
                {
                    RootKey = root.Name,
                    SubKeyPath = subKeyPath,
                    ValueName = valueName,
                    OriginalValueBase64 = SerializeValue(currentValue, kind),
                    ValueKind = kind.ToString(),
                    CapturedAt = DateTime.UtcNow,
                    ActionKey = ""
                });

                return currentValue;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>Tags the most recently captured entry with the action that will modify it.</summary>
        public void TagLastEntry(string actionKey)
        {
            if (_entries.Count > 0)
                _entries[^1].ActionKey = actionKey;
        }

        /// <summary>Persists all captured rollback entries to disk.</summary>
        public void Save()
        {
            if (_entries.Count == 0) return;
            try
            {
                // Merge with any existing rollback data
                var existing = LoadEntries();
                existing.AddRange(_entries);

                string dir = Path.GetDirectoryName(RollbackPath)!;
                Directory.CreateDirectory(dir);

                var options = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(RollbackPath, JsonSerializer.Serialize(existing, options));
            }
            catch { /* best-effort */ }
        }

        /// <summary>
        /// Restores all registry values from the most recent rollback session.
        /// Returns a list of (actionKey, success, message) results.
        /// </summary>
        public static List<(string ActionKey, bool Success, string Message)> RestoreAll()
        {
            var results = new List<(string, bool, string)>();
            var entries = LoadEntries();
            if (entries.Count == 0)
            {
                results.Add(("", false, "No rollback data found"));
                return results;
            }

            foreach (var entry in entries)
            {
                try
                {
                    var root = ResolveRoot(entry.RootKey);
                    if (root == null)
                    {
                        results.Add((entry.ActionKey, false, $"Unknown root key: {entry.RootKey}"));
                        continue;
                    }

                    if (string.IsNullOrEmpty(entry.OriginalValueBase64))
                    {
                        // Value didn't exist before — delete it
                        using var key = root.OpenSubKey(entry.SubKeyPath, writable: true);
                        key?.DeleteValue(entry.ValueName, throwOnMissingValue: false);
                        results.Add((entry.ActionKey, true, $"Removed {entry.ValueName}"));
                    }
                    else
                    {
                        // Restore original value
                        using var key = root.OpenSubKey(entry.SubKeyPath, writable: true);
                        if (key == null)
                        {
                            results.Add((entry.ActionKey, false, $"Key not found: {entry.SubKeyPath}"));
                            continue;
                        }

                        var kind = Enum.Parse<RegistryValueKind>(entry.ValueKind);
                        object value = DeserializeValue(entry.OriginalValueBase64, kind);
                        key.SetValue(entry.ValueName, value, kind);
                        results.Add((entry.ActionKey, true, $"Restored {entry.ValueName}"));
                    }
                }
                catch (Exception ex)
                {
                    results.Add((entry.ActionKey, false, $"Failed: {ex.Message}"));
                }
            }

            // Clear rollback data after restore
            try { File.Delete(RollbackPath); } catch { }

            return results;
        }

        /// <summary>Whether rollback data exists on disk from a previous optimization.</summary>
        public static bool HasSavedRollback()
        {
            try { return File.Exists(RollbackPath) && new FileInfo(RollbackPath).Length > 2; }
            catch { return false; }
        }

        /// <summary>Clears persisted rollback data without restoring.</summary>
        public static void ClearSavedRollback()
        {
            try { File.Delete(RollbackPath); } catch { }
        }

        // ── Serialization helpers ──────────────────────────────────

        private static string SerializeValue(object value, RegistryValueKind kind)
        {
            if (value == null) return null;
            return kind switch
            {
                RegistryValueKind.Binary => Convert.ToBase64String((byte[])value),
                RegistryValueKind.DWord => ((int)value).ToString(),
                RegistryValueKind.QWord => ((long)value).ToString(),
                RegistryValueKind.String or RegistryValueKind.ExpandString => (string)value,
                _ => value.ToString()
            };
        }

        private static object DeserializeValue(string serialized, RegistryValueKind kind)
        {
            return kind switch
            {
                RegistryValueKind.Binary => Convert.FromBase64String(serialized),
                RegistryValueKind.DWord => int.Parse(serialized),
                RegistryValueKind.QWord => long.Parse(serialized),
                RegistryValueKind.String or RegistryValueKind.ExpandString => serialized,
                _ => serialized
            };
        }

        private static RegistryKey ResolveRoot(string rootName)
        {
            if (rootName.Contains("CURRENT_USER", StringComparison.OrdinalIgnoreCase))
                return Registry.CurrentUser;
            if (rootName.Contains("LOCAL_MACHINE", StringComparison.OrdinalIgnoreCase))
                return Registry.LocalMachine;
            return null;
        }

        private static List<RollbackEntry> LoadEntries()
        {
            try
            {
                if (!File.Exists(RollbackPath)) return new();
                string json = File.ReadAllText(RollbackPath);
                return JsonSerializer.Deserialize<List<RollbackEntry>>(json) ?? new();
            }
            catch { return new(); }
        }

        private class RollbackEntry
        {
            public string RootKey { get; set; } = "";
            public string SubKeyPath { get; set; } = "";
            public string ValueName { get; set; } = "";
            public string OriginalValueBase64 { get; set; }
            public string ValueKind { get; set; } = "";
            public DateTime CapturedAt { get; set; }
            public string ActionKey { get; set; } = "";
        }
    }
}
