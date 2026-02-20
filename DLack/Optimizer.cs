using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace DLack
{
    public class Optimizer
    {
        public event Action<string> OnLog;

        private const int ProcessTimeoutMs = 10_000;

        // ═══════════════════════════════════════════════════════════════
        //  PUBLIC API
        // ═══════════════════════════════════════════════════════════════

        public OptimizationResult RunOptimizations(ScanResult beforeScan)
        {
            if (beforeScan == null)
                throw new ArgumentNullException(nameof(beforeScan));

            var result = new OptimizationResult
            {
                BeforeScan = beforeScan,
                Timestamp = DateTime.Now
            };

            OnLog?.Invoke("=== Starting Optimizations ===");

            // 1. High Performance power plan
            if (!string.Equals(beforeScan.PowerPlan, "High performance", StringComparison.OrdinalIgnoreCase))
            {
                result.PowerPlanChanged = SetHighPerformancePowerPlan();
            }
            else
            {
                OnLog?.Invoke("⊘ Power plan already set to High Performance — skipping");
            }

            // 2. Performance Visual Mode
            result.VisualModeApplied = ApplyPerformanceVisualMode();

            // 3. Clear temp files
            result.TempFilesClearedMB = ClearTempFiles();

            // 4. Empty recycle bin
            result.RecycleBinEmptied = EmptyRecycleBin();

            // 5. Restart Windows Explorer
            result.ExplorerRestarted = RestartExplorer();

            // 6. Flush DNS
            result.DnsFlushed = FlushDns();

            int successCount = (result.PowerPlanChanged ? 1 : 0)
                             + (result.VisualModeApplied ? 1 : 0)
                             + (result.TempFilesClearedMB > 0 ? 1 : 0)
                             + (result.RecycleBinEmptied ? 1 : 0)
                             + (result.ExplorerRestarted ? 1 : 0)
                             + (result.DnsFlushed ? 1 : 0);

            OnLog?.Invoke($"=== Optimizations Complete ({successCount}/6 succeeded) ===");

            return result;
        }

        // ═══════════════════════════════════════════════════════════════
        //  OPTIMIZATION ACTIONS
        // ═══════════════════════════════════════════════════════════════

        private bool SetHighPerformancePowerPlan()
        {
            try
            {
                OnLog?.Invoke("Setting High Performance power plan...");

                const string highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c";
                bool ok = RunProcess("powercfg.exe", $"/setactive {highPerfGuid}");

                if (ok)
                    OnLog?.Invoke("✓ Power plan set to High Performance");
                else
                    OnLog?.Invoke("✗ Power plan change may have failed");

                return ok;
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"✗ Failed to set power plan: {ex.Message}");
                return false;
            }
        }

        private bool ApplyPerformanceVisualMode()
        {
            try
            {
                OnLog?.Invoke("Applying Performance Visual Mode...");
                bool applied = true;

                // Set to "Custom" visual effects
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"))
                {
                    if (key == null)
                    {
                        OnLog?.Invoke("⚠ Could not open VisualEffects registry key");
                        applied = false;
                    }
                    else
                    {
                        key.SetValue("VisualFXSetting", 2, RegistryValueKind.DWord);
                    }
                }

                // Disable animations but keep font smoothing enabled
                using (var key = Registry.CurrentUser.CreateSubKey(@"Control Panel\Desktop"))
                {
                    if (key == null)
                    {
                        OnLog?.Invoke("⚠ Could not open Desktop registry key");
                        applied = false;
                    }
                    else
                    {
                        key.SetValue("UserPreferencesMask",
                            new byte[] { 0x90, 0x12, 0x03, 0x80, 0x10, 0x00, 0x00, 0x00 },
                            RegistryValueKind.Binary);
                    }
                }

                // Disable minimize animation
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Control Panel\Desktop\WindowMetrics"))
                {
                    if (key == null)
                    {
                        OnLog?.Invoke("⚠ Could not open WindowMetrics registry key");
                        applied = false;
                    }
                    else
                    {
                        key.SetValue("MinAnimate", "0", RegistryValueKind.String);
                    }
                }

                if (applied)
                    OnLog?.Invoke("✓ Performance Visual Mode applied");
                else
                    OnLog?.Invoke("⚠ Performance Visual Mode partially applied");

                return applied;
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"✗ Failed to apply visual mode: {ex.Message}");
                return false;
            }
        }

        private long ClearTempFiles()
        {
            long totalBytes = 0;
            int fileCount = 0;

            try
            {
                OnLog?.Invoke("Clearing temp files...");

                // User temp
                var (bytes1, count1) = ClearFolder(Path.GetTempPath());
                totalBytes += bytes1;
                fileCount += count1;

                // Windows temp (requires admin)
                string winTemp = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp");
                var (bytes2, count2) = ClearFolder(winTemp);
                totalBytes += bytes2;
                fileCount += count2;

                long totalMB = totalBytes / (1024 * 1024);
                OnLog?.Invoke($"✓ Cleared {fileCount} files ({FormatBytes(totalBytes)})");

                return totalMB;
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"✗ Temp cleanup error: {ex.Message}");
                return totalBytes / (1024 * 1024);
            }
        }

        private (long bytesFreed, int filesDeleted) ClearFolder(string path)
        {
            long bytesFreed = 0;
            int filesDeleted = 0;

            if (!Directory.Exists(path)) return (0, 0);

            try
            {
                // Safe enumeration — skips folders we can't access
                var files = Directory.EnumerateFiles(path, "*", new EnumerationOptions
                {
                    IgnoreInaccessible = true,
                    RecurseSubdirectories = true
                });

                foreach (var file in files)
                {
                    try
                    {
                        var info = new FileInfo(file);
                        long size = info.Length;    // Capture size BEFORE delete
                        info.Delete();
                        bytesFreed += size;
                        filesDeleted++;
                    }
                    catch { } // Skip locked/in-use files
                }

                // Clean empty subdirectories
                try
                {
                    foreach (var dir in Directory.EnumerateDirectories(path, "*", new EnumerationOptions
                    {
                        IgnoreInaccessible = true,
                        RecurseSubdirectories = true
                    }))
                    {
                        try
                        {
                            if (Directory.GetFileSystemEntries(dir).Length == 0)
                                Directory.Delete(dir);
                        }
                        catch { }
                    }
                }
                catch { }
            }
            catch { }

            return (bytesFreed, filesDeleted);
        }

        private bool EmptyRecycleBin()
        {
            try
            {
                OnLog?.Invoke("Emptying Recycle Bin...");

                const uint SHERB_NOCONFIRMATION = 0x0001;
                const uint SHERB_NOPROGRESSUI = 0x0002;
                const uint SHERB_NOSOUND = 0x0004;

                int hr = SHEmptyRecycleBin(IntPtr.Zero, null,
                    SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);

                // S_OK = 0, or 0x80070012 = bin already empty (not a real error)
                if (hr == 0 || hr == unchecked((int)0x80070012))
                {
                    OnLog?.Invoke("✓ Recycle Bin emptied");
                    return true;
                }

                OnLog?.Invoke($"⚠ Recycle Bin returned HRESULT 0x{hr:X8}");
                return false;
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"✗ Failed to empty recycle bin: {ex.Message}");
                return false;
            }
        }

        private bool RestartExplorer()
        {
            try
            {
                OnLog?.Invoke("Restarting Windows Explorer...");

                var explorers = Process.GetProcessesByName("explorer");

                if (explorers.Length == 0)
                {
                    OnLog?.Invoke("⊘ No Explorer process found — starting fresh");
                    Process.Start("explorer.exe");
                    return true;
                }

                foreach (var proc in explorers)
                {
                    try
                    {
                        proc.Kill();
                        proc.WaitForExit(5000);
                    }
                    catch { }
                    finally
                    {
                        proc.Dispose();
                    }
                }

                // Brief pause then restart
                System.Threading.Thread.Sleep(1500);
                Process.Start("explorer.exe");

                OnLog?.Invoke("✓ Explorer restarted");
                return true;
            }
            catch (Exception ex)
            {
                // Try to restart explorer even if kill failed
                try { Process.Start("explorer.exe"); } catch { }
                OnLog?.Invoke($"✗ Failed to restart Explorer: {ex.Message}");
                return false;
            }
        }

        private bool FlushDns()
        {
            try
            {
                OnLog?.Invoke("Flushing DNS cache...");

                bool ok = RunProcess("ipconfig.exe", "/flushdns");

                if (ok)
                    OnLog?.Invoke("✓ DNS cache flushed");
                else
                    OnLog?.Invoke("✗ DNS flush may have failed");

                return ok;
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"✗ Failed to flush DNS: {ex.Message}");
                return false;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  UTILITY
        // ═══════════════════════════════════════════════════════════════

        /// <summary>Runs a process silently and returns true if exit code is 0.</summary>
        private bool RunProcess(string fileName, string arguments)
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null) return false;

            bool exited = process.WaitForExit(ProcessTimeoutMs);
            if (!exited)
            {
                try { process.Kill(entireProcessTree: true); }
                catch { }
                return false;
            }

            return exited && process.ExitCode == 0;
        }

        private static string FormatBytes(long bytes)
        {
            if (bytes >= 1024 * 1024 * 1024)
                return $"{bytes / (1024.0 * 1024 * 1024):F1} GB";
            if (bytes >= 1024 * 1024)
                return $"{bytes / (1024.0 * 1024):F1} MB";
            if (bytes >= 1024)
                return $"{bytes / 1024.0:F0} KB";
            return $"{bytes} bytes";
        }

        // ═══════════════════════════════════════════════════════════════
        //  P/INVOKE
        // ═══════════════════════════════════════════════════════════════

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);
    }

    // ═══════════════════════════════════════════════════════════════
    //  RESULT MODEL
    // ═══════════════════════════════════════════════════════════════

    public class OptimizationResult
    {
        public DateTime Timestamp { get; set; }
        public ScanResult BeforeScan { get; set; }
        public ScanResult AfterScan { get; set; }

        public bool PowerPlanChanged { get; set; }
        public bool VisualModeApplied { get; set; }
        public long TempFilesClearedMB { get; set; }
        public bool RecycleBinEmptied { get; set; }
        public bool ExplorerRestarted { get; set; }
        public bool DnsFlushed { get; set; }

        public double CpuImprovement => BeforeScan != null && AfterScan != null
            ? Math.Round(BeforeScan.AvgCpu - AfterScan.AvgCpu, 1)
            : 0;

        public long RamFreedMB => BeforeScan != null && AfterScan != null
            ? BeforeScan.RamUsedMB - AfterScan.RamUsedMB
            : 0;

        public double DiskFreedPercent => BeforeScan != null && AfterScan != null
            ? Math.Round(AfterScan.DiskFreePercent - BeforeScan.DiskFreePercent, 1)
            : 0;
    }
}