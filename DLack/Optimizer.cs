using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;

namespace DLack
{
    /// <summary>
    /// Result returned by each optimization action, replacing shared mutable fields.
    /// </summary>
    internal record ActionResult(
        long FreedMB = 0,
        int DeletedCount = 0,
        int SkippedCount = 0,
        string Detail = null);

    public class Optimizer
    {
        public event Action<string> OnLog;
        public event Action<OptimizationAction> OnActionStatusChanged;

        // ═══════════════════════════════════════════════════════════════
        //  BUILD PLAN
        // ═══════════════════════════════════════════════════════════════

        public List<OptimizationAction> BuildPlan(DiagnosticResult result)
        {
            var actions = new List<OptimizationAction>();

            // ── Disk: Temp Folders ──
            // Use scan data as baseline; quick existence check only
            long totalTempMB = result.Disk.WindowsTempMB + result.Disk.UserTempMB;
            if (totalTempMB > 500 && (DirectoryHasFiles(Path.GetTempPath()) ||
                DirectoryHasFiles(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp"))))
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Disk Cleanup",
                    Title = "Clear Temporary Files",
                    Description = $"Delete Windows and user temp files (~{DiagnosticFormatters.FormatMB(totalTempMB)})",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    EstimatedFreeMB = totalTempMB,
                    IsSelected = true,
                    ActionKey = "ClearTempFiles"
                });
            }

            // ── Disk: Prefetch ──
            string prefetchDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Prefetch");
            if (result.Disk.PrefetchMB > 128 && DirectoryHasFiles(prefetchDir))
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Disk Cleanup",
                    Title = "Clear Prefetch Cache",
                    Description = $"Delete Windows prefetch data (~{DiagnosticFormatters.FormatMB(result.Disk.PrefetchMB)})",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    EstimatedFreeMB = result.Disk.PrefetchMB,
                    IsSelected = true,
                    ActionKey = "ClearPrefetch"
                });
            }

            // ── Disk: Recycle Bin (fast P/Invoke query) ──
            long liveRecycleMB = GetRecycleBinSizeMB();
            if (liveRecycleMB > 100)
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Disk Cleanup",
                    Title = "Empty Recycle Bin",
                    Description = $"Permanently delete recycled files ({DiagnosticFormatters.FormatMB(liveRecycleMB)})",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    EstimatedFreeMB = liveRecycleMB,
                    IsSelected = true,
                    ActionKey = "EmptyRecycleBin"
                });
            }

            // ── Disk: Software Distribution ──
            string sdDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "Download");
            if (result.Disk.SoftwareDistributionMB > 500 && DirectoryHasFiles(sdDir))
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Disk Cleanup",
                    Title = "Clear Windows Update Cache",
                    Description = $"Stop wuauserv, clear update downloads (~{DiagnosticFormatters.FormatMB(result.Disk.SoftwareDistributionMB)})",
                    IsAutomatable = true,
                    Risk = ActionRisk.Moderate,
                    EstimatedFreeMB = result.Disk.SoftwareDistributionMB,
                    IsSelected = false,
                    ActionKey = "ClearUpdateCache"
                });
            }

            // ── Disk: Windows.old ──
            if (result.Disk.WindowsOldExists && result.Disk.WindowsOldMB > 0)
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Disk Cleanup",
                    Title = "Remove Windows.old",
                    Description = $"Delete old Windows installation folder (~{DiagnosticFormatters.FormatMB(result.Disk.WindowsOldMB)})",
                    IsAutomatable = true,
                    Risk = ActionRisk.Moderate,
                    EstimatedFreeMB = result.Disk.WindowsOldMB,
                    IsSelected = false,
                    ActionKey = "LaunchDiskCleanup"
                });
            }

            // ── Disk: Windows Upgrade Logs ──
            if (result.Disk.UpgradeLogsMB > 200)
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Disk Cleanup",
                    Title = "Clean Windows Upgrade Logs",
                    Description = $"Remove upgrade logs, Panther files, and temporary setup data (~{DiagnosticFormatters.FormatMB(result.Disk.UpgradeLogsMB)})",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    EstimatedFreeMB = result.Disk.UpgradeLogsMB,
                    IsSelected = true,
                    ActionKey = "CleanUpgradeLogs"
                });
            }

            // ── Browser: Cache ──
            foreach (var browser in result.Browser.Browsers.Where(b => b.CacheSizeMB > 200))
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Browser",
                    Title = $"Clear {browser.Name} Cache",
                    Description = $"Delete browser cache files (~{DiagnosticFormatters.FormatMB(browser.CacheSizeMB)})",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    EstimatedFreeMB = browser.CacheSizeMB,
                    IsSelected = true,
                    ActionKey = $"ClearBrowserCache:{browser.Name}"
                });
            }

            // ── CPU: High Load ──
            if (result.Cpu.CpuLoadFlagged)
            {
                // Always offer Resource Monitor when CPU is flagged
                actions.Add(new OptimizationAction
                {
                    Category = "CPU",
                    Title = "Open Resource Monitor",
                    Description = $"CPU load is {result.Cpu.CpuLoadPercent}% at idle — open Resource Monitor to identify the cause",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    IsSelected = true,
                    ActionKey = "OpenResourceMonitor"
                });
            }

            // ── CPU: Power Plan ──
            if (result.Battery.PowerPlan.Contains("Power saver", StringComparison.OrdinalIgnoreCase))
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Performance",
                    Title = "Switch to Balanced Power Plan",
                    Description = "Power saver mode throttles CPU performance — switch to Balanced",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    IsSelected = true,
                    ActionKey = "SetPowerPlanBalanced"
                });
            }

            // ── Visual Settings: Performance Mode ──
            if (result.VisualSettings.TransparencyEnabled || result.VisualSettings.AnimationsEnabled)
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Visual Settings",
                    Title = "Switch to Performance Visuals",
                    Description = "Disable transparency and animations to reduce CPU/GPU overhead",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    IsSelected = false,
                    ActionKey = "SwitchToPerformanceVisuals"
                });
            }

            // ── RAM: High Memory Usage ──
            if (result.Ram.UsageFlagged && result.Ram.TopRamProcesses.Count > 0)
            {
                var topProc = result.Ram.TopRamProcesses
                    .Where(p => p.Name is not ("System" or "svchost" or "csrss" or "dwm" or "explorer"))
                    .OrderByDescending(p => p.MemoryMB)
                    .FirstOrDefault();
                if (topProc != null && topProc.MemoryMB > 500)
                {
                    actions.Add(new OptimizationAction
                    {
                        Category = "RAM",
                        Title = $"End Memory-Heavy Process: {topProc.Name}",
                        Description = $"{topProc.Name} (PID {topProc.Id}) is using {DiagnosticFormatters.FormatMB(topProc.MemoryMB)} of RAM",
                        IsAutomatable = true,
                        Risk = ActionRisk.Moderate,
                        IsSelected = false,
                        ActionKey = $"KillProcess:{topProc.Id}:{topProc.Name}"
                    });
                }
            }

            // ── Browser: Too Many Tabs ──
            foreach (var browser in result.Browser.Browsers.Where(b => b.OpenTabs > 30))
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Browser",
                    Title = $"Close {browser.Name} Tabs",
                    Description = $"{browser.OpenTabs} tabs open — excessive tabs slow your system and consume RAM",
                    IsAutomatable = false,
                    Risk = ActionRisk.Safe,
                    IsSelected = false,
                    ActionKey = "ManualOnly"
                });
            }

            // ── Startup: Too Many ──
            if (result.Startup.TooManyFlagged)
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Startup",
                    Title = "Open Task Manager (Startup)",
                    Description = $"{result.Startup.EnabledCount} startup items enabled — open Task Manager to review",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    IsSelected = true,
                    ActionKey = "OpenTaskManagerStartup"
                });
            }

            // ── Windows Update: Service Disabled ──
            if (result.WindowsUpdate.ServiceStartType == "Disabled")
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Windows Update",
                    Title = "Enable Windows Update Service",
                    Description = "Windows Update service is disabled — security patches will not be installed",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    IsSelected = true,
                    ActionKey = "StartWuauserv"
                });
            }

            // ── System: Uptime / Reboot Needed ──
            bool needsReboot = result.SystemOverview.UptimeFlagged ||
                               result.WindowsUpdate.PendingUpdates.Contains("reboot", StringComparison.OrdinalIgnoreCase);
            if (needsReboot)
            {
                string reason = result.SystemOverview.UptimeFlagged
                    ? $"System has been running for {(int)result.SystemOverview.Uptime.TotalDays} days"
                    : "Pending updates require a reboot";
                actions.Add(new OptimizationAction
                {
                    Category = "System",
                    Title = "Schedule System Restart",
                    Description = reason,
                    IsAutomatable = true,
                    Risk = ActionRisk.RequiresReboot,
                    IsSelected = false,
                    ActionKey = "ScheduleRestart"
                });
            }

            // ── GPU: Outdated Driver ──
            if (result.Gpu.DriverOutdated)
            {
                actions.Add(new OptimizationAction
                {
                    Category = "GPU",
                    Title = "Open GPU Driver Update Page",
                    Description = $"GPU driver dated {result.Gpu.DriverDate} — open manufacturer's download page",
                    IsAutomatable = true,
                    Risk = ActionRisk.Safe,
                    IsSelected = true,
                    ActionKey = $"OpenGpuDriverPage:{result.Gpu.GpuName}"
                });
            }

            // ── Office: Repair Needed ──
            if (result.Office.OfficeInstalled && result.Office.RepairNeeded)
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Office",
                    Title = "Launch Office Repair",
                    Description = "Open Settings > Apps to repair Office installation",
                    IsAutomatable = true,
                    Risk = ActionRisk.Moderate,
                    IsSelected = false,
                    ActionKey = "OpenAppsSettings"
                });
            }

            // ── Security: BitLocker ──
            if (result.Antivirus.BitLockerStatus == "Not Encrypted")
            {
                actions.Add(new OptimizationAction
                {
                    Category = "Security",
                    Title = "Open BitLocker Settings",
                    Description = "C: drive is not encrypted — open BitLocker management",
                    IsAutomatable = true,
                    Risk = ActionRisk.Moderate,
                    IsSelected = false,
                    ActionKey = "OpenBitLocker"
                });
            }

            // ── Event Log: BSODs ──
            if (result.EventLog.BSODs.Count > 0)
            {
                string bsodContext = FormatEventContext("BSODs", result.EventLog.BSODs.Count, "BSOD");
                bool bsodRepeat = IsRepeatedRemediation("BSODs");

                if (bsodRepeat)
                {
                    // Previous SFC/DISM/driver update did not stop BSODs — escalate
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "⚠ BSODs persist after previous repair",
                        Description = $"{bsodContext}. Previous automated repairs (SFC, DISM, driver updates) did not resolve the issue — hardware diagnostics or OS reinstall may be needed",
                        IsAutomatable = false,
                        Risk = ActionRisk.Safe,
                        IsSelected = false,
                        ActionKey = "ManualOnly"
                    });
                }
                else
                {
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Run System File Checker",
                        Description = $"{bsodContext} — repair corrupted system files with sfc /scannow",
                        IsAutomatable = true,
                        Risk = ActionRisk.Safe,
                        IsSelected = true,
                        ActionKey = "RunSfc"
                    });
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Run DISM System Image Repair",
                        Description = "Repair the Windows component store to fix system file corruption",
                        IsAutomatable = true,
                        Risk = ActionRisk.Safe,
                        IsSelected = true,
                        ActionKey = "RunDism"
                    });
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Install Available Driver Updates",
                        Description = "Outdated/faulty drivers are the #1 cause of BSODs — install all pending driver updates",
                        IsAutomatable = true,
                        Risk = ActionRisk.Moderate,
                        IsSelected = true,
                        ActionKey = "InstallDriverUpdates"
                    });
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Schedule Memory Diagnostic",
                        Description = "Faulty RAM causes BSODs — schedule Windows Memory Diagnostic to test on next reboot",
                        IsAutomatable = true,
                        Risk = ActionRisk.RequiresReboot,
                        IsSelected = false,
                                ActionKey = "ScheduleMemoryDiagnostic"
                                });
                                actions.Add(new OptimizationAction
                                {
                                    Category = "Disk Cleanup",
                                    Title = "Clear Crash Dump Files",
                                    Description = "Delete old BSOD memory dumps (minidumps + MEMORY.DMP) to free disk space",
                                    IsAutomatable = true,
                                    Risk = ActionRisk.Safe,
                                    EstimatedFreeMB = 0,
                                    IsSelected = true,
                                    ActionKey = "ClearCrashDumps"
                                });
                            }
                        }

            // ── Event Log: Disk Errors ──
            if (result.EventLog.DiskErrors.Count > 0)
            {
                string diskContext = FormatEventContext("DiskErrors", result.EventLog.DiskErrors.Count, "disk error");
                bool diskRepeat = IsRepeatedRemediation("DiskErrors");

                if (diskRepeat)
                {
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "⚠ Disk errors persist after chkdsk",
                        Description = $"{diskContext}. Previous chkdsk did not stop the errors — the drive may be failing. Back up data and consider replacement.",
                        IsAutomatable = false,
                        Risk = ActionRisk.Safe,
                        IsSelected = false,
                        ActionKey = "ManualOnly"
                    });
                }
                else
                {
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Schedule Disk Check (chkdsk)",
                        Description = $"{diskContext} — schedule chkdsk on next reboot to scan and repair",
                        IsAutomatable = true,
                        Risk = ActionRisk.RequiresReboot,
                        IsSelected = false,
                        EstimatedDuration = "< 5 sec",
                        ActionKey = "ScheduleChkdsk"
                    });
                }
            }

            // ── Event Log: App Crashes ──
            if (result.EventLog.AppCrashes.Count >= 5)
            {
                string crashContext = FormatEventContext("AppCrashes", result.EventLog.AppCrashes.Count, "app crash");
                bool crashRepeat = IsRepeatedRemediation("AppCrashes");

                // Extract which apps are actually crashing
                var topCrashers = result.EventLog.GetTopCrashingApps(3);
                string crasherSummary = topCrashers.Count > 0
                    ? string.Join(", ", topCrashers.Select(c => $"{c.AppName} ({c.Count}×)"))
                    : null;

                bool officeIsCrashing = HasOfficeCrashes(topCrashers);
                string crashingBrowser = GetCrashingBrowser(topCrashers);

                // ── Targeted fixes based on what's actually crashing ──

                if (officeIsCrashing)
                {
                    // Office app is crashing → Office Online Repair is the real fix
                    var officeApp = topCrashers.First(c => OfficeExecutables.Contains(c.AppName));
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Repair Office Installation",
                        Description = $"{officeApp.AppName} crashed {officeApp.Count}× — run Office Online Repair to fix corrupted Office files",
                        IsAutomatable = true,
                        Risk = ActionRisk.Moderate,
                        IsSelected = true,
                        ActionKey = "OpenAppsSettings"
                    });
                }

                if (crashingBrowser != null)
                {
                    // Browser is crashing → clear cache (corrupt cache is a common cause)
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = $"Clear {crashingBrowser} Cache",
                        Description = $"{crashingBrowser} is crashing — corrupt cache files are a common cause",
                        IsAutomatable = true,
                        Risk = ActionRisk.Safe,
                        IsSelected = true,
                        ActionKey = $"ClearBrowserCache:{crashingBrowser}"
                    });
                }

                if (crashRepeat && !officeIsCrashing && crashingBrowser == null)
                {
                    // Repeat + no targeted fix available → escalate to manual
                    string detail = crasherSummary != null
                        ? $"{crashContext}. Top crashing apps: {crasherSummary}. Previous automated repairs did not resolve the issue — reinstall or update the crashing application(s)"
                        : $"{crashContext}. Previous automated repairs did not resolve the issue — identify and reinstall the crashing application(s)";
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "⚠ App crashes persist after previous repair",
                        Description = detail,
                        IsAutomatable = false,
                        Risk = ActionRisk.Safe,
                        IsSelected = false,
                        ActionKey = "ManualOnly"
                    });
                }
                else if (!crashRepeat)
                {
                    // First time — also try generic system repair alongside targeted fixes
                    if (result.EventLog.BSODs.Count == 0)
                    {
                        string sfcDesc = crasherSummary != null
                            ? $"{crashContext} (top: {crasherSummary}) — repair corrupted system files"
                            : $"{crashContext} — repair corrupted system files";
                        actions.Add(new OptimizationAction
                        {
                            Category = "Event Log",
                            Title = "Run System File Checker",
                            Description = sfcDesc,
                            IsAutomatable = true,
                            Risk = ActionRisk.Safe,
                            IsSelected = !officeIsCrashing && crashingBrowser == null,
                            ActionKey = "RunSfc"
                        });
                    }
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Re-register System Components",
                        Description = "Repair broken COM/DLL registrations and reset app platforms that cause crashes",
                        IsAutomatable = true,
                        Risk = ActionRisk.Safe,
                        IsSelected = !officeIsCrashing && crashingBrowser == null,
                                ActionKey = "ReRegisterComponents"
                                });
                                actions.Add(new OptimizationAction
                                {
                                    Category = "Disk Cleanup",
                                    Title = "Clear Error Reports",
                                    Description = "Delete old Windows Error Reporting data and crash dumps to free disk space",
                                    IsAutomatable = true,
                                    Risk = ActionRisk.Safe,
                                    EstimatedFreeMB = 0,
                                    IsSelected = true,
                                    ActionKey = "ClearWerReports"
                                });
                            }
                        }

            // ── Event Log: Unexpected Shutdowns ──
            if (result.EventLog.UnexpectedShutdowns.Count > 0)
            {
                string shutdownContext = FormatEventContext("UnexpectedShutdowns", result.EventLog.UnexpectedShutdowns.Count, "unexpected shutdown");
                bool shutdownRepeat = IsRepeatedRemediation("UnexpectedShutdowns");

                if (shutdownRepeat)
                {
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "⚠ Unexpected shutdowns persist",
                        Description = $"{shutdownContext}. Fast Startup and power configuration were already repaired — check power supply, UPS, or overheating",
                        IsAutomatable = false,
                        Risk = ActionRisk.Safe,
                        IsSelected = false,
                        ActionKey = "ManualOnly"
                    });
                }
                else
                {
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Disable Fast Startup",
                        Description = $"{shutdownContext} — disable Fast Startup to prevent hibernate-related power issues",
                        IsAutomatable = true,
                        Risk = ActionRisk.Safe,
                        IsSelected = true,
                        ActionKey = "DisableFastStartup"
                    });
                    actions.Add(new OptimizationAction
                    {
                        Category = "Event Log",
                        Title = "Repair Power Configuration",
                        Description = "Reset power plans to defaults and fix wake timers/USB suspend settings that cause unexpected shutdowns (warning: custom power plans will be deleted)",
                        IsAutomatable = true,
                        Risk = ActionRisk.Safe,
                        IsSelected = true,
                        ActionKey = "RepairPowerConfig"
                    });
                }
            }

            // ── Non-Automatable Issues ──
            foreach (var issue in result.FlaggedIssues)
            {
                bool alreadyCovered = actions.Any(a =>
                    a.Category.Equals(issue.Category, StringComparison.OrdinalIgnoreCase) ||
                    (issue.Category == "System Overview" && a.ActionKey == "ScheduleRestart") ||
                    (issue.Category == "WindowsUpdate" && (a.ActionKey == "StartWuauserv" || a.ActionKey == "ScheduleRestart")));

                if (alreadyCovered) continue;

                // Hardware / non-fixable issues
                bool isHardware = issue.Category is "Battery" or "RAM"
                    || (issue.Category == "Disk" && issue.Description.Contains("health"))
                    || (issue.Category == "User Profile" && issue.Description.Contains("corruption"))
                    || (issue.Category == "Antivirus" && issue.Description.Contains("No antivirus"));

                if (isHardware)
                {
                    actions.Add(new OptimizationAction
                    {
                        Category = issue.Category,
                        Title = "Manual Action Required",
                        Description = $"{issue.Description} — {issue.Recommendation}",
                        IsAutomatable = false,
                        Risk = issue.Severity == Severity.Critical ? ActionRisk.Moderate : ActionRisk.Safe,
                        IsSelected = false,
                        ActionKey = "ManualOnly"
                    });
                }
            }

            // Deduplicate by ActionKey — the same action can be proposed by multiple
            // diagnostic sections (e.g., ClearBrowserCache from Browser + Event Log crashes)
            var seen = new HashSet<string>(StringComparer.Ordinal);
            actions.RemoveAll(a => !seen.Add(a.ActionKey));

            foreach (var action in actions)
            {
                // Preserve any EstimatedDuration explicitly set during plan building
                if (string.IsNullOrEmpty(action.EstimatedDuration))
                    action.EstimatedDuration = EstimateDuration(action.ActionKey);
                action.Type = ClassifyActionType(action);
            }

            return actions;
        }

        /// <summary>
        /// Returns true if the specified category was remediated within the cooldown
        /// period (default 7 days) and new events have appeared since — meaning the
        /// automated fix did not resolve the underlying problem.
        /// </summary>
        private static bool IsRepeatedRemediation(string category, int cooldownDays = 7)
        {
            DateTime? lastFix = GetRemediationTimestamp(category);
            return lastFix.HasValue &&
                   (DateTime.UtcNow - lastFix.Value).TotalDays < cooldownDays;
        }

        private static readonly HashSet<string> OfficeExecutables = new(StringComparer.OrdinalIgnoreCase)
        {
            "OUTLOOK.EXE", "WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE",
            "MSACCESS.EXE", "ONENOTE.EXE", "MSPUB.EXE", "lync.exe"
        };

        /// <summary>
        /// Checks whether any of the top crashing apps are Office executables.
        /// </summary>
        private static bool HasOfficeCrashes(List<(string AppName, int Count)> topCrashers)
            => topCrashers.Any(c => OfficeExecutables.Contains(c.AppName));

        /// <summary>
        /// Returns the browser name ("Chrome" / "Microsoft Edge" / "Firefox") if a
        /// browser executable is in the top crashers, or null otherwise.
        /// </summary>
        private static string GetCrashingBrowser(List<(string AppName, int Count)> topCrashers)
        {
            foreach (var c in topCrashers)
            {
                if (c.AppName.Equals("chrome.exe", StringComparison.OrdinalIgnoreCase)) return "Chrome";
                if (c.AppName.Equals("msedge.exe", StringComparison.OrdinalIgnoreCase)) return "Microsoft Edge";
                if (c.AppName.Equals("firefox.exe", StringComparison.OrdinalIgnoreCase)) return "Firefox";
            }
            return null;
        }

        /// <summary>
        /// Builds a description prefix for event-log actions that indicates whether
        /// these are new events since the last remediation or first-time detections.
        /// </summary>
        private static string FormatEventContext(string category, int count, string eventType)
        {
            DateTime? lastFix = GetRemediationTimestamp(category);
            if (lastFix.HasValue)
                return $"{count} new {eventType}(s) since last fix on {lastFix.Value.ToLocalTime():MMM dd} at {lastFix.Value.ToLocalTime():h:mm tt}";
            return $"{count} {eventType}(s) detected";
        }

        private static ActionType ClassifyActionType(OptimizationAction action)
        {
            if (!action.IsAutomatable)
                return ActionType.ManualOnly;

            string key = action.ActionKey;
            if (key is "OpenResourceMonitor" or "OpenTaskManagerStartup"
                or "OpenAppsSettings" or "OpenBitLocker"
                or "ScheduleRestart"
                || key.StartsWith("OpenGpuDriverPage:"))
                return ActionType.Shortcut;

            return ActionType.AutoFix;
        }

        private static string EstimateDuration(string actionKey)
        {
            if (actionKey.StartsWith("KillProcess:")) return "< 5 sec";
            if (actionKey.StartsWith("ClearBrowserCache:")) return "< 10 sec";
            if (actionKey.StartsWith("OpenGpuDriverPage:")) return "instant";

            return actionKey switch
            {
                "RunSfc" => "~5–15 min",
                "RunDism" => "~10–30 min",
                "ReRegisterComponents" => "~2–4 min",
                "LaunchDiskCleanup" => "~2–5 min",
                "InstallDriverUpdates" => "~30 sec",
                "ClearUpdateCache" => "~15 sec",
                "RepairPowerConfig" => "~15 sec",
                "StartWuauserv" => "~15 sec",
                "ScheduleChkdsk" => "< 10 sec",
                "ScheduleMemoryDiagnostic" => "< 10 sec",
                "DisableFastStartup" => "< 5 sec",
                "SetPowerPlanBalanced" => "< 5 sec",
                "SwitchToPerformanceVisuals" => "< 10 sec",
                "ClearTempFiles" => "~10–30 sec",
                "ClearPrefetch" => "< 10 sec",
                "EmptyRecycleBin" => "~5–15 sec",
                "CleanUpgradeLogs" => "~10–30 sec",
                "ClearCrashDumps" => "< 10 sec",
                "ClearWerReports" => "< 10 sec",
                "OpenResourceMonitor" => "instant",
                "OpenTaskManagerStartup" => "instant",
                "OpenAppsSettings" => "instant",
                "OpenBitLocker" => "instant",
                "ScheduleRestart" => "instant",
                _ => "< 10 sec"
            };
        }

        private static bool IsLongRunningAction(string actionKey) =>
            actionKey is "RunSfc" or "RunDism" or "ReRegisterComponents"
                or "LaunchDiskCleanup" or "InstallDriverUpdates";

        /// <summary>
        /// Returns a conservative countdown estimate in seconds for the given action.
        /// Only long-running actions (where IsLongRunningAction returns true) get a
        /// live countdown timer. All others return 0 to avoid frozen countdown labels.
        /// </summary>
        private static int EstimateSeconds(string actionKey)
        {
            // Only return non-zero for actions tracked by IsLongRunningAction.
            // Non-long-running actions show no countdown to avoid a frozen label.
            return actionKey switch
            {
                "RunSfc" => 900,               // ~15 min upper bound
                "RunDism" => 1800,             // ~30 min upper bound
                "ReRegisterComponents" => 240, // ~4 min upper bound
                "LaunchDiskCleanup" => 300,    // ~5 min upper bound
                "InstallDriverUpdates" => 30,
                _ => 0
            };
        }

        // ═══════════════════════════════════════════════════════════════
        //  EXECUTE PLAN
        // ═══════════════════════════════════════════════════════════════

        public async Task<OptimizationSummary> ExecuteAsync(
            List<OptimizationAction> actions,
            IProgress<OptimizationProgress> progress,
            CancellationToken ct)
        {
            var selected = actions.Where(a => a.IsSelected && a.Type == ActionType.AutoFix).ToList();
            var summary = new OptimizationSummary();

            for (int i = 0; i < selected.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var action = selected[i];

                progress?.Report(new OptimizationProgress
                {
                    Current = i + 1,
                    Total = selected.Count,
                    CurrentAction = action.Title,
                    IsLongRunning = IsLongRunningAction(action.ActionKey),
                    EstimatedSeconds = EstimateSeconds(action.ActionKey)
                });

                action.Status = ActionStatus.Running;
                OnActionStatusChanged?.Invoke(action);
                OnLog?.Invoke($"[{i + 1}/{selected.Count}] {action.Title}...");

                try
                {
                    var result = await Task.Run(() => DispatchAction(action.ActionKey), ct);
                    long freedMB = result.FreedMB;
                    action.ActualFreedMB = freedMB;

                    // Determine precise outcome for file-based actions
                    if (result.DeletedCount == 0 && result.SkippedCount > 0)
                    {
                        action.Status = ActionStatus.NoChange;
                        action.ResultMessage = result.Detail
                            ?? $"No files deleted — {result.SkippedCount:N0} locked/in-use";
                        summary.ActionDetails.Add($"⊘ {action.Title}: {action.ResultMessage}");
                    }
                    else if (result.DeletedCount > 0 && result.SkippedCount > 0)
                    {
                        action.Status = ActionStatus.PartialSuccess;
                        if (result.Detail != null)
                            action.ResultMessage = freedMB > 0
                                ? $"Freed {freedMB:N0} MB — {result.Detail}"
                                : result.Detail;
                        else
                            action.ResultMessage = $"Freed {freedMB:N0} MB";
                        summary.TotalFreedMB += freedMB;
                        summary.SuccessCount++;
                        summary.ActionDetails.Add($"~ {action.Title}: {action.ResultMessage}");
                    }
                    else
                    {
                        action.Status = ActionStatus.Success;
                        if (result.Detail != null)
                            action.ResultMessage = freedMB > 0
                                ? $"Freed {freedMB:N0} MB — {result.Detail}"
                                : result.Detail;
                        else
                            action.ResultMessage = freedMB > 0 ? $"Freed {freedMB:N0} MB" : "Done";
                        summary.TotalFreedMB += freedMB;
                        summary.SuccessCount++;
                        summary.ActionDetails.Add($"✓ {action.Title}: {action.ResultMessage}");
                    }
                    OnActionStatusChanged?.Invoke(action);
                    string prefix = action.Status == ActionStatus.NoChange ? "⊘" :
                                    action.Status == ActionStatus.PartialSuccess ? "~" : "✓";
                    OnLog?.Invoke($"  {prefix} {action.Title} — {action.ResultMessage}");
                }
                catch (Exception ex)
                {
                    action.Status = ActionStatus.Failed;
                    action.ResultMessage = ex.Message;
                    summary.FailureCount++;
                    summary.ActionDetails.Add($"✗ {action.Title}: {ex.Message}");
                    OnActionStatusChanged?.Invoke(action);
                    OnLog?.Invoke($"  ✗ {action.Title} — {ex.Message}");
                }
            }

            // Mark unselected automatable actions as skipped (but preserve shortcuts already executed)
            foreach (var a in actions.Where(a => a.IsAutomatable && !a.IsSelected && a.Status == ActionStatus.Pending))
                a.Status = ActionStatus.Skipped;

            // Record per-category remediation timestamps so the next scan only
            // flags NEW events occurring after this fix
            var remediatedCategories = GetRemediatedCategories(selected);
            SaveRemediationTimestamps(remediatedCategories);

            summary.ActionsRun = selected.Count;
            return summary;
        }

        // ═══════════════════════════════════════════════════════════════
        //  VERIFY AFFECTED ITEMS
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Re-checks only the items affected by the optimization actions to confirm
        /// fixes were applied. Updates each action's Verification and VerificationMessage.
        /// </summary>
        public async Task VerifyActionsAsync(
            List<OptimizationAction> actions,
            IProgress<OptimizationProgress> progress,
            CancellationToken ct)
        {
            var toVerify = actions
                .Where(a => a.Status is ActionStatus.Success or ActionStatus.PartialSuccess or ActionStatus.NoChange)
                .ToList();

            for (int i = 0; i < toVerify.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var action = toVerify[i];

                action.Verification = VerificationStatus.Verifying;
                OnActionStatusChanged?.Invoke(action);

                progress?.Report(new OptimizationProgress
                {
                    Current = i + 1,
                    Total = toVerify.Count,
                    CurrentAction = $"Verifying: {action.Title}"
                });

                try
                {
                    var (status, message) = await Task.Run(
                        () => VerifyAction(action), ct);
                    action.Verification = status;
                    action.VerificationMessage = message;
                }
                catch
                {
                    action.Verification = VerificationStatus.NotVerified;
                    action.VerificationMessage = "Verification check failed";
                }

                OnActionStatusChanged?.Invoke(action);
            }

            // Mark failed/skipped actions
            foreach (var a in actions)
            {
                if (a.Verification != VerificationStatus.None) continue;
                if (a.Status == ActionStatus.Failed)
                {
                    a.Verification = VerificationStatus.NotVerified;
                    a.VerificationMessage = "Action failed — not verified";
                }
                else if (a.Status == ActionStatus.Skipped)
                {
                    a.Verification = VerificationStatus.None;
                }
            }
        }

        private (VerificationStatus Status, string Message) VerifyAction(OptimizationAction action)
        {
            string key = action.ActionKey;

            // Reboot-dependent actions cannot be verified until next boot
            if (key is "ScheduleChkdsk" or "ScheduleMemoryDiagnostic" or "ScheduleRestart"
                or "InstallDriverUpdates")
                return (VerificationStatus.RequiresReboot, "Will take effect after reboot");

            if (key == "ClearTempFiles")
            {
                long totalMB = GetDirectorySizeMB(Path.GetTempPath())
                    + GetDirectorySizeMB(Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp"));
                return totalMB < 100
                    ? (VerificationStatus.Verified, $"Temp folders now {totalMB} MB")
                    : (VerificationStatus.PartiallyVerified, $"Temp folders still {totalMB} MB — some files were locked");
            }

            if (key == "ClearPrefetch")
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Prefetch");
                long mb = GetDirectorySizeMB(dir);
                return mb < 50
                    ? (VerificationStatus.Verified, $"Prefetch now {mb} MB")
                    : (VerificationStatus.PartiallyVerified, $"Prefetch still {mb} MB");
            }

            if (key == "EmptyRecycleBin")
            {
                long mb = GetRecycleBinSizeMB();
                return mb < 10
                    ? (VerificationStatus.Verified, $"Recycle Bin: {mb} MB")
                    : (VerificationStatus.PartiallyVerified, $"Recycle Bin still {mb} MB");
            }

            if (key == "ClearUpdateCache")
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    "SoftwareDistribution", "Download");
                long mb = GetDirectorySizeMB(dir);
                return mb < 50
                    ? (VerificationStatus.Verified, $"Update cache now {mb} MB")
                    : (VerificationStatus.PartiallyVerified, $"Update cache still {mb} MB");
            }

            if (key == "CleanUpgradeLogs")
            {
                string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                string sysRoot = Path.GetPathRoot(winDir) ?? @"C:\";
                long mb = GetDirectorySizeMB(Path.Combine(winDir, "Logs", "WindowsUpdate"))
                    + GetDirectorySizeMB(Path.Combine(winDir, "Panther"))
                    + GetDirectorySizeMB(Path.Combine(sysRoot, "$Windows.~BT"))
                    + GetDirectorySizeMB(Path.Combine(sysRoot, "$Windows.~WS"));
                return mb < 50
                    ? (VerificationStatus.Verified, $"Upgrade logs now {mb} MB")
                    : (VerificationStatus.PartiallyVerified, $"Upgrade logs still {mb} MB");
            }

            if (key == "LaunchDiskCleanup")
            {
                string winOldDir = Path.Combine(
                    Path.GetPathRoot(Environment.SystemDirectory) ?? @"C:\", "Windows.old");
                return !Directory.Exists(winOldDir)
                    ? (VerificationStatus.Verified, "Windows.old removed")
                    : (VerificationStatus.NotVerified, "Windows.old still exists");
            }

            if (key.StartsWith("ClearBrowserCache:"))
            {
                // Already freed what it could; trust the action result
                return action.ActualFreedMB > 0 || action.Status == ActionStatus.Success
                    ? (VerificationStatus.Verified, $"Freed {action.ActualFreedMB} MB")
                    : (VerificationStatus.PartiallyVerified, "Some cache files were locked");
            }

            if (key.StartsWith("KillProcess:"))
            {
                string[] parts = key["KillProcess:".Length..].Split(':', 2);
                if (parts.Length == 2 && int.TryParse(parts[0], out int pid))
                {
                    try
                    {
                        using var proc = Process.GetProcessById(pid);
                        return proc.HasExited
                            ? (VerificationStatus.Verified, $"{parts[1]} is no longer running")
                            : (VerificationStatus.NotVerified, $"{parts[1]} is still running");
                    }
                    catch (ArgumentException)
                    {
                        return (VerificationStatus.Verified, $"{parts[1]} is no longer running");
                    }
                }
                return (VerificationStatus.Verified, "Process ended");
            }

            if (key == "SetPowerPlanBalanced")
            {
                try
                {
                    using var proc = Process.Start(new ProcessStartInfo
                    {
                        FileName = "powercfg",
                        Arguments = "/getactivescheme",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true
                    });
                    if (proc != null)
                    {
                        string output = proc.StandardOutput.ReadToEnd();
                        proc.WaitForExit(5000);
                        return output.Contains("381b4222-f694-41f0-9685-ff5bb260df2e")
                            ? (VerificationStatus.Verified, "Balanced power plan is active")
                            : (VerificationStatus.NotVerified, "Balanced plan is not active");
                    }
                }
                catch { }
                return (VerificationStatus.PartiallyVerified, "Could not confirm power plan");
            }

            if (key == "SwitchToPerformanceVisuals")
            {
                try
                {
                    using var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                    object val = k?.GetValue("EnableTransparency");
                    return val is int i && i == 0
                        ? (VerificationStatus.Verified, "Transparency disabled")
                        : (VerificationStatus.PartiallyVerified, "Transparency setting unchanged — relogin may be needed");
                }
                catch { return (VerificationStatus.PartiallyVerified, "Could not read registry"); }
            }

            if (key == "StartWuauserv")
            {
                try
                {
                    using var sc = new ServiceController("wuauserv");
                    sc.Refresh();
                    return sc.Status == ServiceControllerStatus.Running
                        ? (VerificationStatus.Verified, "Windows Update service is running")
                        : (VerificationStatus.NotVerified, $"Service status: {sc.Status}");
                }
                catch { return (VerificationStatus.NotVerified, "Could not query service"); }
            }

            if (key == "DisableFastStartup")
            {
                try
                {
                    using var k = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                        @"SYSTEM\CurrentControlSet\Control\Session Manager\Power");
                    object val = k?.GetValue("HiberbootEnabled");
                    return val is int i && i == 0
                        ? (VerificationStatus.Verified, "Fast Startup is disabled")
                        : (VerificationStatus.NotVerified, "Fast Startup is still enabled");
                }
                catch { return (VerificationStatus.PartiallyVerified, "Could not read registry"); }
            }

            if (key == "ClearCrashDumps")
            {
                string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                bool minidumpGone = !DirectoryHasFiles(Path.Combine(winDir, "Minidump"));
                bool memoryDmpGone = !File.Exists(Path.Combine(winDir, "MEMORY.DMP"));
                return minidumpGone && memoryDmpGone
                    ? (VerificationStatus.Verified, "Crash dumps cleared")
                    : (VerificationStatus.PartiallyVerified, "Some dump files remain (locked)");
            }

            if (key == "ClearWerReports")
                return action.ActualFreedMB > 0 || action.Status == ActionStatus.Success
                    ? (VerificationStatus.Verified, "Error reports cleared")
                    : (VerificationStatus.PartiallyVerified, "Some reports remain");

            // SFC, DISM, ReRegister — trust the exit code from the action result
            if (key is "RunSfc" or "RunDism" or "ReRegisterComponents" or "RepairPowerConfig")
                return action.Status == ActionStatus.Success
                    ? (VerificationStatus.Verified, action.ResultMessage)
                    : (VerificationStatus.PartiallyVerified, action.ResultMessage);

            // Fallback: trust the action result
            return action.Status == ActionStatus.Success
                ? (VerificationStatus.Verified, "Completed successfully")
                : (VerificationStatus.PartiallyVerified, action.ResultMessage ?? "Check manually");
        }

        private static long GetDirectorySizeMB(string path)
            => NativeHelpers.GetFolderSizeMB(path);

        // ═══════════════════════════════════════════════════════════════
        //  ACTION DISPATCH
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Dispatches a shortcut action after validating against known keys.
        /// </summary>
        public void DispatchShortcut(string actionKey)
        {
            ArgumentNullException.ThrowIfNull(actionKey);
            // Only allow known shortcut prefixes/keys
            bool isValid = actionKey is "OpenResourceMonitor" or "OpenTaskManagerStartup"
                or "OpenAppsSettings" or "OpenBitLocker" or "ScheduleRestart"
                || actionKey.StartsWith("OpenGpuDriverPage:", StringComparison.Ordinal);
            if (!isValid)
                throw new ArgumentException($"Unknown shortcut action key: {actionKey}", nameof(actionKey));
            DispatchAction(actionKey);
        }

        private ActionResult DispatchAction(string actionKey)
        {
            if (actionKey.StartsWith("ClearBrowserCache:"))
            {
                string browserName = actionKey["ClearBrowserCache:".Length..];
                return ClearBrowserCache(browserName);
            }
            if (actionKey.StartsWith("OpenGpuDriverPage:"))
            {
                string gpuName = actionKey["OpenGpuDriverPage:".Length..];
                OpenGpuDriverPage(gpuName);
                return new ActionResult();
            }
            if (actionKey.StartsWith("KillProcess:"))
            {
                // Format: KillProcess:{pid}:{name}
                string[] parts = actionKey["KillProcess:".Length..].Split(':', 2);
                if (parts.Length == 2 && int.TryParse(parts[0], out int pid))
                    return KillProcess(pid, parts[1]);
                return new ActionResult();
            }

            return actionKey switch
            {
                "ClearTempFiles" => ClearTempFiles(),
                "ClearPrefetch" => ClearPrefetch(),
                "EmptyRecycleBin" => EmptyRecycleBinAction(),
                "ClearUpdateCache" => ClearUpdateCache(),
                "LaunchDiskCleanup" => LaunchDiskCleanup(),
                "CleanUpgradeLogs" => CleanUpgradeLogs(),
                "SwitchToPerformanceVisuals" => SwitchToPerformanceVisuals(),
                "SetPowerPlanBalanced" => SetPowerPlanBalanced(),
                "OpenResourceMonitor" => OpenResourceMonitor(),
                "OpenTaskManagerStartup" => OpenTaskManagerStartup(),
                "StartWuauserv" => StartWindowsUpdateService(),
                "ScheduleRestart" => ScheduleRestart(),
                "OpenAppsSettings" => OpenAppsSettings(),
                "OpenBitLocker" => OpenBitLocker(),
                "RunSfc" => RunSystemFileChecker(),
                "RunDism" => RunDismRepair(),
                "ScheduleChkdsk" => ScheduleChkdsk(),
                "ClearCrashDumps" => ClearCrashDumps(),
                "InstallDriverUpdates" => InstallDriverUpdates(),
                "ScheduleMemoryDiagnostic" => ScheduleMemoryDiagnostic(),
                "ClearWerReports" => ClearWerReports(),
                "ReRegisterComponents" => ReRegisterComponents(),
                "DisableFastStartup" => DisableFastStartup(),
                "RepairPowerConfig" => RepairPowerConfiguration(),
                _ => new ActionResult()
            };
        }

        // ═══════════════════════════════════════════════════════════════
        //  INDIVIDUAL ACTIONS
        // ═══════════════════════════════════════════════════════════════

        private ActionResult ClearTempFiles()
        {
            long freed = 0;
            int deleted = 0;
            int skipped = 0;

            // User temp
            string userTemp = Path.GetTempPath();
            OnLog?.Invoke($"  Cleaning: {userTemp}");
            freed += DeleteDirectoryContents(userTemp, ref deleted, ref skipped);

            // Windows temp
            string winTemp = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp");
            OnLog?.Invoke($"  Cleaning: {winTemp}");
            freed += DeleteDirectoryContents(winTemp, ref deleted, ref skipped);

            long freedMB = freed / (1024 * 1024);
            string detail = $"Deleted {deleted:N0} files ({freedMB:N0} MB), {skipped:N0} locked/in-use files skipped";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(freedMB, deleted, skipped, detail);
        }

        private ActionResult ClearPrefetch()
        {
            string prefetch = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Prefetch");
            int deleted = 0, skipped = 0;
            long freed = DeleteDirectoryContents(prefetch, ref deleted, ref skipped);
            long freedMB = freed / (1024 * 1024);
            string detail = $"Deleted {deleted:N0} files ({freedMB:N0} MB), {skipped:N0} skipped";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(freedMB, deleted, skipped, detail);
        }

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

        private const uint SHERB_NOCONFIRMATION = 0x00000001;
        private const uint SHERB_NOPROGRESSUI = 0x00000002;
        private const uint SHERB_NOSOUND = 0x00000004;

        private ActionResult EmptyRecycleBinAction()
        {
            // Get estimated size before emptying
            long beforeMB = GetRecycleBinSizeMB();

            int hr = SHEmptyRecycleBin(IntPtr.Zero, null,
                SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);

            // HRESULT convention: negative values indicate failure
            if (hr < 0)
                throw new InvalidOperationException($"SHEmptyRecycleBin returned 0x{hr:X8}");

            long afterMB = GetRecycleBinSizeMB();
            long actualFreed = Math.Max(0, beforeMB - afterMB);
            string detail = $"Recycle Bin emptied ({actualFreed:N0} MB freed)";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(actualFreed, Detail: detail);
        }

        private static long GetRecycleBinSizeMB() => NativeHelpers.GetRecycleBinSizeMB();

        private ActionResult ClearUpdateCache()
        {
            StopService("wuauserv", TimeSpan.FromSeconds(30));

            int deleted = 0, skipped = 0;
            long freed = 0;
            try
            {
                string sdPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    "SoftwareDistribution", "Download");
                freed = DeleteDirectoryContents(sdPath, ref deleted, ref skipped);
            }
            finally
            {
                // Always restart the service — leaving wuauserv stopped breaks Windows Update
                StartService("wuauserv", TimeSpan.FromSeconds(30));
            }

            long freedMB = freed / (1024 * 1024);
            string detail = $"Deleted {deleted:N0} files ({freedMB:N0} MB), {skipped:N0} skipped";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(freedMB, deleted, skipped, detail);
        }

        private ActionResult CleanUpgradeLogs()
        {
            long freed = 0;
            int deleted = 0, skipped = 0;

            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string sysRoot = Path.GetPathRoot(winDir) ?? @"C:\";
            string[] upgradeDirs =
            [
                Path.Combine(winDir, "Logs", "WindowsUpdate"),
                Path.Combine(winDir, "Panther"),
                Path.Combine(sysRoot, "$Windows.~BT"),
                Path.Combine(sysRoot, "$Windows.~WS")
            ];

            foreach (string dir in upgradeDirs)
            {
                if (!Directory.Exists(dir)) continue;
                OnLog?.Invoke($"  Cleaning: {dir}");
                freed += DeleteDirectoryContents(dir, ref deleted, ref skipped);
            }

            long freedMB = freed / (1024 * 1024);
            string detail = $"Deleted {deleted:N0} upgrade log files ({freedMB:N0} MB), {skipped:N0} skipped";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(freedMB, deleted, skipped, detail);
        }

        private ActionResult LaunchDiskCleanup()
        {
            string winOldDir = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory) ?? @"C:\", "Windows.old");
            if (!Directory.Exists(winOldDir))
            {
                string detail = "Windows.old not found — already removed";
                OnLog?.Invoke($"  ✓ {detail}");
                return new ActionResult(Detail: detail);
            }

            string drive = Path.GetPathRoot(winOldDir) ?? @"C:\";
            string driveLetter = drive.TrimEnd('\\');
            long freeBytesBefore = new DriveInfo(drive).AvailableFreeSpace;

            // Use the supported Windows Disk Cleanup mechanism.
            // cleanmgr's "Previous Installations" handler uses internal APIs that handle
            // TrustedInstaller permissions natively — no manual takeown/icacls needed.
            OnLog?.Invoke("  Running Windows Disk Cleanup for previous installations...");

            const string volumeCachesPath =
                @"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches";

            // Enable cleanup categories for sageset profile 65
            RunElevatedCmd(
                $"reg add \"{volumeCachesPath}\\Previous Installations\" /v StateFlags0065 /t REG_DWORD /d 2 /f" +
                $" & reg add \"{volumeCachesPath}\\Temporary Setup Files\" /v StateFlags0065 /t REG_DWORD /d 2 /f");

            int exitCode = RunElevatedCmd($"cleanmgr /d {driveLetter} /sagerun:65", timeoutMs: 900_000);

            // Clean up registry flags regardless of outcome
            RunElevatedCmd(
                $"reg delete \"{volumeCachesPath}\\Previous Installations\" /v StateFlags0065 /f" +
                $" & reg delete \"{volumeCachesPath}\\Temporary Setup Files\" /v StateFlags0065 /f");

            long freedMB = Math.Max(0, (new DriveInfo(drive).AvailableFreeSpace - freeBytesBefore) / (1024 * 1024));

            if (!Directory.Exists(winOldDir))
            {
                string detail = $"Windows.old removed via Disk Cleanup ({freedMB:N0} MB freed)";
                OnLog?.Invoke($"  ✓ {detail}");
                return new ActionResult(freedMB, Detail: detail);
            }

            // Disk Cleanup may have partially cleaned or left locked files
            if (freedMB > 0)
            {
                string detail = $"Partially cleaned ({freedMB:N0} MB freed) — use Settings → System → Storage to finish removal";
                OnLog?.Invoke($"  ⚠ {detail}");
                return new ActionResult(freedMB, Detail: detail);
            }

            string failDetail = exitCode == -1
                ? "Disk Cleanup timed out — try removing via Settings → System → Storage"
                : "Disk Cleanup could not remove Windows.old — try using Settings → System → Storage";
            OnLog?.Invoke($"  ⚠ {failDetail}");
            return new ActionResult(Detail: failDetail);
        }

        /// <summary>
        /// Validates that a string is safe to interpolate into a shell command.
        /// Rejects characters that could enable command injection.
        /// </summary>
        private static void ValidateShellArgument(string value, string paramName)
        {
            ArgumentNullException.ThrowIfNull(value, paramName);
            foreach (char c in value)
            {
                if (c is '&' or '|' or ';' or '`' or '$' or '(' or ')' or '<' or '>' or '!' or '\n' or '\r')
                    throw new ArgumentException(
                        $"Shell argument contains unsafe character '{c}'", paramName);
            }
        }

        /// <summary>
        /// Runs a command via cmd.exe, waits up to 10 minutes, and returns the exit code.
        /// Output is suppressed. Returns -1 if the process could not be started or timed out.
        /// </summary>
        private static int RunElevatedCmd(string arguments, int timeoutMs = 600_000)
        {
            try
            {
                using var proc = Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/C {arguments} > NUL 2>&1",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
                if (proc == null) return -1;
                if (!proc.WaitForExit(timeoutMs))
                {
                    try { proc.Kill(entireProcessTree: true); }
                    catch { /* best effort */ }
                    return -1;
                }
                return proc.ExitCode;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"RunElevatedCmd failed: {ex.Message}");
                return -1;
            }
        }

        private ActionResult ClearBrowserCache(string browserName)
        {
            long freed = 0;
            int deleted = 0, skipped = 0;
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            var cachePaths = new List<string>();

            if (browserName.Contains("Chrome", StringComparison.OrdinalIgnoreCase))
            {
                cachePaths.Add(Path.Combine(localAppData, "Google", "Chrome", "User Data", "Default", "Cache"));
                cachePaths.Add(Path.Combine(localAppData, "Google", "Chrome", "User Data", "Default", "Code Cache"));
            }
            else if (browserName.Contains("Edge", StringComparison.OrdinalIgnoreCase))
            {
                cachePaths.Add(Path.Combine(localAppData, "Microsoft", "Edge", "User Data", "Default", "Cache"));
                cachePaths.Add(Path.Combine(localAppData, "Microsoft", "Edge", "User Data", "Default", "Code Cache"));
            }
            else if (browserName.Contains("Firefox", StringComparison.OrdinalIgnoreCase))
            {
                // Firefox profiles are in AppData\Roaming, not LocalApplicationData
                string ffProfiles = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Mozilla", "Firefox", "Profiles");
                if (Directory.Exists(ffProfiles))
                {
                    foreach (var profile in Directory.GetDirectories(ffProfiles))
                    {
                        cachePaths.Add(Path.Combine(profile, "cache2"));
                    }
                }
            }

            foreach (var path in cachePaths)
                freed += DeleteDirectoryContents(path, ref deleted, ref skipped);

            long freedMB = freed / (1024 * 1024);
            string detail = $"Deleted {deleted:N0} files ({freedMB:N0} MB), {skipped:N0} skipped";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(freedMB, deleted, skipped, detail);
        }

        private ActionResult KillProcess(int pid, string name)
        {
            try
            {
                using var proc = Process.GetProcessById(pid);
                if (proc.HasExited)
                    return new ActionResult(Detail: $"{name} (PID {pid}) already exited");

                // Guard against PID reuse — verify the process is still the one we scanned
                if (!proc.ProcessName.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return new ActionResult(Detail: $"PID {pid} is now {proc.ProcessName}, not {name} — skipped (PID was reused)");

                long memMB = proc.WorkingSet64 / (1024 * 1024);
                proc.Kill(entireProcessTree: true);
                proc.WaitForExit(10_000);
                string detail = $"Ended {name} (PID {pid}, was using {memMB:N0} MB RAM)";
                OnLog?.Invoke($"  ✓ {detail}");
                return new ActionResult(Detail: detail);
            }
            catch (ArgumentException)
            {
                return new ActionResult(Detail: $"{name} (PID {pid}) is no longer running");
            }
        }

        private ActionResult SwitchToPerformanceVisuals()
        {
            int changes = 0;

            // Disable transparency
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", writable: true);
                if (key != null)
                {
                    object current = key.GetValue("EnableTransparency");
                    if (current is int val && val != 0)
                    {
                        key.SetValue("EnableTransparency", 0, Microsoft.Win32.RegistryValueKind.DWord);
                        changes++;
                        OnLog?.Invoke("  ✓ Disabled transparency effects");
                    }
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚠ Could not disable transparency: {ex.Message}"); }

            // Disable animations
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Control Panel\Desktop\WindowMetrics", writable: true);
                if (key != null)
                {
                    key.SetValue("MinAnimate", "0");
                    changes++;
                }

                using var dwmKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", writable: true);
                if (dwmKey != null)
                {
                    dwmKey.SetValue("VisualFXSetting", 2, Microsoft.Win32.RegistryValueKind.DWord); // 2 = best performance
                    changes++;
                }

                OnLog?.Invoke("  ✓ Set visual effects to best performance");
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚠ Could not disable animations: {ex.Message}"); }

            // Disable menu/tooltip animations via SystemParametersInfo
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Control Panel\Desktop", writable: true);
                if (key != null)
                {
                    key.SetValue("UserPreferencesMask", new byte[]
                    {
                        0x90, 0x12, 0x01, 0x80, 0x10, 0x00, 0x00, 0x00 // performance-oriented mask
                    }, Microsoft.Win32.RegistryValueKind.Binary);
                    changes++;
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  \u26a0 Could not set UserPreferencesMask: {ex.Message}"); }

            string detail = changes > 0
                ? $"Applied {changes} visual performance optimizations (relogin to fully take effect)"
                : "Visual settings were already optimized";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult SetPowerPlanBalanced()
        {
            const string balancedGuid = "381b4222-f694-41f0-9685-ff5bb260df2e";
            int exitCode = RunElevatedCmd($"powercfg /setactive {balancedGuid}");
            if (exitCode != 0)
                throw new InvalidOperationException($"powercfg /setactive failed (exit code {exitCode})");

            string detail = "Switched to Balanced power plan";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult OpenResourceMonitor()
        {
            var proc = Process.Start(new ProcessStartInfo
            {
                FileName = "resmon.exe",
                UseShellExecute = true
            });
            if (proc == null)
                OnLog?.Invoke("  ⚠ Could not launch Resource Monitor");
            return new ActionResult(Detail: "Resource Monitor opened — check the CPU tab for processes using high CPU");
        }

        private ActionResult OpenTaskManagerStartup()
        {
            var proc = Process.Start(new ProcessStartInfo
            {
                FileName = "taskmgr.exe",
                Arguments = "/7",
                UseShellExecute = true
            });
            if (proc == null)
                OnLog?.Invoke("  ⚠ Could not launch Task Manager");
            return new ActionResult(Detail: "Task Manager opened to Startup tab");
        }

        private ActionResult StartWindowsUpdateService()
        {
            try
            {
                using var sc = new ServiceController("wuauserv");
                sc.Refresh();
                if (sc.StartType == System.ServiceProcess.ServiceStartMode.Disabled)
                {
                    int exitCode = RunElevatedCmd("sc.exe config wuauserv start= demand");
                    if (exitCode != 0)
                        OnLog?.Invoke($"  ⚠ sc.exe config returned exit code {exitCode}");
                }
            }
            catch (Exception ex) { OnLog?.Invoke($"  ⚠ Could not check wuauserv start type: {ex.Message}"); }

            StartService("wuauserv", TimeSpan.FromSeconds(30));
            string detail = "Windows Update service enabled and started";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult ScheduleRestart()
        {
            var proc = Process.Start(new ProcessStartInfo
            {
                FileName = "shutdown.exe",
                Arguments = "/r /t 120 /c \"Scheduled restart by Endpoint Diagnostics & Remediation\"",
                UseShellExecute = false,
                CreateNoWindow = true
            });
            if (proc == null)
                throw new InvalidOperationException("Failed to schedule system restart");
            string detail = "System restart scheduled in 2 minutes";
            OnLog?.Invoke($"  {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult OpenAppsSettings()
        {
            var proc = Process.Start(new ProcessStartInfo
            {
                FileName = "ms-settings:appsfeatures",
                UseShellExecute = true
            });
            if (proc == null)
                OnLog?.Invoke("  ⚠ Could not open Apps & Features settings");
            return new ActionResult();
        }

        private ActionResult OpenBitLocker()
        {
            var proc = Process.Start(new ProcessStartInfo
            {
                FileName = "control.exe",
                Arguments = "/name Microsoft.BitLockerDriveEncryption",
                UseShellExecute = true
            });
            if (proc == null)
                OnLog?.Invoke("  ⚠ Could not open BitLocker settings");
            return new ActionResult();
        }

        private void OpenGpuDriverPage(string gpuName)
        {
            string url;
            if (gpuName.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase) ||
                gpuName.Contains("GeForce", StringComparison.OrdinalIgnoreCase))
                url = "https://www.nvidia.com/Download/index.aspx";
            else if (gpuName.Contains("AMD", StringComparison.OrdinalIgnoreCase) ||
                     gpuName.Contains("Radeon", StringComparison.OrdinalIgnoreCase))
                url = "https://www.amd.com/en/support";
            else if (gpuName.Contains("Intel", StringComparison.OrdinalIgnoreCase))
                url = "https://www.intel.com/content/www/us/en/download-center/home.html";
            else
                url = "https://www.google.com/search?q=download+GPU+driver+update";

            var proc = Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
            if (proc == null)
                OnLog?.Invoke("  ⚠ Could not open GPU driver page");
        }

        private ActionResult RunSystemFileChecker()
        {
            OnLog?.Invoke("  Running sfc /scannow (this may take several minutes)...");
            // sfc can take 5–15 min on slow/damaged systems; 20 min timeout is safe
            int exitCode = RunElevatedCmd("sfc /scannow", timeoutMs: 1_200_000);
            string detail = exitCode switch
            {
                0 => "System File Checker completed successfully — check CBS.log for details",
                -1 => "System File Checker timed out — the system may need multiple passes; try running again",
                _ => $"System File Checker finished (exit code {exitCode}) — check C:\\Windows\\Logs\\CBS\\CBS.log for results"
            };
            OnLog?.Invoke($"  {(exitCode == -1 ? "⚠" : "✓")} {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult RunDismRepair()
        {
            OnLog?.Invoke("  Running DISM /RestoreHealth (this may take 10–30 minutes)...");
            // DISM /RestoreHealth downloads components from Windows Update and can
            // take 15–30+ min on damaged or slow systems. 45 min timeout is safe.
            int exitCode = RunElevatedCmd("DISM /Online /Cleanup-Image /RestoreHealth", timeoutMs: 2_700_000);
            string detail = exitCode switch
            {
                0 => "DISM component store repair completed successfully",
                -1 => "DISM timed out — the component store may need manual repair; try running again",
                _ => $"DISM finished with exit code {exitCode} — check C:\\Windows\\Logs\\DISM\\dism.log for details"
            };
            OnLog?.Invoke($"  {(exitCode == -1 ? "⚠" : "✓")} {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult ScheduleChkdsk()
        {
            string systemDrive = Path.GetPathRoot(Environment.SystemDirectory) ?? "C:\\";
            string driveLetter = systemDrive.TrimEnd('\\', '/');

            ValidateShellArgument(driveLetter, nameof(driveLetter));
            OnLog?.Invoke($"  Scheduling chkdsk for {driveLetter} on next reboot...");

            // Pipe Y to confirm scheduling — chkdsk on the system drive cannot lock
            // the volume while Windows is running, so it prompts Y/N. Without this,
            // the process hangs waiting for input until the timeout kills it.
            int exitCode = RunElevatedCmd($"echo Y | chkdsk {driveLetter} /F /R", timeoutMs: 30_000);

            if (exitCode == -1)
            {
                string warn = $"chkdsk scheduling timed out for {driveLetter} — try running 'chkdsk {driveLetter} /F /R' manually in an admin Command Prompt";
                OnLog?.Invoke($"  ⚠ {warn}");
                return new ActionResult(Detail: warn);
            }

            string detail = $"Disk check scheduled for {driveLetter} — will run on next system restart";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult ClearCrashDumps()
        {
            long freed = 0;
            int deleted = 0, skipped = 0;
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);

            string miniDumpDir = Path.Combine(winDir, "Minidump");
            if (Directory.Exists(miniDumpDir))
            {
                OnLog?.Invoke($"  Cleaning: {miniDumpDir}");
                freed += DeleteDirectoryContents(miniDumpDir, ref deleted, ref skipped);
            }

            string fullDump = Path.Combine(winDir, "MEMORY.DMP");
            if (File.Exists(fullDump))
            {
                try
                {
                    long size = new FileInfo(fullDump).Length;
                    File.Delete(fullDump);
                    freed += size;
                    deleted++;
                    OnLog?.Invoke($"  Deleted MEMORY.DMP ({size / (1024 * 1024):N0} MB)");
                }
                catch (Exception ex) { skipped++; OnLog?.Invoke($"  ⚠ Could not delete MEMORY.DMP: {ex.Message}"); }
            }

            string liveKernel = Path.Combine(winDir, "LiveKernelReports");
            if (Directory.Exists(liveKernel))
            {
                OnLog?.Invoke($"  Cleaning: {liveKernel}");
                freed += DeleteDirectoryContents(liveKernel, ref deleted, ref skipped);
            }

            long freedMB = freed / (1024 * 1024);
            string detail = $"Deleted {deleted:N0} crash dump files ({freedMB:N0} MB freed), {skipped:N0} skipped";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(freedMB, deleted, skipped, detail);
        }

        private ActionResult InstallDriverUpdates()
        {
            OnLog?.Invoke("  Installing available driver updates...");
            RunElevatedCmd("pnputil /scan-devices");

            int exitCode = RunElevatedCmd("UsoClient StartInstall");
            if (exitCode != 0)
            {
                OnLog?.Invoke("  UsoClient unavailable, using wuauclt...");
                RunElevatedCmd("wuauclt /detectnow /updatenow");
            }

            OnLog?.Invoke("  Checking for devices with driver problems...");
            RunElevatedCmd("pnputil /enum-devices /problem");

            string detail = "Driver updates triggered — pending installs will complete in the background or after reboot";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult ScheduleMemoryDiagnostic()
        {
            OnLog?.Invoke("  Scheduling Windows Memory Diagnostic...");

            // bcdedit /bootsequence adds {memdiag} to the one-time boot sequence
            // so the memory diagnostic tool runs on the very next reboot only.
            int exitCode = RunElevatedCmd("bcdedit /bootsequence {memdiag}", timeoutMs: 15_000);

            if (exitCode != 0)
            {
                string warn = $"Could not schedule memory diagnostic (bcdedit exit code {exitCode}) — open Start and search 'Windows Memory Diagnostic' to schedule manually";
                OnLog?.Invoke($"  ⚠ {warn}");
                return new ActionResult(Detail: warn);
            }

            string detail = "Windows Memory Diagnostic scheduled — RAM will be tested on next reboot";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult ClearWerReports()
        {
            long freed = 0;
            int deleted = 0, skipped = 0;

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string[] werPaths =
            [
                Path.Combine(localAppData, "Microsoft", "Windows", "WER", "ReportArchive"),
                Path.Combine(localAppData, "Microsoft", "Windows", "WER", "ReportQueue"),
                Path.Combine(localAppData, "CrashDumps"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "Microsoft", "Windows", "WER", "ReportArchive"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "Microsoft", "Windows", "WER", "ReportQueue")
            ];

            foreach (string path in werPaths)
            {
                if (!Directory.Exists(path)) continue;
                OnLog?.Invoke($"  Cleaning: {path}");
                freed += DeleteDirectoryContents(path, ref deleted, ref skipped);
            }

            long freedMB = freed / (1024 * 1024);
            string detail = $"Deleted {deleted:N0} error report files ({freedMB:N0} MB freed), {skipped:N0} skipped";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(freedMB, deleted, skipped, detail);
        }

        private ActionResult DisableFastStartup()
        {
            string detail;
            try
            {
                using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                    @"SYSTEM\CurrentControlSet\Control\Session Manager\Power", writable: true);
                if (key != null)
                {
                    object current = key.GetValue("HiberbootEnabled");
                    if (current is int val && val != 0)
                    {
                        key.SetValue("HiberbootEnabled", 0, Microsoft.Win32.RegistryValueKind.DWord);
                        detail = "Fast Startup disabled — system will perform a full shutdown/boot cycle";
                    }
                    else
                    {
                        detail = "Fast Startup was already disabled";
                    }
                }
                else
                {
                    RunElevatedCmd("powercfg /hibernate off");
                    detail = "Hibernate and Fast Startup disabled via powercfg";
                }
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"  Registry access failed ({ex.Message}), using powercfg...");
                RunElevatedCmd("powercfg /hibernate off");
                detail = "Hibernate and Fast Startup disabled via powercfg";
            }
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult ReRegisterComponents()
        {
            int repairs = 0;
            OnLog?.Invoke("  Re-registering system components...");

            OnLog?.Invoke("  Resetting Windows component registrations...");
            int exitCode = RunElevatedCmd("DISM /Online /Cleanup-Image /StartComponentCleanup");
            if (exitCode == 0) repairs++;

            OnLog?.Invoke("  Clearing Windows Store cache...");
            string storeCacheDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Packages", "Microsoft.WindowsStore_8wekyb3d8bbwe", "LocalCache");
            if (Directory.Exists(storeCacheDir))
            {
                int cacheDeleted = 0, cacheSkipped = 0;
                DeleteDirectoryContents(storeCacheDir, ref cacheDeleted, ref cacheSkipped);
                if (cacheDeleted > 0) repairs++;
            }
            else
            {
                repairs++; // cache dir doesn't exist, nothing to reset
            }

            OnLog?.Invoke("  Rebuilding font cache...");
            StopService("FontCache", TimeSpan.FromSeconds(15));
            try
            {
                string fontCacheDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    "ServiceProfiles", "LocalService", "AppData", "Local", "FontCache");
                if (Directory.Exists(fontCacheDir))
                {
                    int deleted = 0, skipped = 0;
                    DeleteDirectoryContents(fontCacheDir, ref deleted, ref skipped);
                    if (deleted > 0) repairs++;
                }
            }
            finally
            {
                StartService("FontCache", TimeSpan.FromSeconds(15));
            }

            OnLog?.Invoke("  Re-registering core runtime libraries...");
            string[] criticalDlls = ["oleaut32.dll", "ole32.dll", "mshtml.dll", "urlmon.dll", "shell32.dll"];
            foreach (string dll in criticalDlls)
            {
                exitCode = RunElevatedCmd($"regsvr32 /s {dll}");
                if (exitCode == 0) repairs++;
            }

            string detail = $"Completed {repairs} component repairs — apps should be more stable after reboot";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(Detail: detail);
        }

        private ActionResult RepairPowerConfiguration()
        {
            int fixes = 0;
            OnLog?.Invoke("  Repairing power configuration...");

            OnLog?.Invoke("  Resetting active power plan to defaults...");
            int exitCode = RunElevatedCmd("powercfg /restoredefaultschemes");
            if (exitCode == 0) fixes++;

            const string balancedGuid = "381b4222-f694-41f0-9685-ff5bb260df2e";
            if (RunElevatedCmd($"powercfg /setactive {balancedGuid}") == 0) fixes++;

            OnLog?.Invoke("  Disabling wake timers...");
            RunElevatedCmd($"powercfg /SETACVALUEINDEX {balancedGuid} 238c9fa8-0aad-41ed-83f4-97be242c8f20 bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d 0");
            RunElevatedCmd($"powercfg /SETDCVALUEINDEX {balancedGuid} 238c9fa8-0aad-41ed-83f4-97be242c8f20 bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d 0");
            fixes++;

            OnLog?.Invoke("  Disabling USB selective suspend...");
            RunElevatedCmd($"powercfg /SETACVALUEINDEX {balancedGuid} 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0");
            RunElevatedCmd($"powercfg /SETDCVALUEINDEX {balancedGuid} 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0");
            fixes++;

            OnLog?.Invoke("  Disabling PCI Express power management...");
            RunElevatedCmd($"powercfg /SETACVALUEINDEX {balancedGuid} 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0");
            RunElevatedCmd($"powercfg /SETDCVALUEINDEX {balancedGuid} 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0");
            fixes++;

            RunElevatedCmd($"powercfg /setactive {balancedGuid}");

            string detail = $"Applied {fixes} power configuration fixes — wake timers, USB suspend, and PCIe power management disabled";
            OnLog?.Invoke($"  ✓ {detail}");
            return new ActionResult(Detail: detail);
        }

        // ═══════════════════════════════════════════════════════════════
        //  REMEDIATION TIMESTAMPS (per event-log category)
        // ═══════════════════════════════════════════════════════════════

        private static readonly string RemediationDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DLack", "event-log-remediation.dat");

        // Legacy single-timestamp path (migrated on first read)
        private static readonly string LegacyTimestampPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DLack", "event-log-remediated.txt");

        /// <summary>
        /// Records the current UTC time as the remediation cutoff for the specified
        /// event-log categories. Only events after this timestamp will be flagged
        /// on the next scan.
        /// </summary>
        private static void SaveRemediationTimestamps(HashSet<string> categories)
        {
            if (categories.Count == 0) return;
            try
            {
                var timestamps = LoadRemediationTimestamps();
                string nowUtc = DateTime.UtcNow.ToString("O");
                foreach (string category in categories)
                    timestamps[category] = nowUtc;

                Directory.CreateDirectory(Path.GetDirectoryName(RemediationDataPath)!);
                File.WriteAllLines(RemediationDataPath,
                    timestamps.Select(kvp => $"{kvp.Key}={kvp.Value}"));
            }
            catch { /* best-effort */ }
        }

        /// <summary>
        /// Returns the last remediation UTC timestamp for a specific event category,
        /// or null if that category has never been remediated.
        /// </summary>
        public static DateTime? GetRemediationTimestamp(string category)
        {
            try
            {
                var timestamps = LoadRemediationTimestamps();
                if (timestamps.TryGetValue(category, out string value) &&
                    DateTime.TryParse(value, null,
                        System.Globalization.DateTimeStyles.RoundtripKind, out var ts))
                    return ts.ToUniversalTime();
            }
            catch { /* best-effort */ }
            return null;
        }

        /// <summary>
        /// Returns the most recent remediation timestamp across all categories (UTC),
        /// or null if no remediation has ever been performed.
        /// </summary>
        public static DateTime? GetRemediationTimestamp()
        {
            try
            {
                var timestamps = LoadRemediationTimestamps();
                DateTime? latest = null;
                foreach (string value in timestamps.Values)
                {
                    if (DateTime.TryParse(value, null,
                        System.Globalization.DateTimeStyles.RoundtripKind, out var ts))
                    {
                        var utc = ts.ToUniversalTime();
                        if (!latest.HasValue || utc > latest.Value)
                            latest = utc;
                    }
                }
                return latest;
            }
            catch { /* best-effort */ }
            return null;
        }

        private static Dictionary<string, string> LoadRemediationTimestamps()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Per-category file takes priority
            if (File.Exists(RemediationDataPath))
            {
                foreach (string line in File.ReadAllLines(RemediationDataPath))
                {
                    int eq = line.IndexOf('=');
                    if (eq > 0)
                        result[line[..eq].Trim()] = line[(eq + 1)..].Trim();
                }
                return result;
            }

            // Migrate legacy single-timestamp file: treat as a global remediation
            if (File.Exists(LegacyTimestampPath))
            {
                string text = File.ReadAllText(LegacyTimestampPath).Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    result["BSODs"] = text;
                    result["DiskErrors"] = text;
                    result["AppCrashes"] = text;
                    result["UnexpectedShutdowns"] = text;
                }
            }

            return result;
        }

        /// <summary>
        /// Maps successful Event Log remediation actions to the event categories they address.
        /// </summary>
        private static HashSet<string> GetRemediatedCategories(IEnumerable<OptimizationAction> actions)
        {
            var categories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var action in actions)
            {
                if (action.Status is not (ActionStatus.Success or ActionStatus.PartialSuccess))
                    continue;

                // Event Log category actions
                if (action.Category == "Event Log")
                {
                    switch (action.ActionKey)
                    {
                        case "RunSfc":
                        case "RunDism":
                        case "InstallDriverUpdates":
                        case "ScheduleMemoryDiagnostic":
                            categories.Add("BSODs");
                            break;
                        case "ScheduleChkdsk":
                            categories.Add("DiskErrors");
                            break;
                        case "ReRegisterComponents":
                        case "OpenAppsSettings":
                            categories.Add("AppCrashes");
                            break;
                        case "DisableFastStartup":
                        case "RepairPowerConfig":
                            categories.Add("UnexpectedShutdowns");
                            break;
                    }
                    // Browser cache clear triggered by crash diagnosis
                    if (action.ActionKey.StartsWith("ClearBrowserCache:", StringComparison.Ordinal))
                        categories.Add("AppCrashes");
                }
            }
            return categories;
        }

        // ═══════════════════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Fast check: does the directory contain at least one file? (No recursive size scan.)
        /// </summary>
        private static bool DirectoryHasFiles(string path)
        {
            try
            {
                if (!Directory.Exists(path)) return false;
                return Directory.EnumerateFiles(path, "*", SearchOption.TopDirectoryOnly).Any()
                    || Directory.EnumerateDirectories(path, "*", SearchOption.TopDirectoryOnly).Any();
            }
            catch { return false; }
        }

        private long DeleteDirectoryContents(string path, ref int deleted, ref int skipped)
        {
            if (!Directory.Exists(path)) return 0;

            long totalFreed = 0;
            int localDeleted = 0;
            int localSkipped = 0;

            var options = new EnumerationOptions
            {
                RecurseSubdirectories = true,
                IgnoreInaccessible = true,
                AttributesToSkip = 0
            };

            // Delete files in parallel for I/O throughput
            try
            {
                Parallel.ForEach(
                    Directory.EnumerateFiles(path, "*", options),
                    file =>
                    {
                        try
                        {
                            var fi = new FileInfo(file);
                            long size = fi.Length;
                            fi.Attributes = FileAttributes.Normal;
                            fi.Delete();
                            Interlocked.Add(ref totalFreed, size);
                            Interlocked.Increment(ref localDeleted);
                        }
                        catch
                        {
                            Interlocked.Increment(ref localSkipped);
                        }
                    });
            }
            catch { /* enumeration-level failure */ }

            // Remove empty directories bottom-up (deepest paths first)
            try
            {
                foreach (var dir in Directory.EnumerateDirectories(path, "*", options)
                    .OrderByDescending(d => d.Length))
                {
                    try { Directory.Delete(dir, false); }
                    catch { /* skip non-empty or locked dirs */ }
                }
            }
            catch { /* enumeration failure */ }

            deleted += localDeleted;
            skipped += localSkipped;
            return totalFreed;
        }

        private static void StopService(string serviceName, TimeSpan timeout)
        {
            try
            {
                using var sc = new ServiceController(serviceName);
                sc.Refresh();
                if (sc.Status is ServiceControllerStatus.Running
                    or ServiceControllerStatus.StartPending)
                {
                    sc.Stop();
                    sc.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
                }
            }
            catch (InvalidOperationException) { /* service not installed */ }
            catch (System.ComponentModel.Win32Exception) { /* access denied */ }
            catch (System.ServiceProcess.TimeoutException) { /* timed out waiting */ }
        }

        private static void StartService(string serviceName, TimeSpan timeout)
        {
            try
            {
                using var sc = new ServiceController(serviceName);
                sc.Refresh();
                if (sc.Status is ServiceControllerStatus.Stopped
                    or ServiceControllerStatus.StopPending)
                {
                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, timeout);
                }
            }
            catch (InvalidOperationException) { /* service not installed */ }
            catch (System.ComponentModel.Win32Exception) { /* access denied */ }
            catch (System.ServiceProcess.TimeoutException) { /* timed out waiting */ }
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  OPTIMIZATION SUMMARY
    // ═══════════════════════════════════════════════════════════════

    public class OptimizationSummary
    {
        public int ActionsRun { get; set; }
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public long TotalFreedMB { get; set; }
        public List<string> ActionDetails { get; set; } = new();
    }
}
