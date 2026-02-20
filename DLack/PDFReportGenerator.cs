using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DLack
{
    public class PDFReportGenerator
    {
        // ── Brand Colors ─────────────────────────────────────────────
        private const string BluePrimary = "#9B4D2B";
        private const string GreenGood = "#1B7A3D";
        private const string RedBad = "#B91C1C";
        private const string AmberWarn = "#B4820A";
        private const string GrayMuted = "#6B7280";
        private const string GrayLight = "#F9FAFB";
        private const string GrayBorder = "#E5E7EB";
        private const string TextDark = "#111827";

        public void Generate(string filePath, DiagnosticResult result)
        {
            ArgumentNullException.ThrowIfNull(result);
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be empty.", nameof(filePath));

            string directory = System.IO.Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
                System.IO.Directory.CreateDirectory(directory);

            try
            {
                QuestPDF.Settings.License = LicenseType.Community;

                Document.Create(container =>
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        page.Margin(2, Unit.Centimetre);
                        page.DefaultTextStyle(x => x.FontSize(10).FontColor(TextDark));

                        page.Header().Element(c => ComposeHeader(c, result));
                        page.Content().Element(c => ComposeContent(c, result));
                        page.Footer().Element(ComposeFooter);
                    });
                })
                .GeneratePdf(filePath);
            }
            catch (UnauthorizedAccessException)
            {
                throw new InvalidOperationException(
                    $"Access denied when writing PDF to: {filePath}\n\n" +
                    "Ensure you have write permissions to this directory.");
            }
            catch (System.IO.IOException ex)
            {
                throw new InvalidOperationException(
                    $"IO error when writing PDF to: {filePath}\n\n" +
                    $"Details: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Unexpected error while generating PDF:\n\n{ex.Message}", ex);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  HEADER
        // ═══════════════════════════════════════════════════════════════

        private void ComposeHeader(IContainer container, DiagnosticResult result)
        {
            container.Column(col =>
            {
                col.Item().Height(4).Background(BluePrimary);

                col.Item().PaddingTop(12).PaddingBottom(8).Row(row =>
                {
                    row.RelativeItem().Column(left =>
                    {
                        var logoBytes = LoadLogoBytes();
                        left.Item().Row(titleRow =>
                        {
                            if (logoBytes != null)
                            {
                                titleRow.ConstantItem(130).Height(26).Image(logoBytes).FitArea();
                                titleRow.ConstantItem(10);
                            }
                            titleRow.RelativeItem().AlignBottom()
                                .Text("IT Diagnostic Report")
                                .FontSize(18).Bold().FontColor(BluePrimary);
                        });
                        left.Item().PaddingTop(2)
                            .Text("Comprehensive System Health Assessment")
                            .FontSize(10).FontColor(GrayMuted);
                    });

                    row.ConstantItem(180).AlignRight().Column(right =>
                    {
                        right.Item().Text($"{result.Timestamp:MMMM dd, yyyy}").FontSize(10).FontColor(TextDark);
                        right.Item().Text($"{result.Timestamp:hh:mm tt}").FontSize(9).FontColor(GrayMuted);
                    });
                });

                col.Item().Height(1).Background(GrayBorder);
            });
        }

        private static byte[] LoadLogoBytes()
        {
            try
            {
                var stream = Application.GetResourceStream(
                    new Uri("pack://application:,,,/Assets/pillsbury-logo-color.png"));
                if (stream != null)
                {
                    using var ms = new System.IO.MemoryStream();
                    stream.Stream.CopyTo(ms);
                    return ms.ToArray();
                }
            }
            catch { }
            try
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                string logoPath = System.IO.Path.Combine(exeDir, "Assets", "pillsbury-logo-color.png");
                if (System.IO.File.Exists(logoPath))
                    return System.IO.File.ReadAllBytes(logoPath);
            }
            catch { }
            return null;
        }

        // ═══════════════════════════════════════════════════════════════
        //  CONTENT
        // ═══════════════════════════════════════════════════════════════

        private void ComposeContent(IContainer container, DiagnosticResult result)
        {
            container.PaddingVertical(10).Column(col =>
            {
                col.Spacing(14);

                ComposeSummary(col, result);
                ComposeSystemOverview(col, result);
                ComposeCpuDiagnostics(col, result);
                ComposeRamDiagnostics(col, result);
                ComposeGpuDiagnostics(col, result);
                ComposeDiskDiagnostics(col, result);
                ComposeBattery(col, result);
                ComposeStartup(col, result);
                ComposeVisualSettings(col, result);
                ComposeAntivirus(col, result);
                ComposeWindowsUpdate(col, result);
                ComposeNetwork(col, result);
                ComposeNetworkDrives(col, result);
                ComposeOutlook(col, result);
                ComposeBrowsers(col, result);
                ComposeUserProfile(col, result);
                ComposeOffice(col, result);
                ComposeInstalledSoftware(col, result);
                ComposeEventLog(col, result);
                ComposeRecommendations(col, result);
                ComposeOptimizationPlan(col, result);
            });
        }

        // ── Summary ──────────────────────────────────────────────────

        private void ComposeSummary(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Executive Summary"));

            int critical = r.FlaggedIssues.Count(i => i.Severity == Severity.Critical);
            int warning = r.FlaggedIssues.Count(i => i.Severity == Severity.Warning);
            int info = r.FlaggedIssues.Count(i => i.Severity == Severity.Info);

            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(6);

                string scoreColor = r.HealthScore >= 70 ? GreenGood :
                                    r.HealthScore >= 50 ? AmberWarn : RedBad;
                string scoreLabel = r.HealthScore >= 85 ? "Excellent" :
                                    r.HealthScore >= 70 ? "Good" :
                                    r.HealthScore >= 50 ? "Fair" : "Poor";

                inner.Item().Row(scoreRow =>
                {
                    scoreRow.AutoItem().Text($"Health Score: {r.HealthScore}/100")
                        .FontSize(16).Bold().FontColor(scoreColor);
                    scoreRow.AutoItem().PaddingLeft(8).AlignBottom()
                        .Text($"({scoreLabel})")
                        .FontSize(10).FontColor(scoreColor);
                });

                inner.Item().Text($"Diagnostic scan completed at {r.Timestamp:h:mm:ss tt} on {r.Timestamp:yyyy-MM-dd} (took {r.ScanDurationSeconds}s)")
                    .FontSize(10).FontColor(TextDark);
                if (!string.IsNullOrEmpty(r.ScannedUser))
                    inner.Item().Text($"Scanned as: {r.ScannedUser}")
                        .FontSize(9).FontColor(GrayMuted);
                inner.Item().Text($"Total issues flagged: {r.FlaggedIssues.Count}")
                    .FontSize(10).Bold().FontColor(TextDark);

                inner.Item().PaddingTop(4).Row(row =>
                {
                    if (critical > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(RedBad);
                            badge.AutoItem().PaddingLeft(4).Text($"{critical} Critical").FontSize(9).Bold().FontColor(RedBad);
                        });
                    }
                    if (warning > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(AmberWarn);
                            badge.AutoItem().PaddingLeft(4).Text($"{warning} Warning").FontSize(9).Bold().FontColor(AmberWarn);
                        });
                    }
                    if (info > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(BluePrimary);
                            badge.AutoItem().PaddingLeft(4).Text($"{info} Informational").FontSize(9).FontColor(GrayMuted);
                        });
                    }
                });

                if (r.FlaggedIssues.Count == 0)
                    inner.Item().PaddingTop(4).Text("No issues detected \u2014 system appears healthy")
                        .FontSize(10).Bold().FontColor(GreenGood);
            });
        }

        // ── System Overview ──────────────────────────────────────────

        private void ComposeSystemOverview(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "System Overview"));
            var so = r.SystemOverview;

            col.Item().Background(GrayLight).Padding(12).Row(row =>
            {
                row.RelativeItem().Column(left =>
                {
                    InfoRow(left, "Computer", so.ComputerName);
                    InfoRow(left, "Manufacturer", so.Manufacturer);
                    InfoRow(left, "Model", so.Model);
                    InfoRow(left, "Serial", so.SerialNumber);
                });
                row.RelativeItem().Column(right =>
                {
                    InfoRow(right, "Windows", so.WindowsVersion);
                    InfoRow(right, "Build", so.WindowsBuild);
                    InfoRow(right, "CPU", so.CpuModel);
                    InfoRow(right, "Clock Speed", so.CpuClockSpeed);
                    InfoRow(right, "RAM", FormatMB(so.TotalRamMB));
                    InfoRow(right, "Uptime", FormatUptime(so.Uptime) + (so.UptimeFlagged ? " ⚠" : ""));
                });
            });
        }

        // ── CPU Diagnostics ──────────────────────────────────────────

        private void ComposeCpuDiagnostics(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "CPU Diagnostics"));

            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(4);
                string cpuColor = r.Cpu.CpuLoadFlagged ? (r.Cpu.CpuLoadPercent > 50 ? RedBad : AmberWarn) : GreenGood;
                inner.Item().Text($"CPU Load: {r.Cpu.CpuLoadPercent}%")
                    .FontSize(10).Bold().FontColor(cpuColor);

                // Thermal data
                if (r.Cpu.CpuTemperatureC > 0)
                {
                    string tempColor = r.Cpu.TemperatureFlagged ? RedBad : GreenGood;
                    inner.Item().Text($"CPU Temperature: {r.Cpu.CpuTemperatureC}°C")
                        .FontSize(9).FontColor(tempColor);
                }
                else
                {
                    InfoRow(inner, "CPU Temperature", "Not available");
                }
                InfoRow(inner, "Throttling", r.Cpu.IsThrottling ? "Yes ⚠" : "No");
                InfoRow(inner, "Fan Status", r.Cpu.FanStatus);
            });

            if (r.Cpu.TopCpuProcesses.Count > 0)
            {
                col.Item().Table(table =>
                {
                    table.ColumnsDefinition(cd =>
                    {
                        cd.RelativeColumn(3f);
                        cd.RelativeColumn(1.5f);
                        cd.RelativeColumn(1.5f);
                    });
                    table.Header(h =>
                    {
                        foreach (var label in new[] { "Process", "PID", "Memory" })
                            h.Cell().Background(BluePrimary).Padding(6)
                                .Text(label).FontSize(9).Bold().FontColor(Colors.White);
                    });
                    int i = 0;
                    foreach (var p in r.Cpu.TopCpuProcesses.Take(10))
                    {
                        bool alt = i++ % 2 == 1;
                        table.Cell().Element(c => DataCell(c, alt)).Text(p.Name).FontSize(9);
                        table.Cell().Element(c => DataCell(c, alt)).Text(p.Id.ToString()).FontSize(9);
                        table.Cell().Element(c => DataCell(c, alt)).AlignRight()
                            .Text(FormatMB(p.MemoryMB)).FontSize(9);
                    }
                });
            }
        }

        // ── RAM Diagnostics

        private void ComposeRamDiagnostics(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "RAM Diagnostics"));
            var ram = r.Ram;

            string ramColor = ram.UsageFlagged ? (ram.PercentUsed > 95 ? RedBad : AmberWarn) : GreenGood;

            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(4);
                inner.Item().Text($"RAM Usage: {ram.PercentUsed}% ({FormatMB(ram.UsedMB)} / {FormatMB(ram.TotalMB)})")
                    .FontSize(10).Bold().FontColor(ramColor);
                inner.Item().Text($"Available: {FormatMB(ram.AvailableMB)}").FontSize(9).FontColor(TextDark);
                if (ram.InsufficientRam)
                    inner.Item().Text($"⚠ Total RAM ({FormatMB(ram.TotalMB)}) may be insufficient")
                        .FontSize(9).FontColor(AmberWarn);
                inner.Item().Text($"Memory Diagnostic: {ram.MemoryDiagnosticStatus}")
                    .FontSize(9).FontColor(GrayMuted);
            });

            if (ram.TopRamProcesses.Count > 0)
            {
                col.Item().Table(table =>
                {
                    table.ColumnsDefinition(cd =>
                    {
                        cd.RelativeColumn(3f);
                        cd.RelativeColumn(1.5f);
                        cd.RelativeColumn(1.5f);
                    });
                    table.Header(h =>
                    {
                        foreach (var label in new[] { "Process", "PID", "Memory" })
                            h.Cell().Background(BluePrimary).Padding(6)
                                .Text(label).FontSize(9).Bold().FontColor(Colors.White);
                    });
                    int i = 0;
                    foreach (var p in ram.TopRamProcesses.Take(10))
                    {
                        bool alt = i++ % 2 == 1;
                        table.Cell().Element(c => DataCell(c, alt)).Text(p.Name).FontSize(9);
                        table.Cell().Element(c => DataCell(c, alt)).Text(p.Id.ToString()).FontSize(9);
                        table.Cell().Element(c => DataCell(c, alt)).AlignRight()
                            .Text(FormatMB(p.MemoryMB)).FontSize(9);
                    }
                });
            }
        }

        // ── Disk Diagnostics

        private void ComposeDiskDiagnostics(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Disk Diagnostics"));

            // Drives table
            col.Item().Table(table =>
            {
                table.ColumnsDefinition(cd =>
                {
                    cd.RelativeColumn(1f);   // Drive
                    cd.RelativeColumn(1.5f); // Total
                    cd.RelativeColumn(1.5f); // Used
                    cd.RelativeColumn(1.5f); // Free
                    cd.RelativeColumn(1f);   // % Used
                    cd.RelativeColumn(1f);   // Type
                    cd.RelativeColumn(1f);   // Health
                });
                table.Header(h =>
                {
                    foreach (var label in new[] { "Drive", "Total", "Used", "Free", "% Used", "Type", "Health" })
                        h.Cell().Background(BluePrimary).Padding(5)
                            .Text(label).FontSize(8).Bold().FontColor(Colors.White);
                });
                int i = 0;
                foreach (var d in r.Disk.Drives)
                {
                    bool alt = i++ % 2 == 1;
                    string usedColor = d.UsageFlagged ? (d.PercentUsed > 95 ? RedBad : AmberWarn) : TextDark;
                    table.Cell().Element(c => DataCell(c, alt)).Text(d.DriveLetter).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(FormatMB(d.TotalMB)).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(FormatMB(d.UsedMB)).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(FormatMB(d.FreeMB)).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text($"{d.PercentUsed}%").FontSize(9).FontColor(usedColor);
                    table.Cell().Element(c => DataCell(c, alt)).Text(d.DriveType).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(d.HealthStatus).FontSize(9);
                }
            });

            // Folder sizes
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                InfoRow(inner, "Windows Temp", FormatMB(r.Disk.WindowsTempMB));
                InfoRow(inner, "User Temp", FormatMB(r.Disk.UserTempMB));
                InfoRow(inner, "SW Distribution", FormatMB(r.Disk.SoftwareDistributionMB));
                InfoRow(inner, "Prefetch", FormatMB(r.Disk.PrefetchMB));
                InfoRow(inner, "Recycle Bin", FormatMB(r.Disk.RecycleBinMB));
                if (r.Disk.UpgradeLogsMB > 0)
                    InfoRow(inner, "Upgrade Logs", FormatMB(r.Disk.UpgradeLogsMB));
                if (r.Disk.WindowsOldExists)
                    InfoRow(inner, "Windows.old", $"{FormatMB(r.Disk.WindowsOldMB)} ⚠");
                if (r.Disk.DiskActivityFlagged)
                    inner.Item().Text($"⚠ Disk activity: {r.Disk.DiskActivityPercent}% (above 50% at idle)")
                        .FontSize(9).FontColor(AmberWarn);
            });
        }

        // ── GPU Diagnostics ──────────────────────────────────────────

        private void ComposeGpuDiagnostics(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "GPU & Display Diagnostics"));

            // Show detected displays
            if (r.Gpu.Displays.Count > 0)
            {
                col.Item().Background(GrayLight).Padding(12).Column(inner =>
                {
                    inner.Spacing(3);
                    inner.Item().Text("Connected Displays")
                        .FontSize(10).Bold().FontColor(TextDark);
                    foreach (var display in r.Gpu.Displays)
                    {
                        string label = display.IsPrimary ? "Primary" : "Secondary";
                        InfoRow(inner, $"{label} Display", display.MonitorName);
                        InfoRow(inner, "  Resolution", $"{display.Resolution} @ {display.RefreshRateText}");
                    }
                });
            }

            // Show all detected GPUs
            if (r.Gpu.AllGpus.Count > 0)
            {
                foreach (var gpuInfo in r.Gpu.AllGpus)
                {
                    col.Item().Background(GrayLight).Padding(12).Column(inner =>
                    {
                        inner.Spacing(3);
                        string label = gpuInfo.IsPrimary ? "Primary" : "Secondary";
                        inner.Item().Text($"{gpuInfo.Name} ({label})")
                            .FontSize(10).Bold().FontColor(TextDark);
                        InfoRow(inner, "Driver Version", gpuInfo.DriverVersion);
                        InfoRow(inner, "Driver Date", gpuInfo.DriverDate + (gpuInfo.DriverOutdated ? " ⚠ outdated" : ""));
                        InfoRow(inner, "Status", gpuInfo.AdapterStatus);
                        if (gpuInfo.DedicatedVideoMemoryMB > 0)
                            InfoRow(inner, "Dedicated VRAM", FormatMB(gpuInfo.DedicatedVideoMemoryMB));

                        if (gpuInfo.IsPrimary)
                        {
                            if (r.Gpu.GpuTemperatureC > 0)
                            {
                                string tempColor = r.Gpu.TemperatureFlagged ? RedBad : GreenGood;
                                inner.Item().Text($"GPU Temperature: {r.Gpu.GpuTemperatureC}°C")
                                    .FontSize(9).FontColor(tempColor);
                            }
                            else
                            {
                                InfoRow(inner, "GPU Temperature", "Not available");
                            }

                            if (r.Gpu.GpuUsagePercent > 0)
                            {
                                string usageColor = r.Gpu.UsageFlagged ? RedBad : GreenGood;
                                inner.Item().Text($"GPU Usage: {r.Gpu.GpuUsagePercent}%")
                                    .FontSize(9).FontColor(usageColor);
                            }
                        }

                        if (gpuInfo.DriverOutdated)
                            inner.Item().Text("⚠ GPU driver is over 1 year old — consider updating")
                                .FontSize(9).FontColor(AmberWarn);
                    });
                }
            }
            else
            {
                // Fallback: single GPU
                col.Item().Background(GrayLight).Padding(12).Column(inner =>
                {
                    inner.Spacing(3);
                    inner.Item().Text(r.Gpu.GpuName)
                        .FontSize(10).Bold().FontColor(TextDark);
                    InfoRow(inner, "Driver Version", r.Gpu.DriverVersion);
                    InfoRow(inner, "Driver Date", r.Gpu.DriverDate + (r.Gpu.DriverOutdated ? " ⚠ outdated" : ""));
                    InfoRow(inner, "Status", r.Gpu.AdapterStatus);
                    if (r.Gpu.DedicatedVideoMemoryMB > 0)
                        InfoRow(inner, "Dedicated VRAM", FormatMB(r.Gpu.DedicatedVideoMemoryMB));
                    InfoRow(inner, "Resolution", r.Gpu.Resolution);
                    InfoRow(inner, "Refresh Rate", r.Gpu.RefreshRate);

                    if (r.Gpu.GpuTemperatureC > 0)
                    {
                        string tempColor = r.Gpu.TemperatureFlagged ? RedBad : GreenGood;
                        inner.Item().Text($"GPU Temperature: {r.Gpu.GpuTemperatureC}°C")
                            .FontSize(9).FontColor(tempColor);
                    }
                    else
                    {
                        InfoRow(inner, "GPU Temperature", "Not available");
                    }

                    if (r.Gpu.GpuUsagePercent > 0)
                    {
                        string usageColor = r.Gpu.UsageFlagged ? RedBad : GreenGood;
                        inner.Item().Text($"GPU Usage: {r.Gpu.GpuUsagePercent}%")
                            .FontSize(9).FontColor(usageColor);
                    }

                    foreach (var adapter in r.Gpu.AdditionalAdapters)
                        inner.Item().Text($"+ {adapter}").FontSize(8).FontColor(GrayMuted);

                    if (r.Gpu.DriverOutdated)
                        inner.Item().Text("⚠ GPU driver is over 1 year old — consider updating")
                            .FontSize(9).FontColor(AmberWarn);
                });
            }
        }

        // ── Battery ────────────────────────────────────────────────

        private void ComposeBattery(ColumnDescriptor col, DiagnosticResult r)
        {
            if (!r.Battery.HasBattery) return;

            col.Item().Element(c => SectionHeader(c, "Battery Health"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                string healthColor = r.Battery.HealthFlagged ? RedBad : GreenGood;
                inner.Item().Text($"Battery Health: {r.Battery.HealthPercent}%")
                    .FontSize(10).Bold().FontColor(healthColor);
                InfoRow(inner, "Design Capacity", $"{r.Battery.DesignCapacityMWh} mWh");
                InfoRow(inner, "Full Charge", $"{r.Battery.FullChargeCapacityMWh} mWh");
                InfoRow(inner, "Power Source", r.Battery.PowerSource);
                InfoRow(inner, "Power Plan", r.Battery.PowerPlan);
            });
        }

        // ── Startup ──────────────────────────────────────────────────

        private void ComposeStartup(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Startup Programs"));

            string countColor = r.Startup.TooManyFlagged ? AmberWarn : GreenGood;
            col.Item().Background(GrayLight).Padding(8).Column(inner =>
            {
                inner.Item().Text($"Enabled: {r.Startup.EnabledCount} / Total: {r.Startup.Entries.Count}")
                    .FontSize(10).Bold().FontColor(countColor);
            });

            if (r.Startup.Entries.Count > 0)
            {
                col.Item().Table(table =>
                {
                    table.ColumnsDefinition(cd =>
                    {
                        cd.RelativeColumn(4f);
                        cd.RelativeColumn(1f);
                    });
                    table.Header(h =>
                    {
                        h.Cell().Background(BluePrimary).Padding(5).Text("Name").FontSize(8).Bold().FontColor(Colors.White);
                        h.Cell().Background(BluePrimary).Padding(5).Text("Status").FontSize(8).Bold().FontColor(Colors.White);
                    });
                    int i = 0;
                    foreach (var entry in r.Startup.Entries.Take(20))
                    {
                        bool alt = i++ % 2 == 1;
                        table.Cell().Element(c => DataCell(c, alt)).Text(entry.Name).FontSize(8);
                        table.Cell().Element(c => DataCell(c, alt))
                            .Text(entry.Enabled ? "Enabled" : "Disabled")
                            .FontSize(8).FontColor(entry.Enabled ? TextDark : GrayMuted);
                    }
                });
            }
        }

        // ── Visual Settings ──────────────────────────────────────────

        private void ComposeVisualSettings(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Visual & Display Settings"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                InfoRow(inner, "Visual Effects", r.VisualSettings.VisualEffectsSetting);
                InfoRow(inner, "Transparency", r.VisualSettings.TransparencyEnabled ? "Enabled" : "Disabled");
                InfoRow(inner, "Animations", r.VisualSettings.AnimationsEnabled ? "Enabled" : "Disabled");
            });
        }

        // ── Antivirus ────────────────────────────────────────────────

        private void ComposeAntivirus(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Antivirus & Security"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                InfoRow(inner, "Antivirus", r.Antivirus.AntivirusName);
                InfoRow(inner, "Full Scan Running", r.Antivirus.FullScanRunning ? "Yes" : "No");
                InfoRow(inner, "BitLocker", r.Antivirus.BitLockerStatus);
            });
        }

        // ── Windows Update ───────────────────────────────────────────

        private void ComposeWindowsUpdate(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Windows Update"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                InfoRow(inner, "Pending Updates", r.WindowsUpdate.PendingUpdates);

                // Show service status with start type for clarity
                string wuService = r.WindowsUpdate.UpdateServiceStatus;
                string wuStartType = r.WindowsUpdate.ServiceStartType;
                if (!string.IsNullOrEmpty(wuStartType) && wuStartType != "Unknown")
                    wuService = $"{wuService} ({wuStartType})";
                InfoRow(inner, "Service Status", wuService);

                InfoRow(inner, "Cache Size", FormatMB(r.WindowsUpdate.UpdateCacheMB));
            });
        }

        // ── Network ──────────────────────────────────────────────────

        private void ComposeNetwork(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Network Diagnostics"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                InfoRow(inner, "Connection", r.Network.ConnectionType);
                if (r.Network.ConnectionType == "WiFi")
                {
                    InfoRow(inner, "WiFi Band", r.Network.WifiBand);
                    InfoRow(inner, "Signal", r.Network.WifiSignalStrength);
                    InfoRow(inner, "Link Speed", r.Network.LinkSpeed);
                }
                InfoRow(inner, "Adapter Speed", r.Network.AdapterSpeed);

                // DNS Response with color
                string dnsColor = RateDnsColor(r.Network.DnsResponseTime);
                inner.Item().Text(t =>
                {
                    t.Span("DNS Response: ").FontSize(9).FontColor(TextDark);
                    t.Span(r.Network.DnsResponseTime).FontSize(9).Bold().FontColor(dnsColor);
                });

                // Ping Latency with color
                string pingColor = RatePingColor(r.Network.PingLatency);
                inner.Item().Text(t =>
                {
                    t.Span("Ping Latency: ").FontSize(9).FontColor(TextDark);
                    t.Span(r.Network.PingLatency).FontSize(9).Bold().FontColor(pingColor);
                });
                InfoRow(inner, "VPN Active", r.Network.VpnActive ? $"Yes — {r.Network.VpnClient}" : "No");

                if (r.Network.DownloadSpeedMbps > 0 || r.Network.UploadSpeedMbps > 0)
                {
                    string dlColor = r.Network.DownloadSpeedMbps < 10 ? RedBad :
                                     r.Network.DownloadSpeedMbps < 50 ? AmberWarn : GreenGood;
                    string ulColor = r.Network.UploadSpeedMbps < 5 ? RedBad :
                                     r.Network.UploadSpeedMbps < 20 ? AmberWarn : GreenGood;

                    inner.Item().Text($"Download Speed: {r.Network.DownloadSpeedMbps:F2} Mbps")
                        .FontSize(9).FontColor(dlColor);
                    inner.Item().Text($"Upload Speed: {r.Network.UploadSpeedMbps:F2} Mbps")
                        .FontSize(9).FontColor(ulColor);
                }
                else if (!string.IsNullOrEmpty(r.Network.SpeedTestError))
                {
                    InfoRow(inner, "Speed Test", $"Failed — {r.Network.SpeedTestError}");
                }
            });
        }

        // ── Network Drives ───────────────────────────────────────────

        private void ComposeNetworkDrives(ColumnDescriptor col, DiagnosticResult r)
        {
            if (r.NetworkDrives.Drives.Count == 0) return;

            col.Item().Element(c => SectionHeader(c, "Network Drive Health"));

            col.Item().Table(table =>
            {
                table.ColumnsDefinition(cd =>
                {
                    cd.RelativeColumn(1f);   // Drive
                    cd.RelativeColumn(2.5f); // UNC Path
                    cd.RelativeColumn(1f);   // Status
                    cd.RelativeColumn(1f);   // Latency
                    cd.RelativeColumn(1f);   // % Used
                    cd.RelativeColumn(1.5f); // Free Space
                });
                table.Header(h =>
                {
                    foreach (var label in new[] { "Drive", "Path", "Status", "Latency", "% Used", "Free" })
                        h.Cell().Background(BluePrimary).Padding(5)
                            .Text(label).FontSize(8).Bold().FontColor(Colors.White);
                });
                int i = 0;
                foreach (var nd in r.NetworkDrives.Drives)
                {
                    bool alt = i++ % 2 == 1;
                    string statusColor = !nd.IsAccessible ? RedBad : nd.LatencyFlagged ? AmberWarn : GreenGood;
                    string usedColor = nd.SpaceFlagged ? AmberWarn : TextDark;

                    table.Cell().Element(c => DataCell(c, alt)).Text(nd.DriveLetter).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(nd.UncPath).FontSize(8);
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(nd.IsAccessible ? "Online" : "Offline ⚠").FontSize(9).FontColor(statusColor);
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(nd.IsAccessible ? $"{nd.LatencyMs} ms" : "—").FontSize(9)
                        .FontColor(nd.LatencyFlagged ? AmberWarn : TextDark);
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(nd.IsAccessible ? $"{nd.PercentUsed}%" : "—").FontSize(9).FontColor(usedColor);
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(nd.IsAccessible ? FormatMB(nd.FreeMB) : "—").FontSize(9);
                }
            });

            // Offline files info
            var offlineDrives = r.NetworkDrives.Drives.Where(d => d.OfflineFilesEnabled).ToList();
            if (offlineDrives.Count > 0)
            {
                col.Item().Background(GrayLight).Padding(8).Column(inner =>
                {
                    inner.Spacing(2);
                    inner.Item().Text("Offline Files Sync").FontSize(9).Bold().FontColor(TextDark);
                    foreach (var nd in offlineDrives)
                        InfoRow(inner, $"  {nd.DriveLetter}", $"Sync: {nd.SyncStatus}");
                });
            }
        }

        // ── Outlook ──────────────────────────────────────────────────

        private void ComposeOutlook(ColumnDescriptor col, DiagnosticResult r)
        {
            if (!r.Outlook.OutlookInstalled) return;

            col.Item().Element(c => SectionHeader(c, "Outlook / Email"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                foreach (var df in r.Outlook.DataFiles)
                {
                    string sizeColor = df.SizeFlagged ? AmberWarn : TextDark;
                    inner.Item().Text($"{System.IO.Path.GetFileName(df.Path)}: {FormatMB(df.SizeMB)}" +
                        (df.SizeFlagged ? " ⚠ (above 5 GB)" : ""))
                        .FontSize(9).FontColor(sizeColor);
                }
                InfoRow(inner, "Add-ins", $"{r.Outlook.AddInCount}");
            });
        }

        // ── Browsers ─────────────────────────────────────────────────

        private void ComposeBrowsers(ColumnDescriptor col, DiagnosticResult r)
        {
            if (r.Browser.Browsers.Count == 0) return;

            col.Item().Element(c => SectionHeader(c, "Browser Check"));
            col.Item().Table(table =>
            {
                table.ColumnsDefinition(cd =>
                {
                    cd.RelativeColumn(2f);
                    cd.RelativeColumn(1f);
                    cd.RelativeColumn(1.5f);
                    cd.RelativeColumn(1.5f);
                });
                table.Header(h =>
                {
                    foreach (var label in new[] { "Browser", "Open Tabs", "Cache", "Extensions" })
                        h.Cell().Background(BluePrimary).Padding(5)
                            .Text(label).FontSize(8).Bold().FontColor(Colors.White);
                });
                int i = 0;
                foreach (var b in r.Browser.Browsers)
                {
                    bool alt = i++ % 2 == 1;
                    table.Cell().Element(c => DataCell(c, alt)).Text(b.Name).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(b.OpenTabs.ToString()).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(FormatMB(b.CacheSizeMB)).FontSize(9);
                    table.Cell().Element(c => DataCell(c, alt)).Text(b.ExtensionCount.ToString()).FontSize(9);
                }
            });
        }

        // ── User Profile ─────────────────────────────────────────────

        private void ComposeUserProfile(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "User Profile"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                InfoRow(inner, "Profile Size", FormatMB(r.UserProfile.ProfileSizeMB));
                InfoRow(inner, "Profile Age", r.UserProfile.ProfileAge);
                InfoRow(inner, "Desktop Items", $"{r.UserProfile.DesktopItemCount}" +
                    (r.UserProfile.DesktopItemsFlagged ? " ⚠ (excessive)" : ""));
                if (r.UserProfile.CorruptionDetected)
                    inner.Item().Text("⚠ Profile corruption detected").FontSize(9).Bold().FontColor(RedBad);
            });
        }

        // ── Office ───────────────────────────────────────────────────

        private void ComposeOffice(ColumnDescriptor col, DiagnosticResult r)
        {
            if (!r.Office.OfficeInstalled) return;

            col.Item().Element(c => SectionHeader(c, "Office Diagnostics"));
            col.Item().Background(GrayLight).Padding(12).Column(inner =>
            {
                inner.Spacing(3);
                InfoRow(inner, "Version", r.Office.OfficeVersion);
                InfoRow(inner, "Repair Needed", r.Office.RepairNeeded ? "Yes ⚠" : "No");
            });
        }

        // ── Installed Software ───────────────────────────────────────

        private void ComposeInstalledSoftware(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Installed Software"));

            col.Item().Background(GrayLight).Padding(8).Column(summary =>
            {
                summary.Spacing(3);
                summary.Item().Text($"Total installed applications: {r.InstalledSoftware.TotalCount}")
                    .FontSize(10).Bold().FontColor(TextDark);
                if (r.InstalledSoftware.EOLApps.Count > 0)
                    summary.Item().Text($"⚠ {r.InstalledSoftware.EOLApps.Count} end-of-life application(s): {string.Join(", ", r.InstalledSoftware.EOLApps.Select(a => a.Name))}")
                        .FontSize(9).FontColor(AmberWarn);
                if (r.InstalledSoftware.BloatwareApps.Count > 0)
                    summary.Item().Text($"ℹ {r.InstalledSoftware.BloatwareApps.Count} potential bloatware: {string.Join(", ", r.InstalledSoftware.BloatwareApps.Select(a => a.Name))}")
                        .FontSize(9).FontColor(GrayMuted);
                if (r.InstalledSoftware.RuntimeApps.Count > 0)
                {
                    summary.Item().Text($"ℹ {r.InstalledSoftware.RuntimeApps.Count} Visual C++ Redistributables installed:")
                        .FontSize(9).FontColor(GrayMuted);
                    foreach (var rt in r.InstalledSoftware.RuntimeApps)
                        summary.Item().PaddingLeft(12).Text($"• {rt.Name}  ({rt.Version})")
                            .FontSize(8).FontColor(GrayMuted);
                }
            });

            if (r.InstalledSoftware.Applications.Count > 0)
            {
                const int maxApps = 200;
                var appsToShow = r.InstalledSoftware.Applications.Take(maxApps).ToList();
                int remaining = r.InstalledSoftware.Applications.Count - appsToShow.Count;

                col.Item().Table(table =>
                {
                    table.ColumnsDefinition(cd =>
                    {
                        cd.RelativeColumn(3f);
                        cd.RelativeColumn(1.5f);
                        cd.RelativeColumn(2f);
                    });
                    table.Header(h =>
                    {
                        foreach (var label in new[] { "Application", "Version", "Publisher" })
                            h.Cell().Background(BluePrimary).Padding(5)
                                .Text(label).FontSize(8).Bold().FontColor(Colors.White);
                    });
                    int i = 0;
                    foreach (var app in appsToShow)
                    {
                        bool alt = i++ % 2 == 1;
                        table.Cell().Element(c => DataCell(c, alt)).Text(app.Name).FontSize(8);
                        table.Cell().Element(c => DataCell(c, alt)).Text(app.Version).FontSize(8);
                        table.Cell().Element(c => DataCell(c, alt)).Text(app.Publisher).FontSize(8);
                    }
                });

                if (remaining > 0)
                    col.Item().PaddingTop(4)
                        .Text($"+ {remaining} more application(s) not shown")
                        .FontSize(8).Italic().FontColor(GrayMuted);
            }
        }

        // ── Event Log Analysis ──────────────────────────────────────

        private void ComposeEventLog(ColumnDescriptor col, DiagnosticResult r)
        {
            col.Item().Element(c => SectionHeader(c, "Event Log Analysis (30 Days)"));

            var el = r.EventLog;
            bool hasEvents = el.TotalEventCount > 0;
            DateTime? lastRemediation = Optimizer.GetRemediationTimestamp();

            // Summary counts
            col.Item().Background(GrayLight).Padding(8).Column(summary =>
            {
                summary.Spacing(3);

                string bsodColor = el.BSODs.Count > 0 ? RedBad : GreenGood;
                string shutdownColor = el.UnexpectedShutdowns.Count >= 3 ? RedBad :
                    el.UnexpectedShutdowns.Count > 0 ? AmberWarn : GreenGood;
                string diskColor = el.DiskErrors.Count >= 5 ? RedBad :
                    el.DiskErrors.Count > 0 ? AmberWarn : GreenGood;
                string crashColor = el.AppCrashes.Count >= 5 ? AmberWarn : GreenGood;

                summary.Item().Text($"BSODs: {el.BSODs.Count}")
                    .FontSize(10).Bold().FontColor(bsodColor);
                summary.Item().Text($"Unexpected Shutdowns: {el.UnexpectedShutdowns.Count}")
                    .FontSize(10).Bold().FontColor(shutdownColor);
                summary.Item().Text($"Disk Errors: {el.DiskErrors.Count}")
                    .FontSize(10).Bold().FontColor(diskColor);
                summary.Item().Text($"App Crashes: {el.AppCrashes.Count}")
                    .FontSize(10).Bold().FontColor(crashColor);

                if (!hasEvents && lastRemediation.HasValue)
                    summary.Item().PaddingTop(4).Text(
                        $"✓ All clear — remediated on {lastRemediation.Value.ToLocalTime():MMM dd, yyyy} at {lastRemediation.Value.ToLocalTime():h:mm tt}")
                        .FontSize(9).FontColor(GreenGood);
                else if (!hasEvents)
                    summary.Item().PaddingTop(4).Text("✓ No issues found")
                        .FontSize(9).FontColor(GreenGood);
            });

            // Recent entries table
            if (hasEvents)
            {
                var allEntries = new List<(string Category, EventLogEntry Entry)>();
                foreach (var e in el.BSODs) allEntries.Add(("BSOD", e));
                foreach (var e in el.UnexpectedShutdowns) allEntries.Add(("Shutdown", e));
                foreach (var e in el.DiskErrors) allEntries.Add(("Disk Error", e));
                foreach (var e in el.AppCrashes) allEntries.Add(("App Crash", e));

                var sorted = allEntries.OrderByDescending(x => x.Entry.Timestamp).Take(10).ToList();

                col.Item().Table(table =>
                {
                    table.ColumnsDefinition(cd =>
                    {
                        cd.ConstantColumn(100);  // Timestamp
                        cd.ConstantColumn(80);   // Category
                        cd.RelativeColumn();      // Message
                    });
                    table.Header(h =>
                    {
                        foreach (var label in new[] { "Timestamp", "Type", "Details" })
                            h.Cell().Background(BluePrimary).Padding(5)
                                .Text(label).FontSize(8).Bold().FontColor(Colors.White);
                    });
                    int i = 0;
                    foreach (var (category, entry) in sorted)
                    {
                        bool alt = i++ % 2 == 1;
                        string msg = !string.IsNullOrEmpty(entry.Message) ? entry.Message : entry.Source;
                        table.Cell().Element(c => DataCell(c, alt))
                            .Text(entry.Timestamp.ToString("MM/dd h:mm tt")).FontSize(8);
                        table.Cell().Element(c => DataCell(c, alt))
                            .Text(category).FontSize(8).Bold();
                        table.Cell().Element(c => DataCell(c, alt))
                            .Text(msg).FontSize(8);
                    }
                });
            }
        }

        // ── Recommended Actions ─────────────────────────────────────

        private void ComposeRecommendations(ColumnDescriptor col, DiagnosticResult r)
        {
            if (r.FlaggedIssues.Count == 0) return;

            col.Item().Element(c => SectionHeader(c, "Recommended Actions"));

            col.Item().Border(1).BorderColor(AmberWarn).Background("#FFFBEB").Padding(12).Column(inner =>
            {
                inner.Spacing(6);
                foreach (var issue in r.FlaggedIssues)
                {
                    string indicatorColor = issue.Severity switch
                    {
                        Severity.Critical => RedBad,
                        Severity.Warning => AmberWarn,
                        _ => GrayMuted
                    };
                    inner.Item().Row(row =>
                    {
                        row.ConstantItem(8).PaddingTop(3).Height(8).Background(indicatorColor);
                        row.ConstantItem(6);
                        row.RelativeItem().Column(c =>
                        {
                            c.Item().Text($"[{issue.Severity}] {issue.Category}").FontSize(8).Bold().FontColor(indicatorColor);
                            c.Item().Text(issue.Description).FontSize(9).FontColor(TextDark);
                            if (!string.IsNullOrEmpty(issue.Recommendation))
                                c.Item().Text($"\u2192 {issue.Recommendation}").FontSize(8).FontColor(GrayMuted);
                        });
                    });
                }
            });
        }

        // ── Optimization Plan / Results ─────────────────────────────

        private void ComposeOptimizationPlan(ColumnDescriptor col, DiagnosticResult result)
        {
            bool hasResults = result.OptimizationActions != null && result.OptimizationSummary != null;

            // Use stored actions if optimizer ran, otherwise build a fresh proposed plan
            List<OptimizationAction> actions;
            if (hasResults)
            {
                actions = result.OptimizationActions;
            }
            else
            {
                var optimizer = new Optimizer();
                actions = optimizer.BuildPlan(result);
            }

            if (actions.Count == 0) return;

            // ── Section header reflects whether optimizer was run ──
            if (hasResults)
            {
                col.Item().Element(c => SectionHeader(c, "Optimization Results"));
                ComposeOptimizationResultsSummary(col, result.OptimizationSummary, actions);
                ComposeOptimizationResultsTable(col, actions);
            }
            else
            {
                col.Item().Element(c => SectionHeader(c, "Optimization Plan"));
                ComposeOptimizationProposedSummary(col, actions);
                ComposeOptimizationProposedTable(col, actions);

                col.Item().PaddingTop(4)
                    .Text("Run the Optimize tool in the application to execute these actions automatically.")
                    .FontSize(8).Italic().FontColor(GrayMuted);
            }
        }

        /// <summary>Summary banner when the optimizer has been run.</summary>
        private void ComposeOptimizationResultsSummary(
            ColumnDescriptor col, OptimizationSummary summary, List<OptimizationAction> actions)
        {
            int succeeded = actions.Count(a => a.Status == ActionStatus.Success);
            int partial = actions.Count(a => a.Status == ActionStatus.PartialSuccess);
            int noChange = actions.Count(a => a.Status == ActionStatus.NoChange);
            int failed = actions.Count(a => a.Status == ActionStatus.Failed);
            int skipped = actions.Count(a => a.Status == ActionStatus.Skipped);
            int manual = actions.Count(a => !a.IsAutomatable);

            string overallColor = failed == 0 ? GreenGood : AmberWarn;

            col.Item().Background(GrayLight).Padding(10).Column(inner =>
            {
                inner.Spacing(4);

                inner.Item().Row(row =>
                {
                    row.RelativeItem().Column(c =>
                    {
                        c.Item().Text("Optimization completed").FontSize(10).Bold().FontColor(overallColor);
                        if (summary.TotalFreedMB > 0)
                            c.Item().Text($"Total disk space freed: {FormatMB(summary.TotalFreedMB)}")
                                .FontSize(9).Bold().FontColor(GreenGood);
                    });
                });

                inner.Item().PaddingTop(2).Row(row =>
                {
                    if (succeeded > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(GreenGood);
                            badge.AutoItem().PaddingLeft(4)
                                .Text($"{succeeded} Succeeded").FontSize(9).Bold().FontColor(GreenGood);
                        });
                    }
                    if (failed > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(RedBad);
                            badge.AutoItem().PaddingLeft(4)
                                .Text($"{failed} Failed").FontSize(9).Bold().FontColor(RedBad);
                        });
                    }
                    if (partial > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(AmberWarn);
                            badge.AutoItem().PaddingLeft(4)
                                .Text($"{partial} Partial").FontSize(9).Bold().FontColor(AmberWarn);
                        });
                    }
                    if (noChange > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(GrayMuted);
                            badge.AutoItem().PaddingLeft(4)
                                .Text($"{noChange} No Change").FontSize(9).FontColor(GrayMuted);
                        });
                    }
                    if (skipped > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(GrayMuted);
                            badge.AutoItem().PaddingLeft(4)
                                .Text($"{skipped} Skipped").FontSize(9).FontColor(GrayMuted);
                        });
                    }
                    if (manual > 0)
                    {
                        row.AutoItem().PaddingRight(12).Row(badge =>
                        {
                            badge.AutoItem().Width(10).Height(10).Background(AmberWarn);
                            badge.AutoItem().PaddingLeft(4)
                                .Text($"{manual} Manual").FontSize(9).FontColor(AmberWarn);
                        });
                    }
                });
            });
        }

        /// <summary>Results table showing status, actual freed space, and outcome per action.</summary>
        private void ComposeOptimizationResultsTable(ColumnDescriptor col, List<OptimizationAction> actions)
        {
            col.Item().PaddingTop(4).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.ConstantColumn(55);   // Status
                    cols.ConstantColumn(75);   // Category
                    cols.RelativeColumn();      // Action + result message
                    cols.ConstantColumn(60);   // Freed
                });

                table.Header(header =>
                {
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Status").FontSize(8).Bold().FontColor(Colors.White);
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Category").FontSize(8).Bold().FontColor(Colors.White);
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Action").FontSize(8).Bold().FontColor(Colors.White);
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Freed").FontSize(8).Bold().FontColor(Colors.White);
                });

                bool alt = false;
                foreach (var action in actions)
                {
                    string statusLabel = action.Status switch
                    {
                        ActionStatus.Success => "\u2713 Done",
                        ActionStatus.PartialSuccess => "~ Partial",
                        ActionStatus.NoChange => "\u2014 No Change",
                        ActionStatus.Failed => "\u2717 Failed",
                        ActionStatus.Skipped => "Skipped",
                        ActionStatus.Pending => "Not run",
                        _ => action.Status.ToString()
                    };
                    string statusColor = action.Status switch
                    {
                        ActionStatus.Success => GreenGood,
                        ActionStatus.PartialSuccess => AmberWarn,
                        ActionStatus.NoChange => GrayMuted,
                        ActionStatus.Failed => RedBad,
                        ActionStatus.Skipped => GrayMuted,
                        _ => TextDark
                    };

                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(statusLabel).FontSize(8).Bold().FontColor(statusColor);
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(action.Category).FontSize(8).FontColor(TextDark);
                    table.Cell().Element(c => DataCell(c, alt)).Column(c =>
                    {
                        c.Item().Text(action.Title).FontSize(8).Bold().FontColor(TextDark);
                        if (!string.IsNullOrEmpty(action.ResultMessage))
                            c.Item().Text(action.ResultMessage).FontSize(7).FontColor(
                                action.Status == ActionStatus.Failed ? RedBad : GrayMuted);
                        else if (!action.IsAutomatable)
                            c.Item().Text("Manual action required").FontSize(7).Italic().FontColor(AmberWarn);
                    });
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(action.ActualFreedMB > 0 ? FormatMB(action.ActualFreedMB) : "\u2014")
                        .FontSize(8).FontColor(action.ActualFreedMB > 0 ? GreenGood : GrayMuted);

                    alt = !alt;
                }
            });
        }

        /// <summary>Summary banner when the optimizer has NOT been run (proposed plan only).</summary>
        private void ComposeOptimizationProposedSummary(ColumnDescriptor col, List<OptimizationAction> actions)
        {
            long totalEstimatedMB = actions.Where(a => a.IsAutomatable).Sum(a => a.EstimatedFreeMB);
            int autoCount = actions.Count(a => a.IsAutomatable);
            int manualCount = actions.Count(a => !a.IsAutomatable);

            col.Item().Background(GrayLight).Padding(10).Row(row =>
            {
                row.RelativeItem().Column(c =>
                {
                    c.Item().Text($"{autoCount} automatable action{(autoCount == 1 ? "" : "s")}")
                        .FontSize(10).Bold().FontColor(TextDark);
                    if (totalEstimatedMB > 0)
                        c.Item().Text($"~{FormatMB(totalEstimatedMB)} recoverable").FontSize(9).FontColor(GreenGood);
                });
                if (manualCount > 0)
                {
                    row.ConstantItem(160).AlignRight().Column(c =>
                    {
                        c.Item().Text($"{manualCount} manual action{(manualCount == 1 ? "" : "s")} required")
                            .FontSize(9).FontColor(AmberWarn);
                    });
                }
            });
        }

        /// <summary>Proposed action table (before optimizer has been run).</summary>
        private void ComposeOptimizationProposedTable(ColumnDescriptor col, List<OptimizationAction> actions)
        {
            col.Item().PaddingTop(4).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.ConstantColumn(60);  // Risk
                    cols.ConstantColumn(80);  // Category
                    cols.RelativeColumn();     // Description
                    cols.ConstantColumn(60);  // Est. Space
                });

                table.Header(header =>
                {
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Risk").FontSize(8).Bold().FontColor(Colors.White);
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Category").FontSize(8).Bold().FontColor(Colors.White);
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Action").FontSize(8).Bold().FontColor(Colors.White);
                    header.Cell().Background(BluePrimary).Padding(5)
                        .Text("Est. Free").FontSize(8).Bold().FontColor(Colors.White);
                });

                bool alt = false;
                foreach (var action in actions)
                {
                    string riskColor = action.Risk switch
                    {
                        ActionRisk.RequiresReboot => RedBad,
                        ActionRisk.Moderate => AmberWarn,
                        _ => GreenGood
                    };
                    string riskLabel = action.Risk switch
                    {
                        ActionRisk.RequiresReboot => "Reboot",
                        ActionRisk.Moderate => "Moderate",
                        _ => "Safe"
                    };

                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(riskLabel).FontSize(8).Bold().FontColor(riskColor);
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(action.Category).FontSize(8).FontColor(TextDark);
                    table.Cell().Element(c => DataCell(c, alt)).Column(c =>
                    {
                        c.Item().Text(action.Title).FontSize(8).Bold().FontColor(TextDark);
                        c.Item().Text(action.Description).FontSize(7).FontColor(GrayMuted);
                        if (!action.IsAutomatable)
                            c.Item().Text("Manual action required").FontSize(7).Italic().FontColor(AmberWarn);
                    });
                    table.Cell().Element(c => DataCell(c, alt))
                        .Text(action.EstimatedFreeMB > 0 ? FormatMB(action.EstimatedFreeMB) : "\u2014")
                        .FontSize(8).FontColor(action.EstimatedFreeMB > 0 ? GreenGood : GrayMuted);

                    alt = !alt;
                }
            });
        }

        // ═══════════════════════════════════════════════════════════════
        //  FOOTER
        // ═══════════════════════════════════════════════════════════════

        private void ComposeFooter(IContainer container)
        {
            container.Column(col =>
            {
                col.Item().Height(1).Background(GrayBorder);
                col.Item().PaddingTop(6).Row(row =>
                {
                    row.RelativeItem()
                        .Text($"Generated by {DiagnosticFormatters.AppVersion} — Confidential")
                        .FontSize(8).FontColor(GrayMuted);
                    row.RelativeItem().AlignRight().Text(t =>
                    {
                        t.Span("Page ").FontSize(8).FontColor(GrayMuted);
                        t.CurrentPageNumber().FontSize(8).FontColor(GrayMuted);
                        t.Span(" of ").FontSize(8).FontColor(GrayMuted);
                        t.TotalPages().FontSize(8).FontColor(GrayMuted);
                    });
                });
            });
        }

        // ═══════════════════════════════════════════════════════════════
        //  REUSABLE COMPONENTS
        // ═══════════════════════════════════════════════════════════════

        private void SectionHeader(IContainer container, string title)
        {
            container
                .PaddingBottom(6)
                .BorderBottom(2)
                .BorderColor(BluePrimary)
                .PaddingBottom(4)
                .Text(title)
                .FontSize(13)
                .Bold()
                .FontColor(BluePrimary);
        }

        private void InfoRow(ColumnDescriptor col, string label, string value)
        {
            col.Item().Row(row =>
            {
                row.ConstantItem(100)
                    .Text($"{label}:")
                    .FontSize(9).Bold().FontColor(GrayMuted);
                row.RelativeItem()
                    .Text(string.IsNullOrWhiteSpace(value) ? "\u2014" : value)
                    .FontSize(9).FontColor(TextDark);
            });
        }

        private IContainer DataCell(IContainer container, bool isAlt)
        {
            return container
                .Background(isAlt ? GrayLight : Colors.White)
                .BorderBottom(1)
                .BorderColor(GrayBorder)
                .Padding(5);
        }

        private static string FormatUptime(TimeSpan uptime) => DiagnosticFormatters.FormatUptime(uptime);
        private static string FormatMB(long mb) => DiagnosticFormatters.FormatMB(mb);

        private static readonly string[] RatingColors = { GreenGood, AmberWarn, RedBad };

        private static string RateDnsColor(string dnsResponse) =>
            RatingColors[Math.Clamp(DiagnosticFormatters.RateDns(dnsResponse), 0, RatingColors.Length - 1)];

        private static string RatePingColor(string pingLatency) =>
            RatingColors[Math.Clamp(DiagnosticFormatters.RatePing(pingLatency), 0, RatingColors.Length - 1)];
    }
}
