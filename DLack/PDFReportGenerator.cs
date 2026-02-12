using System;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DLack
{
    public class PDFReportGenerator
    {
        // ── Brand Colors ─────────────────────────────────────────────
        private const string BluePrimary = "#3B82F6";
        private const string GreenGood = "#10B981";
        private const string RedBad = "#EF4444";
        private const string AmberWarn = "#F59E0B";
        private const string GrayMuted = "#64748B";
        private const string GrayLight = "#F1F5F9";
        private const string GrayBorder = "#E2E8F0";
        private const string TextDark = "#1E293B";

        public void Generate(string filePath, OptimizationResult result)
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

        // ═══════════════════════════════════════════════════════════════
        //  HEADER
        // ═══════════════════════════════════════════════════════════════

        private void ComposeHeader(IContainer container, OptimizationResult result)
        {
            container.Column(col =>
            {
                // Blue accent bar
                col.Item()
                    .Height(4)
                    .Background(BluePrimary);

                col.Item()
                    .PaddingTop(12)
                    .PaddingBottom(8)
                    .Row(row =>
                    {
                        row.RelativeItem().Column(left =>
                        {
                            left.Item()
                                .Text("DLack Performance Report")
                                .FontSize(22)
                                .Bold()
                                .FontColor(BluePrimary);

                            left.Item()
                                .Text("IT Performance Optimization Summary")
                                .FontSize(10)
                                .FontColor(GrayMuted);
                        });

                        row.ConstantItem(180).AlignRight().Column(right =>
                        {
                            right.Item()
                                .Text($"{result.Timestamp:MMMM dd, yyyy}")
                                .FontSize(10)
                                .FontColor(TextDark);

                            right.Item()
                                .Text($"{result.Timestamp:hh:mm tt}")
                                .FontSize(9)
                                .FontColor(GrayMuted);
                        });
                    });

                // Divider
                col.Item().Height(1).Background(GrayBorder);
            });
        }

        // ═══════════════════════════════════════════════════════════════
        //  CONTENT
        // ═══════════════════════════════════════════════════════════════

        private void ComposeContent(IContainer container, OptimizationResult result)
        {
            container
                .PaddingVertical(10)
                .Column(col =>
                {
                    col.Spacing(14);

                    ComposeSystemInfo(col, result);
                    ComposeComparisonTable(col, result);
                    ComposeActionsPerformed(col, result);
                    ComposeImprovementSummary(col, result);
                    ComposeRecommendations(col, result);
                });
        }

        // ── System Information ──────────────────────────────────────

        private void ComposeSystemInfo(ColumnDescriptor col, OptimizationResult result)
        {
            col.Item().Element(c => SectionHeader(c, "System Information"));

            col.Item()
                .Background(GrayLight)
                .Padding(12)
                .Row(row =>
                {
                    row.RelativeItem().Column(left =>
                    {
                        InfoRow(left, "Computer", Environment.MachineName);
                        InfoRow(left, "User", Environment.UserName);
                    });

                    row.RelativeItem().Column(right =>
                    {
                        InfoRow(right, "OS", Environment.OSVersion.ToString());
                        InfoRow(right, "Uptime", FormatUptime(result.BeforeScan.Uptime));
                    });
                });
        }

        // ── Before / After Comparison Table ─────────────────────────

        private void ComposeComparisonTable(ColumnDescriptor col, OptimizationResult result)
        {
            var before = result.BeforeScan;
            var after = result.AfterScan;

            col.Item().Element(c => SectionHeader(c, "Performance Comparison"));

            col.Item().Table(table =>
            {
                // Define columns
                table.ColumnsDefinition(cd =>
                {
                    cd.RelativeColumn(2.5f);  // Metric
                    cd.RelativeColumn(2f);    // Before
                    cd.RelativeColumn(2f);    // After
                    cd.RelativeColumn(2.5f);  // Change
                });

                // Header row
                table.Header(header =>
                {
                    foreach (var label in new[] { "Metric", "Before", "After", "Change" })
                    {
                        header.Cell()
                            .Background(BluePrimary)
                            .Padding(8)
                            .Text(label)
                            .FontSize(10)
                            .Bold()
                            .FontColor(Colors.White);
                    }
                });

                // CPU row
                double cpuDelta = (after?.AvgCpu ?? before.AvgCpu) - before.AvgCpu;
                TableRow(table, "CPU Average",
                    $"{before.AvgCpu}%",
                    after != null ? $"{after.AvgCpu}%" : "–",
                    cpuDelta, "%", lowerIsBetter: true, isAlt: false);

                // CPU Peak
                double peakDelta = (after?.PeakCpu ?? before.PeakCpu) - before.PeakCpu;
                TableRow(table, "CPU Peak",
                    $"{before.PeakCpu}%",
                    after != null ? $"{after.PeakCpu}%" : "–",
                    peakDelta, "%", lowerIsBetter: true, isAlt: true);

                // RAM %
                double ramPctDelta = (after?.AvgRam ?? before.AvgRam) - before.AvgRam;
                TableRow(table, "RAM Usage",
                    $"{before.AvgRam:F1}%",
                    after != null ? $"{after.AvgRam:F1}%" : "–",
                    ramPctDelta, "%", lowerIsBetter: true, isAlt: false);

                // RAM MB
                double ramMbDelta = (after?.RamUsedMB ?? before.RamUsedMB) - before.RamUsedMB;
                TableRow(table, "RAM Used",
                    $"{before.RamUsedMB:N0} MB",
                    after != null ? $"{after.RamUsedMB:N0} MB" : "–",
                    ramMbDelta, " MB", lowerIsBetter: true, isAlt: true);

                // Disk Free
                double diskDelta = (after?.DiskFreePercent ?? before.DiskFreePercent) - before.DiskFreePercent;
                TableRow(table, "Disk Free",
                    $"{before.DiskFreePercent:F1}%",
                    after != null ? $"{after.DiskFreePercent:F1}%" : "–",
                    diskDelta, "%", lowerIsBetter: false, isAlt: false);

                // Power Plan
                table.Cell().Element(c => DataCell(c, false))
                    .Text("Power Plan").FontSize(10);
                table.Cell().Element(c => DataCell(c, false))
                    .Text(before.PowerPlan ?? "–").FontSize(10);
                table.Cell().Element(c => DataCell(c, false))
                    .Text(after?.PowerPlan ?? "–").FontSize(10);
                table.Cell().Element(c => DataCell(c, false))
                    .Text(before.PowerPlan != after?.PowerPlan ? "Changed" : "–")
                    .FontSize(10)
                    .FontColor(before.PowerPlan != after?.PowerPlan ? GreenGood : GrayMuted);
            });
        }

        // ── Actions Performed ───────────────────────────────────────

        private void ComposeActionsPerformed(ColumnDescriptor col, OptimizationResult result)
        {
            col.Item().Element(c => SectionHeader(c, "Actions Performed"));

            col.Item()
                .Border(1)
                .BorderColor(GrayBorder)
                .Padding(12)
                .Column(inner =>
                {
                    inner.Spacing(4);

                    ActionItem(inner, "High Performance power plan", result.PowerPlanChanged);
                    ActionItem(inner, "Performance Visual Mode applied", result.VisualModeApplied);
                    ActionItem(inner, $"Cleared {result.TempFilesClearedMB} MB temp files",
                        result.TempFilesClearedMB > 0);
                    ActionItem(inner, "Emptied Recycle Bin", result.RecycleBinEmptied);
                    ActionItem(inner, "Restarted Windows Explorer", result.ExplorerRestarted);
                    ActionItem(inner, "Flushed DNS cache", result.DnsFlushed);
                });
        }

        // ── Improvement Summary ─────────────────────────────────────

        private void ComposeImprovementSummary(ColumnDescriptor col, OptimizationResult result)
        {
            if (result.AfterScan == null) return;

            col.Item().Element(c => SectionHeader(c, "Improvement Summary"));

            col.Item()
                .Background(GrayLight)
                .Padding(14)
                .Row(row =>
                {
                    // CPU improvement
                    row.RelativeItem().Element(c =>
                        MetricBox(c, "CPU", result.CpuImprovement, "%", lowerIsBetter: true));

                    row.ConstantItem(10); // spacer

                    // RAM freed
                    row.RelativeItem().Element(c =>
                        MetricBox(c, "RAM Freed", result.RamFreedMB, " MB", lowerIsBetter: false));

                    row.ConstantItem(10);

                    // Disk gained
                    row.RelativeItem().Element(c =>
                        MetricBox(c, "Disk Gained", result.DiskFreedPercent, "%", lowerIsBetter: false));
                });
        }

        // ── Recommendations ─────────────────────────────────────────

        private void ComposeRecommendations(ColumnDescriptor col, OptimizationResult result)
        {
            if (result.BeforeScan.Recommendations == null ||
                result.BeforeScan.Recommendations.Count == 0)
                return;

            col.Item().Element(c => SectionHeader(c, "Recommendations"));

            col.Item()
                .Border(1)
                .BorderColor(AmberWarn)
                .Background("#FFFBEB")
                .Padding(12)
                .Column(inner =>
                {
                    inner.Spacing(4);
                    foreach (var rec in result.BeforeScan.Recommendations)
                    {
                        inner.Item().Row(r =>
                        {
                            r.ConstantItem(16).Text("•").FontColor(AmberWarn).Bold();
                            r.RelativeItem().Text(rec).FontSize(10);
                        });
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

                col.Item()
                    .PaddingTop(6)
                    .Row(row =>
                    {
                        row.RelativeItem()
                            .Text("Generated by DLack v1.0")
                            .FontSize(8)
                            .FontColor(GrayMuted);

                        row.RelativeItem()
                            .AlignRight()
                            .Text(t =>
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
                row.ConstantItem(80)
                    .Text($"{label}:")
                    .FontSize(9)
                    .Bold()
                    .FontColor(GrayMuted);

                row.RelativeItem()
                    .Text(value)
                    .FontSize(9)
                    .FontColor(TextDark);
            });
        }

        private void ActionItem(ColumnDescriptor col, string text, bool performed)
        {
            string icon = performed ? "✓" : "–";
            string color = performed ? GreenGood : GrayMuted;

            col.Item().Row(row =>
            {
                row.ConstantItem(20)
                    .Text(icon)
                    .FontColor(color)
                    .Bold();

                row.RelativeItem()
                    .Text(text)
                    .FontSize(10)
                    .FontColor(performed ? TextDark : GrayMuted);
            });
        }

        private void MetricBox(IContainer container, string label, double value, string unit, bool lowerIsBetter)
        {
            bool isGood = lowerIsBetter ? value < 0 : value > 0;
            string color = Math.Abs(value) < 0.05 ? GrayMuted : (isGood ? GreenGood : RedBad);

            string sign = value > 0 ? "+" : "";
            string formatted = unit == " MB"
                ? $"{sign}{value:N0}{unit}"
                : $"{sign}{value:F1}{unit}";

            container
                .Background(Colors.White)
                .Border(1)
                .BorderColor(GrayBorder)
                .Padding(10)
                .AlignCenter()
                .Column(col =>
                {
                    col.Item()
                        .AlignCenter()
                        .Text(label)
                        .FontSize(9)
                        .FontColor(GrayMuted)
                        .Bold();

                    col.Item()
                        .AlignCenter()
                        .PaddingTop(4)
                        .Text(formatted)
                        .FontSize(16)
                        .Bold()
                        .FontColor(color);
                });
        }

        // ── Table helpers ───────────────────────────────────────────

        private IContainer DataCell(IContainer container, bool isAlt)
        {
            return container
                .Background(isAlt ? GrayLight : Colors.White)
                .BorderBottom(1)
                .BorderColor(GrayBorder)
                .Padding(7);
        }

        private void TableRow(TableDescriptor table,
            string metric, string beforeVal, string afterVal,
            double delta, string unit, bool lowerIsBetter, bool isAlt)
        {
            // Metric name
            table.Cell().Element(c => DataCell(c, isAlt))
                .Text(metric).FontSize(10).Bold();

            // Before value
            table.Cell().Element(c => DataCell(c, isAlt))
                .Text(beforeVal).FontSize(10);

            // After value
            table.Cell().Element(c => DataCell(c, isAlt))
                .Text(afterVal).FontSize(10);

            // Delta with color
            bool isGood = lowerIsBetter ? delta < 0 : delta > 0;
            string color = Math.Abs(delta) < 0.05 ? GrayMuted : (isGood ? GreenGood : RedBad);
            string arrow = Math.Abs(delta) < 0.05 ? "─" : (isGood ? "▼" : "▲");
            string sign = delta > 0 ? "+" : "";

            string formatted = Math.Abs(delta) < 0.05
                ? "No change"
                : unit == " MB"
                    ? $"{arrow} {sign}{delta:N0}{unit}"
                    : $"{arrow} {sign}{delta:F1}{unit}";

            table.Cell().Element(c => DataCell(c, isAlt))
                .Text(formatted)
                .FontSize(10)
                .FontColor(color);
        }

        // ── Utility ─────────────────────────────────────────────────

        private static string FormatUptime(TimeSpan uptime)
        {
            if (uptime.TotalDays >= 1)
                return $"{uptime.Days}d {uptime.Hours}h {uptime.Minutes}m";
            if (uptime.TotalHours >= 1)
                return $"{uptime.Hours}h {uptime.Minutes}m";
            return $"{uptime.Minutes}m";
        }
    }
}