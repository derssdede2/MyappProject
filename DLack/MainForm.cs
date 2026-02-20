using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using FontAwesome.Sharp;

namespace DLack
{
    public class MainForm : Form
    {
        // ── Theme Constants ──────────────────────────────────────────
        private static readonly Color BluePrimary = Color.FromArgb(59, 130, 246);
        private static readonly Color GreenSuccess = Color.FromArgb(16, 185, 129);
        private static readonly Color RedDanger = Color.FromArgb(239, 68, 68);
        private static readonly Color TextDark = Color.FromArgb(30, 41, 59);
        private static readonly Color TextMuted = Color.FromArgb(100, 116, 139);
        private static readonly Color BgMain = Color.FromArgb(240, 242, 245);
        private static readonly Color DisabledBg = Color.FromArgb(226, 232, 240);
        private static readonly Color DisabledFg = Color.FromArgb(100, 116, 139);

        private static readonly Font FontHeader = new Font("Segoe UI", 24, FontStyle.Bold);
        private static readonly Font FontButton = new Font("Segoe UI", 10, FontStyle.Bold);
        private static readonly Font FontSectionHead = new Font("Segoe UI", 12, FontStyle.Bold);
        private static readonly Font FontMetrics = new Font("Segoe UI", 11);
        private static readonly Font FontLog = new Font("Consolas", 9);
        private static readonly Font FontStatus = new Font("Segoe UI", 10);
        private static readonly Font FontSmall = new Font("Segoe UI", 9);

        // ── Layout Constants ─────────────────────────────────────────
        private const int EdgeMargin = 40;
        private const int HeaderHeight = 100;
        private const int StatusHeight = 40;
        private const int CardWidth = 650;
        private const int CardHeight = 590;
        private const int CardPadding = 20;
        private const int CardInner = CardWidth - (CardPadding * 2);
        private const int ButtonHeight = 50;
        private const int IconSz = 24;

        // ── Controls ─────────────────────────────────────────────────
        private Panel headerPanel;
        private Label lblTitle;
        private IconModernButton btnScan, btnOptimize, btnCancel;
        private ModernCard cardMetrics, cardLog;
        private TextBox txtMetrics, txtLog;
        private ModernProgressBar progressBar;
        private Label lblProgress;
        private Panel statusPanel;
        private IconPictureBox statusIconPic;
        private Label lblStatus;

        // ── State ────────────────────────────────────────────────────
        private Scanner scanner;
        private Optimizer optimizer;
        private ScanResult beforeScan;
        private ScanResult afterScan;

        public MainForm()
        {
            Text = "DLack - IT Performance Optimizer";
            ClientSize = new Size(1400, 800);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = BgMain;
            MinimumSize = new Size(1400, 800);
            DoubleBuffered = true;

            BuildUI();
        }

        // ═══════════════════════════════════════════════════════════════
        //  UI CONSTRUCTION
        // ═══════════════════════════════════════════════════════════════

        private void BuildUI()
        {
            SuspendLayout();

            BuildHeader();
            BuildMetricsCard();
            BuildLogCard();
            BuildStatusBar();

            ResumeLayout(true);

            Log("DLack v1.0 - IT Performance Optimizer");
            Log("Running with Administrator privileges");
            Log("Ready to scan.");
        }

        private void BuildHeader()
        {
            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = HeaderHeight,
                BackColor = BluePrimary
            };
            Controls.Add(headerPanel);

            lblTitle = new Label
            {
                Text = "DLack Optimizer",
                Location = new Point(EdgeMargin, 30),
                AutoSize = true,
                Font = FontHeader,
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };
            headerPanel.Controls.Add(lblTitle);

            btnScan = CreateHeaderButton("  Scan", IconChar.Bolt, BluePrimary, Color.White, 950, 130);
            btnScan.Click += BtnScan_Click;

            btnOptimize = CreateHeaderButton("  Optimize", IconChar.CheckCircle, Color.White, GreenSuccess, 1095, 140);
            btnOptimize.Enabled = false;
            btnOptimize.Click += BtnOptimize_Click;

            btnCancel = CreateHeaderButton("  Cancel", IconChar.TimesCircle, Color.White, RedDanger, 1250, 120);
            btnCancel.Enabled = false;
            btnCancel.Click += (s, e) => scanner?.CancelScan();
        }

        private IconModernButton CreateHeaderButton(string text, IconChar icon, Color iconColor, Color bgColor, int x, int width)
        {
            bool isWhiteBg = bgColor.ToArgb() == Color.White.ToArgb();
            var btn = new IconModernButton
            {
                Text = text,
                IconChar = icon,
                IconColor = iconColor,
                IconSize = IconSz,
                Location = new Point(x, 25),
                Size = new Size(width, ButtonHeight),
                BackColor = bgColor,
                ForeColor = isWhiteBg ? BluePrimary : Color.White,
                Font = FontButton,
                DisabledBackColor = DisabledBg,
                DisabledForeColor = DisabledFg,
                DisabledIconColor = DisabledFg
            };
            headerPanel.Controls.Add(btn);
            return btn;
        }

        private (ModernCard card, Panel headerPnl) CreateCard(string title, IconChar icon, int x)
        {
            var card = new ModernCard
            {
                Location = new Point(x, HeaderHeight + 30),
                Size = new Size(CardWidth, CardHeight)
            };
            Controls.Add(card);

            var hdrPanel = new Panel
            {
                Location = new Point(CardPadding, 15),
                Size = new Size(CardInner, 30),
                BackColor = Color.Transparent
            };
            card.Controls.Add(hdrPanel);

            var iconBox = new IconPictureBox
            {
                IconChar = icon,
                IconColor = BluePrimary,
                IconSize = IconSz,
                Location = new Point(0, 3),
                Size = new Size(IconSz, IconSz)
            };
            hdrPanel.Controls.Add(iconBox);

            var lbl = new Label
            {
                Text = title,
                Location = new Point(32, 0),
                AutoSize = true,
                Font = FontSectionHead,
                ForeColor = TextDark
            };
            hdrPanel.Controls.Add(lbl);

            return (card, hdrPanel);
        }

        private void BuildMetricsCard()
        {
            var (card, _) = CreateCard("SYSTEM METRICS", IconChar.ChartLine, EdgeMargin);
            cardMetrics = card;

            progressBar = new ModernProgressBar
            {
                Location = new Point(CardPadding, 55),
                Size = new Size(CardInner, 6),
                Visible = false
            };
            cardMetrics.Controls.Add(progressBar);

            lblProgress = new Label
            {
                Location = new Point(CardPadding, 67),
                Size = new Size(CardInner, 20),
                Font = FontSmall,
                ForeColor = TextMuted,
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = false
            };
            cardMetrics.Controls.Add(lblProgress);

            txtMetrics = new TextBox
            {
                Location = new Point(CardPadding, 95),
                Size = new Size(CardInner, 475),
                Multiline = true,
                Font = FontMetrics,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                Text = "Ready to scan.\r\n\r\nClick 'Scan' button to begin."
            };
            cardMetrics.Controls.Add(txtMetrics);
        }

        private void BuildLogCard()
        {
            var (card, _) = CreateCard("ACTIVITY LOG", IconChar.FileAlt, CardWidth + EdgeMargin + 20);
            cardLog = card;

            txtLog = new TextBox
            {
                Location = new Point(CardPadding, 55),
                Size = new Size(CardInner, 515),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = FontLog,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                ReadOnly = true
            };
            cardLog.Controls.Add(txtLog);
        }

        private void BuildStatusBar()
        {
            statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = StatusHeight,
                BackColor = Color.White
            };
            Controls.Add(statusPanel);

            statusIconPic = new IconPictureBox
            {
                IconChar = IconChar.CheckCircle,
                IconColor = GreenSuccess,
                IconSize = 18,
                Location = new Point(EdgeMargin, 11),
                Size = new Size(18, 18)
            };
            statusPanel.Controls.Add(statusIconPic);

            lblStatus = new Label
            {
                Text = "Ready",
                Location = new Point(EdgeMargin + 25, 10),
                AutoSize = true,
                Font = FontStatus,
                ForeColor = GreenSuccess
            };
            statusPanel.Controls.Add(lblStatus);
        }

        // ═══════════════════════════════════════════════════════════════
        //  BUTTON HANDLERS
        // ═══════════════════════════════════════════════════════════════

        private async void BtnScan_Click(object sender, EventArgs e)
        {
            try
            {
                EnsureScanner();

                SetButtonStates(scan: false, optimize: false, cancel: true);
                UpdateStatus("Scanning...", IconChar.Spinner, BluePrimary);
                ShowProgress(true);

                Log("\r\n=== STARTING SYSTEM SCAN ===");

                beforeScan = await scanner.RunScan(60);

                ShowProgress(false);
                btnCancel.Enabled = false;
                btnScan.Enabled = true;

                if (beforeScan?.SampleCount > 0)
                {
                    DisplayScanResults(beforeScan);
                    btnOptimize.Enabled = true;
                    UpdateStatus("Scan complete", IconChar.CheckCircle, GreenSuccess);
                    Log("✓ Scan complete! Click 'Optimize' to proceed.");
                }
                else
                {
                    UpdateStatus("Cancelled", IconChar.TimesCircle, RedDanger);
                    Log("✗ Scan was cancelled.");
                }
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
        }

        private async void BtnOptimize_Click(object sender, EventArgs e)
        {
            if (beforeScan == null) return;

            try
            {
                EnsureOptimizer();

                SetButtonStates(scan: false, optimize: false, cancel: false);
                UpdateStatus("Optimizing...", IconChar.Cog, BluePrimary);

                Log("\r\n=== RUNNING OPTIMIZATIONS ===");
                var result = await Task.Run(() => optimizer.RunOptimizations(beforeScan));

                Log("Waiting 5 seconds for changes to settle...");
                await Task.Delay(5000);

                // ── Post-optimization re-scan ──
                UpdateStatus("Re-scanning...", IconChar.Spinner, BluePrimary);
                Log("\r\n=== POST-OPTIMIZATION SCAN ===");
                ShowProgress(true);

                EnsureScanner();
                afterScan = await scanner.RunScan(30);

                ShowProgress(false);

                if (afterScan?.SampleCount > 0)
                {
                    result.AfterScan = afterScan;
                    DisplayComparison(beforeScan, afterScan);

                    // ── PDF Report ──
                    UpdateStatus("Generating PDF...", IconChar.FilePdf, BluePrimary);
                    Log("\r\n=== GENERATING PDF REPORT ===");

                    string pdfPath = GeneratePDFReport(result);

                    UpdateStatus("Complete!", IconChar.CheckCircle, GreenSuccess);
                    LogOptimizationSummary(result, pdfPath);

                    MessageBox.Show(
                        $"Optimization complete!\r\n\r\nReport saved to:\r\n{pdfPath}",
                        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    UpdateStatus("Post-scan cancelled", IconChar.TimesCircle, RedDanger);
                    Log("✗ Post-optimization scan was cancelled. PDF report was not generated.");
                }

                btnScan.Enabled = true;
                btnOptimize.Enabled = false;
                btnCancel.Enabled = false;
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  DISPLAY METHODS
        // ═══════════════════════════════════════════════════════════════

        private void Scanner_OnProgress(ScanProgress progress)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<ScanProgress>(Scanner_OnProgress), progress);
                return;
            }

            int pct = progress.Total > 0
                ? (int)((float)progress.Elapsed / progress.Total * 100)
                : 0;

            progressBar.Value = pct;
            lblProgress.Text = $"{progress.Elapsed}s / {progress.Total}s  ({pct}%)";

            SetMetricsText(
                $"SCANNING... {progress.Elapsed} / {progress.Total}s\r\n\r\n" +
                $"  CPU:          {progress.CurrentCpu,6}%\r\n\r\n" +
                $"  RAM:          {progress.CurrentRam,6:F1}%\r\n\r\n" +
                $"  TOP PROCESS:  {progress.TopCpuProcess}");
        }

        private void DisplayScanResults(ScanResult scan)
        {
            string cpuRating = RateValue(scan.AvgCpu, 30, 60, lowerIsBetter: true);
            string ramRating = RateValue(scan.AvgRam, 60, 85, lowerIsBetter: true);
            string diskRating = RateValue(scan.DiskFreePercent, 20, 10, lowerIsBetter: false);

            SetMetricsText(
                $"SCAN RESULTS\r\n" +
                $"{"".PadRight(40, '─')}\r\n\r\n" +

                $"  CPU Average:    {scan.AvgCpu,6}%     {cpuRating}\r\n" +
                $"  CPU Peak:       {scan.PeakCpu,6}%\r\n\r\n" +

                $"  RAM Usage:      {scan.AvgRam,6:F1}%     {ramRating}\r\n" +
                $"  RAM Used:       {scan.RamUsedMB,6:N0} MB\r\n" +
                $"  RAM Total:      {scan.RamTotalMB,6:N0} MB\r\n\r\n" +

                $"  Disk Free:      {scan.DiskFreePercent,6:F1}%     {diskRating}\r\n" +
                $"  Disk Avail:     {scan.DiskFreeMB / 1024,6:N0} GB\r\n\r\n" +

                $"  Uptime:         {FormatUptime(scan.Uptime)}\r\n" +
                $"  Power Plan:     {scan.PowerPlan}\r\n\r\n" +

                $"  Samples:        {scan.SampleCount}");

            Log("\r\n=== RECOMMENDATIONS ===");
            foreach (var rec in scan.Recommendations)
                Log($"  • {rec}");
        }

        private void DisplayComparison(ScanResult before, ScanResult after)
        {
            double cpuDelta = after.AvgCpu - before.AvgCpu;
            long ramDelta = after.RamUsedMB - before.RamUsedMB;
            double diskDelta = after.DiskFreePercent - before.DiskFreePercent;
            long ramFreed = before.RamUsedMB - after.RamUsedMB;

            string divider = "  " + "".PadRight(42, '─');

            SetMetricsText(
                $"OPTIMIZATION RESULTS\r\n" +
                $"{"".PadRight(40, '─')}\r\n\r\n" +

                $"                  {"BEFORE",8}  {"AFTER",8}  {"CHANGE",8}\r\n" +
                $"{divider}\r\n" +
                $"  CPU:            {before.AvgCpu,7}%  {after.AvgCpu,7}%  {FormatDelta(cpuDelta, "%", true)}\r\n" +
                $"  RAM (MB):       {before.RamUsedMB,7:N0}   {after.RamUsedMB,7:N0}   {FormatDelta(ramDelta, " MB", true)}\r\n" +
                $"  Disk Free:      {before.DiskFreePercent,7:F1}%  {after.DiskFreePercent,7:F1}%  {FormatDelta(diskDelta, "%", false)}\r\n\r\n" +

                $"{divider}\r\n" +
                $"  SUMMARY\r\n" +
                $"{divider}\r\n\r\n" +

                FormatSummaryLine("CPU", cpuDelta, "%", lowerIsBetter: true) +
                FormatSummaryLine("RAM", -ramDelta, " MB freed", lowerIsBetter: false) +
                FormatSummaryLine("Disk", diskDelta, "% more free", lowerIsBetter: false) +

                $"\r\n  Power Plan:     {after.PowerPlan}\r\n\r\n" +

                $"  Overall:        {GetOverallVerdict(cpuDelta, ramFreed, diskDelta)}");
        }

        // ═══════════════════════════════════════════════════════════════
        //  FORMATTING HELPERS
        // ═══════════════════════════════════════════════════════════════

        private static string FormatDelta(double delta, string unit, bool lowerIsBetter)
        {
            if (Math.Abs(delta) < 0.05) return "  ─  no change";

            bool improved = lowerIsBetter ? delta < 0 : delta > 0;
            string arrow = improved ? "▼" : "▲";
            string sign = delta > 0 ? "+" : "";

            string formatted = unit == " MB"
                ? $"{sign}{delta:N0}{unit}"
                : $"{sign}{delta:F1}{unit}";

            return $"  {arrow} {formatted}";
        }

        private static string FormatSummaryLine(string label, double delta, string unit, bool lowerIsBetter)
        {
            if (Math.Abs(delta) < 0.05)
                return $"  {label + ":",-14} ─  No change\r\n";

            bool improved = lowerIsBetter ? delta < 0 : delta > 0;
            string icon = improved ? "✓" : "✗";
            double abs = Math.Abs(delta);

            string formatted = unit == " MB freed"
                ? $"{abs:N0}{unit}"
                : $"{abs:F1}{unit}";

            return $"  {icon} {label + ":",-12} {formatted}\r\n";
        }

        private static string RateValue(double value, double warnThreshold, double critThreshold, bool lowerIsBetter)
        {
            bool isCrit = lowerIsBetter ? value >= critThreshold : value <= critThreshold;
            bool isWarn = lowerIsBetter ? value >= warnThreshold : value <= warnThreshold;

            if (isCrit) return "[!! CRITICAL]";
            if (isWarn) return "[!  WARNING]";
            return "[   OK]";
        }

        private static string FormatUptime(TimeSpan uptime)
        {
            if (uptime.TotalDays >= 1)
                return $"{uptime.Days}d {uptime.Hours}h {uptime.Minutes}m";
            if (uptime.TotalHours >= 1)
                return $"{uptime.Hours}h {uptime.Minutes}m";
            return $"{uptime.Minutes}m";
        }

        private static string GetOverallVerdict(double cpuDelta, long ramFreedMB, double diskDelta)
        {
            int score = 0;
            if (cpuDelta < -2) score++;
            if (ramFreedMB > 50) score++;
            if (diskDelta > 1) score++;

            return score switch
            {
                3 => "★★★  Significant improvement!",
                2 => "★★   Good improvement",
                1 => "★    Minor improvement",
                _ => "─    Minimal change detected"
            };
        }

        // ═══════════════════════════════════════════════════════════════
        //  UTILITY METHODS
        // ═══════════════════════════════════════════════════════════════

        private void EnsureScanner()
        {
            if (scanner != null) return;
            Log("Initializing scanner...");
            scanner = new Scanner();
            scanner.OnProgress += Scanner_OnProgress;
            scanner.OnLog += Log;
        }

        private void EnsureOptimizer()
        {
            if (optimizer != null) return;
            optimizer = new Optimizer();
            optimizer.OnLog += Log;
        }

        private void SetMetricsText(string text)
        {
            txtMetrics.Text = text;
            txtMetrics.SelectionStart = 0;
            txtMetrics.SelectionLength = 0;
        }

        private void SetButtonStates(bool scan, bool optimize, bool cancel)
        {
            btnScan.Enabled = scan;
            btnOptimize.Enabled = optimize;
            btnCancel.Enabled = cancel;
        }

        private void ShowProgress(bool visible)
        {
            progressBar.Visible = visible;
            lblProgress.Visible = visible;
        }

        private void UpdateStatus(string text, IconChar icon, Color color)
        {
            lblStatus.Text = text;
            lblStatus.ForeColor = color;
            statusIconPic.IconChar = icon;
            statusIconPic.IconColor = color;
        }

        private void HandleError(Exception ex)
        {
            ShowProgress(false);
            Log($"✗ ERROR: {ex.Message}");
            UpdateStatus("Error", IconChar.ExclamationTriangle, RedDanger);
            SetButtonStates(scan: true, optimize: false, cancel: false);
        }

        private void LogOptimizationSummary(OptimizationResult result, string pdfPath)
        {
            Log("✓ COMPLETE!");
            Log($"  CPU: {result.CpuImprovement:+0.0;-0.0}%");
            Log($"  RAM: {result.RamFreedMB:N0} MB freed");
            Log($"✓ PDF: {pdfPath}");
        }

        private string GeneratePDFReport(OptimizationResult result)
        {
            string computerName = Environment.MachineName;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string filename = $"DLack_{computerName}_{timestamp}.pdf";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (string.IsNullOrWhiteSpace(desktopPath) || !Directory.Exists(desktopPath))
                desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrWhiteSpace(desktopPath) || !Directory.Exists(desktopPath))
                desktopPath = Environment.CurrentDirectory;

            string fullPath = System.IO.Path.Combine(desktopPath, filename);

            new PDFReportGenerator().Generate(fullPath, result);
            return fullPath;
        }

        private void Log(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(Log), message);
                return;
            }

            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (scanner != null)
                {
                    scanner.OnProgress -= Scanner_OnProgress;
                    scanner.OnLog -= Log;
                    scanner.Dispose();
                    scanner = null;
                }

                if (optimizer != null)
                {
                    optimizer.OnLog -= Log;
                    optimizer = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}