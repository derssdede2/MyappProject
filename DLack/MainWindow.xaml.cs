using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace DLack
{
    public partial class MainWindow : Window
    {
        // ── State ────────────────────────────────────────────────────
        private Scanner scanner;
        private DiagnosticResult lastResult;
        private bool _isScanning;
        private int _lastProgressPercent;
        private int _lastPhaseTotal;
        private System.Windows.Threading.DispatcherTimer _elapsedTimer;
        private System.Diagnostics.Stopwatch _scanStopwatch;
        private System.Windows.Documents.Paragraph _logParagraph;
        private BeforeAfterSnapshot _preOptimizationSnapshot;
        // ── Segoe MDL2 Assets Icon Glyphs
        private static readonly FontFamily MdlFont = new("Segoe MDL2 Assets");
        private const string IcoCheck    = "\uE73E"; // ✓ CheckMark
        private const string IcoWarn     = "\uE7BA"; // ⚠ Warning
        private const string IcoError    = "\uE783"; // ⊘ StatusErrorFull
        private const string IcoInfo     = "\uE946"; // ℹ Info
        private const string IcoSearch   = "\uE721"; // Search
        private const string IcoSync     = "\uE72C"; // Sync
        private const string IcoShield   = "\uE72E"; // Shield (Security)
        private const string IcoGlobe    = "\uE774"; // Globe (Network)
        private const string IcoContact  = "\uE77B"; // Contact (User Profile)
        private const string IcoMail     = "\uE715"; // Mail (Outlook)
        private const string IcoPage     = "\uE7C3"; // Page
        private const string IcoPaint    = "\uE790"; // Color (Visual Settings)
        private const string IcoDevice   = "\uE770"; // Laptop (System Overview)
        private const string IcoBattery  = "\uEA93"; // Battery
        private const string IcoApps     = "\uE71D"; // AllApps (Installed Software)
        private const string IcoBright   = "\uE706"; // Brightness
        private const string IcoGrid     = "\uE80A"; // ViewAll (Categories)
        private const string IcoClip     = "\uE8C8"; // Paste (Scan Summary)
        private const string IcoCancel   = "\uE711"; // Cancel
        private const string IcoHealth   = "\uE8BE"; // HeartFill (Health Score)
        private const string IcoCpu      = "\uE713"; // Settings/Gear (CPU)
        private const string IcoRam      = "\uE7F1"; // Processing (RAM)
        private const string IcoDisk     = "\uEDA2"; // HardDrive
        private const string IcoThermo   = "\uE9CA"; // Thermometer (Thermal)
        private const string IcoStartup  = "\uE768"; // Forward (Startup)
        private const string IcoBrowser  = "\uEB41"; // WebBrowser
        private const string IcoOffice   = "\uE8A5"; // SelectAll (Office)
        private const string IcoScan     = "\uE9D9"; // Scan
        private const string IcoExport   = "\uE8B5"; // Export
        private const string IcoGpu      = "\uE7F4"; // TVMonitor (GPU)

        /// <summary>Creates an icon TextBlock using Segoe MDL2 Assets.</summary>
        private static TextBlock MdlIcon(string glyph, double size, Color? color = null)
        {
            var tb = new TextBlock
            {
                Text = glyph,
                FontFamily = MdlFont,
                FontSize = size,
                VerticalAlignment = VerticalAlignment.Center
            };
            if (color.HasValue)
                tb.Foreground = new SolidColorBrush(color.Value);
            return tb;
        }

        public MainWindow()
        {
            InitializeComponent();
            SourceInitialized += MainWindow_SourceInitialized;
            ThemeManager.ThemeChanged += OnThemeChanged;
            btnToggleTheme.Content = ThemeManager.IsDark ? IcoBright : "\uEC46";
            UpdateHeaderLogo();
            lblAppVersion.Text = DiagnosticFormatters.AppVersion;

            // Initialize log paragraph for RichTextBox
            _logParagraph = new System.Windows.Documents.Paragraph();
            txtLog.Document.Blocks.Clear();
            txtLog.Document.Blocks.Add(_logParagraph);

            if (progressFill.Parent is FrameworkElement progressTrack)
                progressTrack.SizeChanged += ProgressTrack_SizeChanged;

            Log(DiagnosticFormatters.AppVersion);
            bool isAdmin = new System.Security.Principal.WindowsPrincipal(
                System.Security.Principal.WindowsIdentity.GetCurrent())
                .IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            Log(isAdmin ? "Running with Administrator privileges" : "Running WITHOUT Administrator privileges (some checks may fail)");
            Log("Ready to scan.");
        }

        // ── Constrain maximized window to working area (respect taskbar) ──

        private void MainWindow_SourceInitialized(object sender, EventArgs e)
        {
            var handle = new WindowInteropHelper(this).Handle;
            HwndSource.FromHwnd(handle)?.AddHook(WndProc);
        }

        private const int WM_GETMINMAXINFO = 0x0024;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X, Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MINMAXINFO
        {
            public POINT ptReserved;
            public POINT ptMaxSize;
            public POINT ptMaxPosition;
            public POINT ptMinTrackSize;
            public POINT ptMaxTrackSize;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_GETMINMAXINFO)
            {
                var mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);
                var monitor = MonitorFromWindow(hwnd, 0x00000002 /* MONITOR_DEFAULTTONEAREST */);
                var mi = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
                if (GetMonitorInfo(monitor, ref mi))
                {
                    mmi.ptMaxPosition.X = mi.rcWork.Left - mi.rcMonitor.Left;
                    mmi.ptMaxPosition.Y = mi.rcWork.Top - mi.rcMonitor.Top;
                    mmi.ptMaxSize.X = mi.rcWork.Right - mi.rcWork.Left;
                    mmi.ptMaxSize.Y = mi.rcWork.Bottom - mi.rcWork.Top;
                }

                // Enforce minimum window size (DIP → physical pixels for WM_GETMINMAXINFO)
                var dpi = VisualTreeHelper.GetDpi(this);
                mmi.ptMinTrackSize.X = (int)(MinWidth * dpi.DpiScaleX);
                mmi.ptMinTrackSize.Y = (int)(MinHeight * dpi.DpiScaleY);

                Marshal.StructureToPtr(mmi, lParam, false);
                handled = true;
            }
            return IntPtr.Zero;
        }

        // ═══════════════════════════════════════════════════════════════
        //  WINDOW CHROME HANDLERS
        // ═══════════════════════════════════════════════════════════════

        private void BtnMinimize_Click(object sender, RoutedEventArgs e)
            => WindowState = WindowState.Minimized;

        private void BtnMaximize_Click(object sender, RoutedEventArgs e)
            => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                btnMaxRestore.Content = "\uE923"; // Restore icon
                btnMaxRestore.ToolTip = "Restore Down";
                WindowBorder.CornerRadius = new CornerRadius(0);
                WindowBorder.BorderThickness = new Thickness(0);
                HeaderBorder.CornerRadius = new CornerRadius(0);
                chromeButtonPanel.Margin = new Thickness(0, 12, 0, 0);
            }
            else
            {
                btnMaxRestore.Content = "\uE922"; // Maximize icon
                btnMaxRestore.ToolTip = "Maximize";
                WindowBorder.CornerRadius = new CornerRadius(14);
                WindowBorder.BorderThickness = new Thickness(1);
                HeaderBorder.CornerRadius = new CornerRadius(13, 13, 0, 0);
                chromeButtonPanel.Margin = new Thickness(0);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
            => Close();

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape && btnCancel.IsEnabled)
            {
                BtnCancel_Click(btnCancel, new RoutedEventArgs());
                e.Handled = true;
                return;
            }

            if (e.KeyboardDevice.Modifiers == System.Windows.Input.ModifierKeys.Control)
            {
                switch (e.Key)
                {
                    case System.Windows.Input.Key.R when btnScan.IsEnabled:
                        BtnScan_Click(btnScan, new RoutedEventArgs());
                        e.Handled = true;
                        break;
                    case System.Windows.Input.Key.E when btnExport.Visibility == Visibility.Visible:
                        BtnExport_Click(btnExport, new RoutedEventArgs());
                        e.Handled = true;
                        break;
                    case System.Windows.Input.Key.T:
                        BtnToggleTheme_Click(btnToggleTheme, new RoutedEventArgs());
                        e.Handled = true;
                        break;
                    case System.Windows.Input.Key.O when btnOptimize.Visibility == Visibility.Visible:
                        BtnOptimize_Click(btnOptimize, new RoutedEventArgs());
                        e.Handled = true;
                        break;
                    case System.Windows.Input.Key.L:
                        _logParagraph.Inlines.Clear();
                        Log("Log cleared.");
                        e.Handled = true;
                        break;
                }
            }
        }

        private void BtnToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            ThemeManager.IsDark = !ThemeManager.IsDark;
            btnToggleTheme.Content = ThemeManager.IsDark ? IcoBright : "\uEC46";
        }

        private void OnThemeChanged()
        {
            UpdateHeaderLogo();
            if (lastResult != null)
                DisplayDiagnosticResults(lastResult);
        }

        private void UpdateHeaderLogo()
        {
            try
            {
                // Both themes have dark headers — always use white logo
                var uri = new Uri("pack://application:,,,/Assets/pillsbury-logo-white.png");
                var bmp = new System.Windows.Media.Imaging.BitmapImage();
                bmp.BeginInit();
                bmp.UriSource = uri;
                bmp.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                bmp.EndInit();
                bmp.Freeze();
                headerLogo.Source = bmp;
                headerLogo.Visibility = Visibility.Visible;
                logoDivider.Visibility = Visibility.Visible;
            }
            catch
            {
                // Logo file missing — hide logo and divider
                headerLogo.Visibility = Visibility.Collapsed;
                logoDivider.Visibility = Visibility.Collapsed;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  HEALTH DASHBOARD
        // ═══════════════════════════════════════════════════════════════

        private void BuildHealthDashboard(DiagnosticResult result)
        {
            healthCards.Children.Clear();

            int critical = result.FlaggedIssues.Count(i => i.Severity == Severity.Critical);
            int warning = result.FlaggedIssues.Count(i => i.Severity == Severity.Warning);
            int info = result.FlaggedIssues.Count(i => i.Severity == Severity.Info);

            int scannedCategories = _lastPhaseTotal > 0 ? _lastPhaseTotal : 19;

            // Health score arc gauge
            Color scoreColor = result.HealthScore >= 70 ? ThemeManager.GoodFg :
                               result.HealthScore >= 50 ? ThemeManager.WarnFg : ThemeManager.CritFg;
            healthCards.Children.Add(BuildArcGaugeCard("Health", result.HealthScore, 100, scoreColor,
                $"{result.HealthScore}", "Overall Score"));

            // Issue count cards
            healthCards.Children.Add(BuildDashCard(IcoCancel, "Critical", critical.ToString(),
                critical > 0 ? ThemeManager.CritFg : ThemeManager.GoodFg, "Issues requiring immediate attention"));
            healthCards.Children.Add(BuildDashCard(IcoWarn, "Warnings", warning.ToString(),
                warning > 0 ? ThemeManager.WarnFg : ThemeManager.GoodFg, "Issues that may affect performance"));
            healthCards.Children.Add(BuildDashCard(IcoGrid, "Scanned", scannedCategories.ToString(),
                ThemeManager.TextMuted, "Diagnostic categories scanned"));

            // Staggered entrance for each card
            for (int i = 0; i < healthCards.Children.Count; i++)
            {
                if (healthCards.Children[i] is FrameworkElement card)
                {
                    card.Opacity = 0;
                    card.RenderTransform = new TranslateTransform(0, 18);
                    int delay = i * 100;
                    card.Loaded += CreateStaggeredEntrance(card, delay);
                }
            }

            healthDashboard.Visibility = Visibility.Visible;
        }

        private static RoutedEventHandler CreateStaggeredEntrance(FrameworkElement card, int delayMs)
        {
            return async (_, _) =>
            {
                if (delayMs > 0)
                    await Task.Delay(delayMs);

                var fade = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(450))
                {
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
                };
                card.BeginAnimation(UIElement.OpacityProperty, fade);

                var slide = new DoubleAnimation(18, 0, TimeSpan.FromMilliseconds(500))
                {
                    EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut }
                };
                ((TranslateTransform)card.RenderTransform).BeginAnimation(TranslateTransform.YProperty, slide);
            };
        }

        /// <summary>Builds a dashboard card with a semicircle arc gauge.</summary>
        private static FrameworkElement BuildArcGaugeCard(
            string title, double value, double max, Color arcColor,
            string centerText, string subtitle)
        {
            // Gauge geometry — all in WPF screen coords (0°=right, clockwise, Y-down)
            const double radius = 34;
            const double strokeWidth = 7;
            const double canvasSize = 90;
            double cx = canvasSize / 2;
            double cy = canvasSize / 2;

            double fraction = Math.Clamp(value / max, 0, 1);
            const double totalSweepDeg = 270.0;

            var card = new Border
            {
                CornerRadius = new CornerRadius(10),
                Background = new SolidColorBrush(ThemeManager.SurfaceAlt),
                BorderBrush = new SolidColorBrush(ThemeManager.Border),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(4, 4, 4, 4),
                MinWidth = 120,
                ToolTip = $"{title}: {centerText}",
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new TranslateTransform(0, 0)
            };
            AttachHoverLift(card);

            var sp = new StackPanel { HorizontalAlignment = HorizontalAlignment.Center };

            // Title
            sp.Children.Add(new TextBlock
            {
                Text = title,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 2)
            });

            // Grid overlay: arc + centered text
            var gaugeGrid = new Grid
            {
                Width = canvasSize,
                Height = canvasSize,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            var canvas = new Canvas { Width = canvasSize, Height = canvasSize };

            // Background track (full 270° arc)
            canvas.Children.Add(CreateArcPath(cx, cy, radius,
                totalSweepDeg, totalSweepDeg, strokeWidth,
                Color.FromArgb(35, arcColor.R, arcColor.G, arcColor.B)));

            // Filled arc (proportional to value) — animated sweep entrance
            double fillSweep = totalSweepDeg * fraction;
            if (fillSweep > 0.5)
            {
                var fillPath = CreateArcPath(cx, cy, radius,
                    totalSweepDeg, fillSweep, strokeWidth, arcColor);

                // Animate via StrokeDashArray/StrokeDashOffset
                double arcLength = 2 * Math.PI * radius * (fillSweep / 360.0);
                fillPath.StrokeDashArray = new DoubleCollection { arcLength / strokeWidth, 9999 };
                fillPath.StrokeDashOffset = arcLength / strokeWidth;
                fillPath.Loaded += (_, _) =>
                {
                    var anim = new DoubleAnimation
                    {
                        From = arcLength / strokeWidth,
                        To = 0,
                        Duration = TimeSpan.FromMilliseconds(1100),
                        EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut },
                        BeginTime = TimeSpan.FromMilliseconds(200)
                    };
                    fillPath.BeginAnimation(Shape.StrokeDashOffsetProperty, anim);
                };
                canvas.Children.Add(fillPath);
            }

            gaugeGrid.Children.Add(canvas);

            // Center value text
            gaugeGrid.Children.Add(new TextBlock
            {
                Text = centerText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(arcColor),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 8)
            });

            sp.Children.Add(gaugeGrid);

            // Subtitle
            sp.Children.Add(new TextBlock
            {
                Text = subtitle,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 10,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, -6, 0, 0)
            });

            card.Child = sp;
            return card;
        }

        /// <summary>
        /// Creates an arc Path for a gauge. Uses WPF screen coordinates
        /// (0° = right/3-o'clock, angles increase clockwise, Y-axis points down).
        /// The gauge starts at bottom-left (135° in screen coords) and sweeps clockwise.
        /// </summary>
        private static Path CreateArcPath(
            double cx, double cy, double radius,
            double totalArcDeg, double sweepDeg,
            double strokeWidth, Color color)
        {
            // 135° in WPF screen coords = bottom-left of the circle
            const double startScreenDeg = 135.0;

            double startRad = startScreenDeg * Math.PI / 180.0;
            double endRad = (startScreenDeg + sweepDeg) * Math.PI / 180.0;

            // WPF screen coords: x = cos, y = sin (Y-down)
            double x1 = cx + radius * Math.Cos(startRad);
            double y1 = cy + radius * Math.Sin(startRad);
            double x2 = cx + radius * Math.Cos(endRad);
            double y2 = cy + radius * Math.Sin(endRad);

            var figure = new PathFigure
            {
                StartPoint = new Point(x1, y1),
                IsClosed = false
            };
            figure.Segments.Add(new ArcSegment
            {
                Point = new Point(x2, y2),
                Size = new Size(radius, radius),
                SweepDirection = SweepDirection.Clockwise,
                IsLargeArc = sweepDeg > 180
            });

            var geometry = new PathGeometry();
            geometry.Figures.Add(figure);

            return new Path
            {
                Data = geometry,
                Stroke = new SolidColorBrush(color),
                StrokeThickness = strokeWidth,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round
            };
        }

        private static FrameworkElement BuildDashCard(string icon, string label, string value, Color valueColor, string tooltip = null)
        {
            var card = new Border
            {
                CornerRadius = new CornerRadius(10),
                Background = new SolidColorBrush(ThemeManager.SurfaceAlt),
                BorderBrush = new SolidColorBrush(ThemeManager.Border),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(18, 12, 18, 12),
                Margin = new Thickness(4, 4, 4, 4),
                MinWidth = 110,
                ToolTip = tooltip,
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new TranslateTransform(0, 0)
            };
            AttachHoverLift(card);

            var sp = new StackPanel { HorizontalAlignment = HorizontalAlignment.Center };

            var iconBlock = MdlIcon(icon, 20, valueColor);
            iconBlock.HorizontalAlignment = HorizontalAlignment.Center;
            iconBlock.Margin = new Thickness(0, 0, 0, 4);
            sp.Children.Add(iconBlock);

            sp.Children.Add(new TextBlock
            {
                Text = value,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(valueColor),
                HorizontalAlignment = HorizontalAlignment.Center
            });

            sp.Children.Add(new TextBlock
            {
                Text = label,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 11,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                HorizontalAlignment = HorizontalAlignment.Center
            });

            card.Child = sp;
            return card;
        }

        /// <summary>Attaches a subtle Y-translate hover animation to a card.</summary>
        private static void AttachHoverLift(Border card)
        {
            card.MouseEnter += (_, _) =>
            {
                var anim = new DoubleAnimation { To = -3, Duration = TimeSpan.FromMilliseconds(200),
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } };
                ((TranslateTransform)card.RenderTransform).BeginAnimation(TranslateTransform.YProperty, anim);
            };
            card.MouseLeave += (_, _) =>
            {
                var anim = new DoubleAnimation { To = 0, Duration = TimeSpan.FromMilliseconds(300),
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } };
                ((TranslateTransform)card.RenderTransform).BeginAnimation(TranslateTransform.YProperty, anim);
            };
        }

        // ═══════════════════════════════════════════════════════════════
        //  BEFORE/AFTER COMPARISON
        // ═══════════════════════════════════════════════════════════════

        private void BuildBeforeAfterComparison(BeforeAfterSnapshot before, BeforeAfterSnapshot after)
        {
            // Insert at the top of metricsPanel, right after the health dashboard
            int insertIndex = 0;

            var card = new Border
            {
                CornerRadius = new CornerRadius(12),
                Background = new SolidColorBrush(ThemeManager.SurfaceAlt),
                BorderBrush = new SolidColorBrush(ThemeManager.AccentBlue),
                BorderThickness = new Thickness(1.5),
                Padding = new Thickness(20, 16, 20, 16),
                Margin = new Thickness(0, 0, 0, 12),
                Opacity = 0,
                RenderTransform = new TranslateTransform(0, 16)
            };

            // Fade + slide entrance
            card.Loaded += (_, _) =>
            {
                var fadeIn = new DoubleAnimation
                {
                    From = 0, To = 1,
                    Duration = TimeSpan.FromMilliseconds(550),
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
                };
                card.BeginAnimation(OpacityProperty, fadeIn);

                var slideUp = new DoubleAnimation
                {
                    From = 16, To = 0,
                    Duration = TimeSpan.FromMilliseconds(600),
                    EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut }
                };
                ((TranslateTransform)card.RenderTransform).BeginAnimation(TranslateTransform.YProperty, slideUp);
            };

            var outerStack = new StackPanel();

            // Header
            var header = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 12) };
            header.Children.Add(MdlIcon("\uE8AB", 16, ThemeManager.AccentBlue)); // CompareIcon
            header.Children.Add(new TextBlock
            {
                Text = "  BEFORE → AFTER OPTIMIZATION",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(ThemeManager.AccentBlue),
                VerticalAlignment = VerticalAlignment.Center
            });
            outerStack.Children.Add(header);

            // Comparison grid
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Arrow column
            var arrowBlock = new TextBlock
            {
                Text = "→",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 24,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 16, 0)
            };
            Grid.SetColumn(arrowBlock, 1);
            grid.Children.Add(arrowBlock);

            // Before column
            var beforePanel = BuildSnapshotColumn("Before", before, ThemeManager.TextMuted);
            Grid.SetColumn(beforePanel, 0);
            grid.Children.Add(beforePanel);

            // After column
            var afterPanel = BuildSnapshotColumn("After", after, ThemeManager.GoodFg);
            Grid.SetColumn(afterPanel, 2);
            grid.Children.Add(afterPanel);

            outerStack.Children.Add(grid);

            // Delta summary
            int scoreDelta = after.HealthScore - before.HealthScore;
            long diskDelta = after.DiskFreeMB - before.DiskFreeMB;
            int issueDelta = after.IssueCount - before.IssueCount;

            var deltaPanel = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 12, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            };

            if (scoreDelta != 0)
                deltaPanel.Children.Add(BuildDeltaBadge("Health", scoreDelta > 0 ? $"+{scoreDelta}" : $"{scoreDelta}", scoreDelta > 0));
            if (diskDelta > 0)
                deltaPanel.Children.Add(BuildDeltaBadge("Disk Freed", $"+{FormatMB(diskDelta)}", true));
            if (issueDelta != 0)
                deltaPanel.Children.Add(BuildDeltaBadge("Issues", issueDelta < 0 ? $"{issueDelta}" : $"+{issueDelta}", issueDelta < 0));

            if (deltaPanel.Children.Count > 0)
                outerStack.Children.Add(deltaPanel);

            card.Child = outerStack;
            metricsPanel.Children.Insert(insertIndex, card);

            Log("\r\n=== BEFORE → AFTER OPTIMIZATION ===");
            Log($"  Health Score: {before.HealthScore} → {after.HealthScore} ({(scoreDelta >= 0 ? "+" : "")}{scoreDelta})");
            Log($"  Disk Free:   {FormatMB(before.DiskFreeMB)} → {FormatMB(after.DiskFreeMB)} ({(diskDelta >= 0 ? "+" : "")}{FormatMB(diskDelta)})");
            Log($"  RAM Usage:   {before.RamUsedPercent}% → {after.RamUsedPercent}%");
            Log($"  Startup:     {before.StartupCount} → {after.StartupCount}");
            Log($"  Issues:      {before.IssueCount} → {after.IssueCount}");
        }

        private static FrameworkElement BuildSnapshotColumn(string title, BeforeAfterSnapshot snap, Color accentColor)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock
            {
                Text = title.ToUpperInvariant(),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 8)
            });

            AddMetricRow(panel, "Health Score", $"{snap.HealthScore}/100", accentColor);
            AddMetricRow(panel, "Disk Free", FormatMB(snap.DiskFreeMB), accentColor);
            AddMetricRow(panel, "RAM Usage", $"{snap.RamUsedPercent}%", accentColor);
            AddMetricRow(panel, "Startup Items", $"{snap.StartupCount}", accentColor);
            AddMetricRow(panel, "Issues", $"{snap.IssueCount}", accentColor);

            return panel;
        }

        private static void AddMetricRow(StackPanel parent, string label, string value, Color valueColor)
        {
            var row = new Grid { Margin = new Thickness(0, 2, 0, 2) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var lbl = new TextBlock
            {
                Text = label,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 11,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(lbl, 0);
            row.Children.Add(lbl);

            var val = new TextBlock
            {
                Text = value,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(valueColor),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetColumn(val, 1);
            row.Children.Add(val);

            parent.Children.Add(row);
        }

        private static Border BuildDeltaBadge(string label, string delta, bool isPositive)
        {
            Color fg = isPositive ? ThemeManager.GoodFg : ThemeManager.WarnFg;
            Color bg = isPositive ? ThemeManager.GoodBg : ThemeManager.WarnBg;

            var badge = new Border
            {
                CornerRadius = new CornerRadius(6),
                Background = new SolidColorBrush(bg),
                Padding = new Thickness(10, 4, 10, 4),
                Margin = new Thickness(4, 0, 4, 0)
            };
            badge.Child = new TextBlock
            {
                Text = $"{label}: {delta}",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(fg)
            };
            return badge;
        }

        // ═══════════════════════════════════════════════════════════════
        //  WHAT CHANGED SINCE LAST SCAN
        // ═══════════════════════════════════════════════════════════════

        private void BuildWhatChangedPanel(ScanSnapshot previous, ScanSnapshot current)
        {
            var deltas = ScanHistoryManager.CompareTo(previous, current);
            // Only show if something actually changed
            if (deltas.All(d => d.Direction == 0)) return;

            var card = new Border
            {
                CornerRadius = new CornerRadius(10),
                Background = new SolidColorBrush(ThemeManager.SurfaceAlt),
                BorderBrush = new SolidColorBrush(ThemeManager.Border),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(16, 12, 16, 12),
                Margin = new Thickness(0, 0, 0, 10),
                Opacity = 0,
                RenderTransform = new TranslateTransform(0, 12)
            };
            card.Loaded += (_, _) =>
            {
                card.BeginAnimation(OpacityProperty,
                    new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(450))
                    { EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } });
                ((TranslateTransform)card.RenderTransform).BeginAnimation(
                    TranslateTransform.YProperty,
                    new DoubleAnimation(12, 0, TimeSpan.FromMilliseconds(500))
                    { EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut } });
            };

            var sp = new StackPanel();

            // Header
            var header = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            header.Children.Add(MdlIcon(IcoSync, 13, ThemeManager.AccentBlue));
            header.Children.Add(new TextBlock
            {
                Text = $"  CHANGES SINCE {previous.Timestamp:MMM dd, h:mm tt}",
                FontFamily = new FontFamily("Segoe UI"), FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(ThemeManager.AccentBlue),
                VerticalAlignment = VerticalAlignment.Center
            });
            sp.Children.Add(header);

            // Delta rows
            foreach (var (metric, before, after, direction) in deltas)
            {
                if (direction == 0) continue;

                Color fg = direction > 0 ? ThemeManager.GoodFg : ThemeManager.CritFg;
                string arrow = direction > 0 ? "▲" : "▼";

                var row = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var label = new TextBlock
                {
                    Text = metric,
                    FontFamily = new FontFamily("Segoe UI"), FontSize = 11,
                    Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(label, 0);
                row.Children.Add(label);

                var val = new TextBlock
                {
                    Text = $"{before} → {after}  {arrow}",
                    FontFamily = new FontFamily("Segoe UI"), FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(fg),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(val, 1);
                row.Children.Add(val);

                sp.Children.Add(row);
            }

            card.Child = sp;
            // Insert after health dashboard, before metrics
            metricsPanel.Children.Insert(0, card);
        }

        // ═══════════════════════════════════════════════════════════════
        //  BUTTON HANDLERS
        // ═══════════════════════════════════════════════════════════════

        private async void BtnScan_Click(object sender, RoutedEventArgs e)
        {
            if (_isScanning) return;
            _isScanning = true;
            try
            {
                // Dispose old scanner to avoid event accumulation
                if (scanner != null)
                {
                    scanner.OnProgress -= Scanner_OnProgress;
                    scanner.OnLog -= Log;
                    scanner.Dispose();
                    scanner = null;
                }
                EnsureScanner();

                SetButtonStates(scan: false, export: false, cancel: true);
                lblExportText.Text = "Export Report";
                btnCancelLabel.Text = "Stop";
                btnCancelIcon.Text = IcoCancel;
                UpdateStatus("Scanning...", IcoSync, ThemeManager.BrushPrimary);
                ShowProgress(true);
                _lastProgressPercent = 0;
                _lastPhaseTotal = 0;
                progressFill.BeginAnimation(WidthProperty, null);
                progressFill.Width = 0;
                lblProgress.Text = "";
                StartElapsedTimer();

                Log("\r\n=== STARTING DIAGNOSTIC SCAN ===");

                lastResult = await scanner.RunScan();

                StopElapsedTimer();
                ShowProgress(false);
                btnCancel.IsEnabled = false;
                btnScan.IsEnabled = true;

                if (lastResult != null)
                {
                    DisplayDiagnosticResults(lastResult);

                    // Update header subtitle with device context
                    var so = lastResult.SystemOverview;
                    lblHeaderSubtitle.Text = $"{so.ComputerName} \u2022 {so.WindowsVersion} \u2022 Last scan {lastResult.Timestamp:h:mm tt}";

                    // Context-sensitive button labels for post-scan state
                    btnScanLabel.Text = "New Scan";
                    btnScanIcon.Text = IcoSync;
                    // Show Before/After comparison if this was a post-optimization rescan
                    if (_preOptimizationSnapshot != null)
                    {
                        var after = new BeforeAfterSnapshot(
                            lastResult.HealthScore,
                            lastResult.Disk.Drives.FirstOrDefault()?.FreeMB ?? 0,
                            lastResult.Ram.PercentUsed,
                            lastResult.Startup.EnabledCount,
                            lastResult.FlaggedIssues.Count);
                        BuildBeforeAfterComparison(_preOptimizationSnapshot, after);
                        _preOptimizationSnapshot = null;
                    }

                    // Save scan snapshot and show "what changed since last scan"
                    var previousSnapshot = ScanHistoryManager.GetPrevious();
                    var currentSnapshot = ScanHistoryManager.CreateSnapshot(lastResult);
                    ScanHistoryManager.Save(currentSnapshot);
                    if (previousSnapshot != null)
                        BuildWhatChangedPanel(previousSnapshot, currentSnapshot);

                    btnExport.Visibility = Visibility.Visible;
                    lblExportText.Text = "Export Scan Report";
                    btnOptimize.Visibility = Visibility.Visible;
                    UpdateStatus("Scan complete", IcoCheck, ThemeManager.BrushSuccess);
                    Log($"✓ Scan complete in {lastResult.ScanDurationSeconds:F1}s! Health Score: {lastResult.HealthScore}/100. {lastResult.FlaggedIssues.Count} issue(s) flagged.");
                    Log("Click 'Export Scan Report' to generate a report.");
                }
                else
                {
                    UpdateStatus("Cancelled", IcoError, ThemeManager.BrushDanger);
                    Log("✗ Scan was cancelled.");
                    btnCancelLabel.Text = "Cancel";
                }
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
            finally
            {
                StopElapsedTimer();
                _isScanning = false;
            }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (lastResult == null) return;

            try
            {
                // Prompt for optional operator notes before generating PDF
                string notes = PromptOperatorNotes();
                if (notes == null) return; // user cancelled
                lastResult.OperatorNotes = notes;

                btnExport.IsEnabled = false;
                btnOptimize.Visibility = Visibility.Collapsed;
                btnScan.IsEnabled = false;
                UpdateStatus("Generating PDF...", IcoPage, ThemeManager.BrushPrimary);
                Log("\r\n=== GENERATING PDF REPORT ===");

                StartElapsedTimer("Generating PDF");

                string pdfPath = await Task.Run(() => GeneratePDFReport(lastResult));

                StopElapsedTimer();
                UpdateStatus("Report saved!", IcoCheck, ThemeManager.BrushSuccess);
                Log($"✓ PDF saved to: {pdfPath}");

                MessageBox.Show(
                    $"Diagnostic report saved!\r\n\r\n{pdfPath}",
                    "Report Generated", MessageBoxButton.OK, MessageBoxImage.Information);

                try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = pdfPath, UseShellExecute = true }); }
                catch (Exception openEx) { Log($"⚠ Could not open PDF viewer: {openEx.Message}"); }
            }
            catch (OperationCanceledException)
            {
                Log("PDF export cancelled.");
                UpdateStatus("Scan complete", IcoCheck, ThemeManager.BrushSuccess);
            }
            catch (Exception ex)
            {
                Log($"✗ PDF Error: {ex.Message}");
                UpdateStatus("PDF Error", IcoWarn, ThemeManager.BrushDanger);
            }
            finally
            {
                StopElapsedTimer();
                btnExport.IsEnabled = true;
                btnOptimize.Visibility = lastResult != null ? Visibility.Visible : Visibility.Collapsed;
                btnScan.IsEnabled = true;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            scanner?.CancelScan();
        }

        private async void BtnOptimize_Click(object sender, RoutedEventArgs e)
        {
            if (lastResult == null) return;

            // If optimization already ran for this scan, show results (review-only, no rescan)
            if (lastResult.OptimizationSummary != null && lastResult.OptimizationActions != null)
            {
                var resultsDialog = new OptimizationResultsWindow(
                    lastResult.OptimizationActions, lastResult.OptimizationSummary,
                    reviewOnly: true) { Owner = this };
                resultsDialog.Show();
                return;
            }

            var optimizer = new Optimizer();
            optimizer.OnLog += Log;

            btnOptimize.Visibility = Visibility.Collapsed;
            UpdateStatus("Building optimization plan...", IcoSync, ThemeManager.BrushPrimary);
            StartElapsedTimer("Building plan");

            var actions = await Task.Run(() => optimizer.BuildPlan(lastResult));

            StopElapsedTimer();

            if (actions.Count == 0)
            {
                UpdateStatus("Scan complete", IcoCheck, ThemeManager.BrushSuccess);
                btnOptimize.Visibility = Visibility.Visible;
                System.Windows.MessageBox.Show(
                    "No actionable optimizations were found based on the current scan results.",
                    "Nothing to Optimize", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            UpdateStatus("Scan complete", IcoCheck, ThemeManager.BrushSuccess);
            btnOptimize.Visibility = Visibility.Visible;

            // Capture pre-optimization snapshot
            var primaryDrive = lastResult.Disk.Drives.FirstOrDefault();
            _preOptimizationSnapshot = new BeforeAfterSnapshot(
                lastResult.HealthScore,
                primaryDrive?.FreeMB ?? 0,
                lastResult.Ram.PercentUsed,
                lastResult.Startup.EnabledCount,
                lastResult.FlaggedIssues.Count);

            var dialog = new OptimizationWindow(actions, optimizer, Log, lastResult) { Owner = this };
            dialog.Closed += (_, _) =>
            {
                optimizer.OnLog -= Log;
                if (dialog.DidRun)
                    lblExportText.Text = "Export Full Report";
                if (dialog.ShouldRescan)
                    BtnScan_Click(btnScan, new RoutedEventArgs());
                else
                    _preOptimizationSnapshot = null;
            };
            dialog.Show();
        }

        // ═══════════════════════════════════════════════════════════════
        //  DISPLAY METHODS
        // ═══════════════════════════════════════════════════════════════

        private void Scanner_OnProgress(ScanProgress progress)
        {
            Dispatcher.BeginInvoke(() =>
            {
                if (!IsLoaded) return;
                _currentScanPhase = progress.CurrentPhase;
                _lastPhaseTotal = progress.Total;
                int pct = progress.Total > 0
                    ? (int)((float)progress.PhaseIndex / progress.Total * 100)
                    : 0;

                SetProgressValue(pct);
                lblProgress.Text = $"{progress.PhaseIndex}/{progress.Total} • {pct}% — {progress.CurrentPhase}";

                SetScanStepCard(progress);
            });
        }

        private void DisplayDiagnosticResults(DiagnosticResult result)
        {
            metricsPanel.Children.Clear();
            metricsScrollViewer.ScrollToTop();

            // Build the health dashboard cards
            BuildHealthDashboard(result);

            // ── Flagged Issues Summary ──
            int critical = result.FlaggedIssues.Count(i => i.Severity == Severity.Critical);
            int warning = result.FlaggedIssues.Count(i => i.Severity == Severity.Warning);
            int info = result.FlaggedIssues.Count(i => i.Severity == Severity.Info);

            metricsPanel.Children.Add(BuildSectionHeader("SCAN SUMMARY", IcoClip));
            metricsPanel.Children.Add(BuildGaugeRow("Health Score", result.HealthScore, "/100", 100,
                result.HealthScore >= 70 ? 0 : result.HealthScore >= 50 ? 1 : 2));
            metricsPanel.Children.Add(BuildInfoRow("Scan Time", result.Timestamp.ToString("h:mm:ss tt")));
            metricsPanel.Children.Add(BuildInfoRow("Duration", $"{result.ScanDurationSeconds}s"));
            metricsPanel.Children.Add(BuildInfoRow("Issues Found", $"{result.FlaggedIssues.Count}"));

            var badgeWrap = new WrapPanel { Margin = new Thickness(4, 4, 0, 4) };
            if (critical > 0)
                badgeWrap.Children.Add(BuildFlagBadge(IcoError, $" {critical} Critical", ThemeManager.CritFg, ThemeManager.CritBg));
            if (warning > 0)
                badgeWrap.Children.Add(BuildFlagBadge(IcoWarn, $" {warning} Warning", ThemeManager.WarnFg, ThemeManager.WarnBg));
            if (info > 0)
                badgeWrap.Children.Add(BuildFlagBadge(IcoInfo, $" {info} Informational", ThemeManager.AccentBlue, ThemeManager.SurfaceAlt));
            if (badgeWrap.Children.Count > 0)
                metricsPanel.Children.Add(badgeWrap);

            // ── Quick-Glance Action Strip ──
            if (critical + warning > 0)
            {
                // Estimate reclaimable disk space
                long reclaimableMB = result.Disk.WindowsTempMB + result.Disk.UserTempMB
                    + result.Disk.PrefetchMB + result.Disk.RecycleBinMB;
                if (result.Disk.WindowsOldExists) reclaimableMB += result.Disk.WindowsOldMB;
                foreach (var b in result.Browser.Browsers)
                    reclaimableMB += b.CacheSizeMB;

                // Estimate fix time (quick heuristic)
                int estMinutes = 1; // base
                if (result.EventLog.BSODs.Count > 0) estMinutes += 8;     // sfc + DISM
                else if (result.EventLog.AppCrashes.Count >= 5) estMinutes += 5; // sfc + re-register
                if (result.EventLog.UnexpectedShutdowns.Count > 0) estMinutes += 1;

                var stripBorder = new Border
                {
                    CornerRadius = new CornerRadius(8),
                    Background = new SolidColorBrush(ThemeManager.SurfaceAlt),
                    BorderBrush = new SolidColorBrush(ThemeManager.Border),
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(14, 10, 14, 10),
                    Margin = new Thickness(0, 6, 0, 2)
                };

                var stripPanel = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Center };

                void AddStripItem(string icon, string label, string value, Color color)
                {
                    var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(8, 2, 8, 2) };
                    sp.Children.Add(MdlIcon(icon, 12, ThemeManager.TextMuted));
                    sp.Children.Add(new TextBlock
                    {
                        Text = $" {label}: ",
                        FontFamily = new FontFamily("Segoe UI"), FontSize = 11.5,
                        Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                        VerticalAlignment = VerticalAlignment.Center
                    });
                    sp.Children.Add(new TextBlock
                    {
                        Text = value,
                        FontFamily = new FontFamily("Segoe UI"), FontSize = 11.5,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = new SolidColorBrush(color),
                        VerticalAlignment = VerticalAlignment.Center
                    });
                    stripPanel.Children.Add(sp);
                }

                // Issue breakdown
                string issueText = "";
                if (critical > 0) issueText += $"{critical} Critical";
                if (warning > 0) issueText += (issueText.Length > 0 ? ", " : "") + $"{warning} Warning";
                AddStripItem(IcoWarn, "Issues", issueText,
                    critical > 0 ? ThemeManager.CritFg : ThemeManager.WarnFg);

                // Estimated time
                AddStripItem("\uE916", "Est. fix time", $"~{estMinutes} min", ThemeManager.AccentBlue);

                // Reclaimable space
                if (reclaimableMB > 100)
                    AddStripItem(IcoDisk, "Reclaimable", $"+{FormatMB(reclaimableMB)}", ThemeManager.AccentGreen);

                stripBorder.Child = stripPanel;
                metricsPanel.Children.Add(stripBorder);
            }

            metricsPanel.Children.Add(BuildDivider());

            // ── System Overview ──
            var so = result.SystemOverview;
            metricsPanel.Children.Add(BuildSectionHeader("SYSTEM OVERVIEW", IcoDevice));
            metricsPanel.Children.Add(BuildInfoRow("Computer", so.ComputerName));
            metricsPanel.Children.Add(BuildInfoRow("Manufacturer", so.Manufacturer));
            metricsPanel.Children.Add(BuildInfoRow("Model", so.Model));
            metricsPanel.Children.Add(BuildInfoRow("Serial", so.SerialNumber));
            metricsPanel.Children.Add(BuildInfoRow("Windows", so.WindowsVersion));
            metricsPanel.Children.Add(BuildInfoRow("Build", so.WindowsBuild));
            metricsPanel.Children.Add(BuildInfoRow("CPU", so.CpuModel));
            metricsPanel.Children.Add(BuildInfoRow("Clock Speed", so.CpuClockSpeed));
            metricsPanel.Children.Add(BuildInfoRow("RAM", FormatMB(so.TotalRamMB)));
            metricsPanel.Children.Add(BuildInfoRow("Uptime", FormatUptime(so.Uptime)));
            metricsPanel.Children.Add(BuildDivider());

            // ── CPU ──
            metricsPanel.Children.Add(BuildSectionHeader("CPU DIAGNOSTICS", IcoCpu));
            metricsPanel.Children.Add(BuildGaugeRow("CPU Load", result.Cpu.CpuLoadPercent, "%", 100,
                RateLevel(result.Cpu.CpuLoadPercent, 15, 50, true)));
            if (result.Cpu.TopCpuProcesses.Count > 0)
            {
                metricsPanel.Children.Add(BuildInfoRow("Top Process",
                    $"{result.Cpu.TopCpuProcesses[0].Name} ({FormatMB(result.Cpu.TopCpuProcesses[0].MemoryMB)})"));
            }
            if (result.Cpu.CpuTemperatureC > 0)
                metricsPanel.Children.Add(BuildGaugeRow("CPU Temp", result.Cpu.CpuTemperatureC, "°C", 100,
                    RateLevel(result.Cpu.CpuTemperatureC, 75, 85, true)));
            else
                metricsPanel.Children.Add(BuildInfoRow("CPU Temp", "Not available"));
            metricsPanel.Children.Add(BuildInfoRow("Throttling", result.Cpu.IsThrottling ? "Yes (!)" : "No"));
            metricsPanel.Children.Add(BuildInfoRow("Fan Status", result.Cpu.FanStatus));
            metricsPanel.Children.Add(BuildDivider());

            // ── RAM ──
            metricsPanel.Children.Add(BuildSectionHeader("RAM DIAGNOSTICS", IcoRam));
            metricsPanel.Children.Add(BuildGaugeRow("RAM Usage", result.Ram.PercentUsed, "%", 100,
                RateLevel(result.Ram.PercentUsed, 70, 85, true)));
            metricsPanel.Children.Add(BuildInfoRow("Total", FormatMB(result.Ram.TotalMB)));
            metricsPanel.Children.Add(BuildInfoRow("Used", FormatMB(result.Ram.UsedMB)));
            metricsPanel.Children.Add(BuildInfoRow("Available", FormatMB(result.Ram.AvailableMB)));
            metricsPanel.Children.Add(BuildInfoRow("Mem Diagnostic", result.Ram.MemoryDiagnosticStatus));
            if (result.Ram.TopRamProcesses.Count > 0)
            {
                var top3 = result.Ram.TopRamProcesses.Take(3);
                foreach (var p in top3)
                    metricsPanel.Children.Add(BuildInfoRow($"  {p.Name}", FormatMB(p.MemoryMB)));
            }
            metricsPanel.Children.Add(BuildDivider());

            // ── GPU ──
            metricsPanel.Children.Add(BuildSectionHeader("GPU DIAGNOSTICS", IcoGpu));

            // Show detected displays
            if (result.Gpu.Displays.Count > 0)
            {
                foreach (var display in result.Gpu.Displays)
                {
                    string label = display.IsPrimary ? "Primary Display" : "Secondary Display";
                    metricsPanel.Children.Add(BuildInfoRow(label, display.MonitorName));
                    metricsPanel.Children.Add(BuildInfoRow("  Resolution", $"{display.Resolution} @ {display.RefreshRateText}"));
                }
            }

            // Show all GPUs with full details
            if (result.Gpu.AllGpus.Count > 0)
            {
                foreach (var gpuInfo in result.Gpu.AllGpus)
                {
                    string label = gpuInfo.IsPrimary ? "Primary GPU" : "GPU";
                    metricsPanel.Children.Add(BuildInfoRow(label, gpuInfo.Name));
                    metricsPanel.Children.Add(BuildInfoRow("  Driver", gpuInfo.DriverVersion));
                    metricsPanel.Children.Add(BuildInfoRow("  Driver Date", gpuInfo.DriverDate + (gpuInfo.DriverOutdated ? " (outdated!)" : "")));
                    metricsPanel.Children.Add(BuildInfoRow("  Status", gpuInfo.AdapterStatus));
                    if (gpuInfo.DedicatedVideoMemoryMB > 0)
                        metricsPanel.Children.Add(BuildInfoRow("  Dedicated VRAM", FormatMB(gpuInfo.DedicatedVideoMemoryMB)));
                    if (gpuInfo.IsPrimary)
                    {
                        if (result.Gpu.GpuTemperatureC > 0)
                            metricsPanel.Children.Add(BuildGaugeRow("  GPU Temp", result.Gpu.GpuTemperatureC, "°C", 100,
                                RateLevel(result.Gpu.GpuTemperatureC, 75, 85, true)));
                        else
                            metricsPanel.Children.Add(BuildInfoRow("  GPU Temp", "Not available"));
                        if (result.Gpu.GpuUsagePercent > 0)
                            metricsPanel.Children.Add(BuildGaugeRow("  GPU Usage", result.Gpu.GpuUsagePercent, "%", 100,
                                RateLevel(result.Gpu.GpuUsagePercent, 80, 90, true)));
                    }
                }
            }
            else
            {
                // Fallback: single GPU display
                metricsPanel.Children.Add(BuildInfoRow("GPU", result.Gpu.GpuName));
                metricsPanel.Children.Add(BuildInfoRow("Driver", result.Gpu.DriverVersion));
                metricsPanel.Children.Add(BuildInfoRow("Driver Date", result.Gpu.DriverDate + (result.Gpu.DriverOutdated ? " (outdated!)" : "")));
                metricsPanel.Children.Add(BuildInfoRow("Status", result.Gpu.AdapterStatus));
                if (result.Gpu.DedicatedVideoMemoryMB > 0)
                    metricsPanel.Children.Add(BuildInfoRow("Dedicated VRAM", FormatMB(result.Gpu.DedicatedVideoMemoryMB)));
                metricsPanel.Children.Add(BuildInfoRow("Resolution", result.Gpu.Resolution));
                metricsPanel.Children.Add(BuildInfoRow("Refresh Rate", result.Gpu.RefreshRate));
                if (result.Gpu.GpuTemperatureC > 0)
                    metricsPanel.Children.Add(BuildGaugeRow("GPU Temp", result.Gpu.GpuTemperatureC, "°C", 100,
                        RateLevel(result.Gpu.GpuTemperatureC, 75, 85, true)));
                else
                    metricsPanel.Children.Add(BuildInfoRow("GPU Temp", "Not available"));
                if (result.Gpu.GpuUsagePercent > 0)
                    metricsPanel.Children.Add(BuildGaugeRow("GPU Usage", result.Gpu.GpuUsagePercent, "%", 100,
                        RateLevel(result.Gpu.GpuUsagePercent, 80, 90, true)));
                foreach (var adapter in result.Gpu.AdditionalAdapters)
                    metricsPanel.Children.Add(BuildInfoRow("  + Adapter", adapter));
            }
            metricsPanel.Children.Add(BuildDivider());

            // ── Disk ──
            metricsPanel.Children.Add(BuildSectionHeader("DISK DIAGNOSTICS", IcoDisk));
            foreach (var drive in result.Disk.Drives)
            {
                metricsPanel.Children.Add(BuildGaugeRow($"{drive.DriveLetter} Used", drive.PercentUsed, "%", 100,
                    RateLevel(drive.PercentUsed, 75, 85, true)));
                metricsPanel.Children.Add(BuildInfoRow($"  Free", $"{FormatMB(drive.FreeMB)} / {FormatMB(drive.TotalMB)}"));
                metricsPanel.Children.Add(BuildInfoRow($"  Type", drive.DriveType));
                metricsPanel.Children.Add(BuildInfoRow($"  Health", drive.HealthStatus));
            }
            if (result.Disk.DiskActivityPercent > 0)
                metricsPanel.Children.Add(BuildGaugeRow("Disk Activity", result.Disk.DiskActivityPercent, "%", 100,
                    RateLevel(result.Disk.DiskActivityPercent, 50, 80, true)));
            metricsPanel.Children.Add(BuildInfoRow("Win Temp", FormatMB(result.Disk.WindowsTempMB)));
            metricsPanel.Children.Add(BuildInfoRow("User Temp", FormatMB(result.Disk.UserTempMB)));
            metricsPanel.Children.Add(BuildInfoRow("Prefetch", FormatMB(result.Disk.PrefetchMB)));
            metricsPanel.Children.Add(BuildInfoRow("Recycle Bin", FormatMB(result.Disk.RecycleBinMB)));
            metricsPanel.Children.Add(BuildInfoRow("SW Dist.", FormatMB(result.Disk.SoftwareDistributionMB)));
            if (result.Disk.UpgradeLogsMB > 0)
                metricsPanel.Children.Add(BuildInfoRow("Upgrade Logs", FormatMB(result.Disk.UpgradeLogsMB)));
            if (result.Disk.WindowsOldExists)
                metricsPanel.Children.Add(BuildInfoRow("Windows.old", FormatMB(result.Disk.WindowsOldMB)));
            metricsPanel.Children.Add(BuildDivider());

            // ── Battery ──
            if (result.Battery.HasBattery)
            {
                metricsPanel.Children.Add(BuildSectionHeader("BATTERY HEALTH", IcoBattery));
                metricsPanel.Children.Add(BuildGaugeRow("Battery Health", result.Battery.HealthPercent, "%", 100,
                    RateLevel(100 - result.Battery.HealthPercent, 30, 50, true)));
                metricsPanel.Children.Add(BuildInfoRow("Power Source", result.Battery.PowerSource));
                metricsPanel.Children.Add(BuildInfoRow("Power Plan", result.Battery.PowerPlan));
                if (result.Battery.DesignCapacityMWh > 0)
                    metricsPanel.Children.Add(BuildInfoRow("Design Cap.", $"{result.Battery.DesignCapacityMWh} mWh"));
                if (result.Battery.FullChargeCapacityMWh > 0)
                    metricsPanel.Children.Add(BuildInfoRow("Full Charge", $"{result.Battery.FullChargeCapacityMWh} mWh"));
                metricsPanel.Children.Add(BuildDivider());
            }

            // ── Startup ──
            metricsPanel.Children.Add(BuildSectionHeader("STARTUP PROGRAMS", IcoStartup));
            metricsPanel.Children.Add(BuildInfoRow("Enabled", $"{result.Startup.EnabledCount}"));
            metricsPanel.Children.Add(BuildInfoRow("Total", $"{result.Startup.Entries.Count}"));
            foreach (var entry in result.Startup.Entries.Take(8))
                metricsPanel.Children.Add(BuildInfoRow($"  {entry.Name}", entry.Enabled ? "Enabled" : "Disabled"));
            if (result.Startup.Entries.Count > 8)
                metricsPanel.Children.Add(BuildInfoRow("", $"... and {result.Startup.Entries.Count - 8} more"));
            metricsPanel.Children.Add(BuildDivider());

            // ── Visual Settings ──
            metricsPanel.Children.Add(BuildSectionHeader("VISUAL SETTINGS", IcoPaint));
            metricsPanel.Children.Add(BuildInfoRow("Effects", result.VisualSettings.VisualEffectsSetting));
            metricsPanel.Children.Add(BuildInfoRow("Transparency", result.VisualSettings.TransparencyEnabled ? "Enabled" : "Disabled"));
            metricsPanel.Children.Add(BuildInfoRow("Animations", result.VisualSettings.AnimationsEnabled ? "Enabled" : "Disabled"));
            metricsPanel.Children.Add(BuildDivider());

            // ── Network ──
            metricsPanel.Children.Add(BuildSectionHeader("NETWORK", IcoGlobe));
            metricsPanel.Children.Add(BuildInfoRow("Connection", result.Network.ConnectionType));
            if (result.Network.ConnectionType == "WiFi")
            {
                metricsPanel.Children.Add(BuildInfoRow("WiFi Band", result.Network.WifiBand));
                metricsPanel.Children.Add(BuildInfoRow("Signal", result.Network.WifiSignalStrength));
                if (result.Network.LinkSpeed != "Unknown")
                    metricsPanel.Children.Add(BuildInfoRow("Link Speed", result.Network.LinkSpeed));
            }
            metricsPanel.Children.Add(BuildInfoRow("Adapter Speed", result.Network.AdapterSpeed));
            metricsPanel.Children.Add(BuildColoredInfoRow("DNS Response", result.Network.DnsResponseTime, RateDns(result.Network.DnsResponseTime)));
            metricsPanel.Children.Add(BuildColoredInfoRow("Ping Latency", result.Network.PingLatency, RatePing(result.Network.PingLatency)));
            metricsPanel.Children.Add(BuildInfoRow("VPN Active", result.Network.VpnActive ? $"Yes — {result.Network.VpnClient}" : "No"));

            // Internet Speed Test
            if (result.Network.DownloadSpeedMbps > 0 || result.Network.UploadSpeedMbps > 0)
            {
                metricsPanel.Children.Add(BuildGaugeRow("Download Speed",
                    result.Network.DownloadSpeedMbps, " Mbps", 500,
                    result.Network.DownloadSpeedMbps < 10 ? 2 : result.Network.DownloadSpeedMbps < 50 ? 1 : 0));
                metricsPanel.Children.Add(BuildGaugeRow("Upload Speed",
                    result.Network.UploadSpeedMbps, " Mbps", 200,
                    result.Network.UploadSpeedMbps < 5 ? 2 : result.Network.UploadSpeedMbps < 20 ? 1 : 0));
            }
            else if (!string.IsNullOrEmpty(result.Network.SpeedTestError))
            {
                metricsPanel.Children.Add(BuildInfoRow("Speed Test", $"Failed — {result.Network.SpeedTestError}"));
            }
            else
            {
                metricsPanel.Children.Add(BuildInfoRow("Speed Test", "Not available"));
            }
            metricsPanel.Children.Add(BuildDivider());

            // ── Network Drives ──
            if (result.NetworkDrives.Drives.Count > 0)
            {
                metricsPanel.Children.Add(BuildSectionHeader("NETWORK DRIVES", IcoGlobe));
                foreach (var nd in result.NetworkDrives.Drives)
                {
                    string status = nd.IsAccessible ? $"OK ({nd.LatencyMs}ms)" : "Unreachable";
                    int statusLevel = !nd.IsAccessible ? 2 : nd.LatencyFlagged ? 1 : 0;
                    metricsPanel.Children.Add(BuildColoredInfoRow(
                        $"{nd.DriveLetter} ({nd.UncPath})", status, statusLevel));

                    if (nd.IsAccessible)
                    {
                        metricsPanel.Children.Add(BuildGaugeRow($"  {nd.DriveLetter} Usage", nd.PercentUsed, "%", 100,
                            RateLevel(nd.PercentUsed, 80, 90, true)));
                        metricsPanel.Children.Add(BuildInfoRow("  Free Space", FormatMB(nd.FreeMB)));
                        if (nd.OfflineFilesEnabled)
                            metricsPanel.Children.Add(BuildInfoRow("  Offline Files", nd.SyncStatus));
                    }
                }
                metricsPanel.Children.Add(BuildDivider());
            }

            // ── Antivirus ──
            metricsPanel.Children.Add(BuildSectionHeader("SECURITY", IcoShield));
            metricsPanel.Children.Add(BuildInfoRow("Antivirus", result.Antivirus.AntivirusName));
            metricsPanel.Children.Add(BuildInfoRow("Full Scan", result.Antivirus.FullScanRunning ? "Running" : "Not running"));
            metricsPanel.Children.Add(BuildInfoRow("BitLocker", result.Antivirus.BitLockerStatus));
            metricsPanel.Children.Add(BuildDivider());

            // ── Windows Update ──
            metricsPanel.Children.Add(BuildSectionHeader("WINDOWS UPDATE", IcoSync));
            metricsPanel.Children.Add(BuildInfoRow("Status", result.WindowsUpdate.PendingUpdates));

            // Show service status with start type for clarity (e.g. "Stopped (Disabled)")
            string wuService = result.WindowsUpdate.UpdateServiceStatus;
            string wuStartType = result.WindowsUpdate.ServiceStartType;
            if (!string.IsNullOrEmpty(wuStartType) && wuStartType != "Unknown")
                wuService = $"{wuService} ({wuStartType})";
            metricsPanel.Children.Add(BuildInfoRow("Service", wuService));

            metricsPanel.Children.Add(BuildInfoRow("Cache Size", FormatMB(result.WindowsUpdate.UpdateCacheMB)));
            metricsPanel.Children.Add(BuildDivider());

            // ── Outlook ──
            if (result.Outlook.OutlookInstalled)
            {
                metricsPanel.Children.Add(BuildSectionHeader("OUTLOOK", IcoMail));
                foreach (var df in result.Outlook.DataFiles)
                    metricsPanel.Children.Add(BuildInfoRow(
                        System.IO.Path.GetFileName(df.Path),
                        FormatMB(df.SizeMB) + (df.SizeFlagged ? " (!)" : "")));
                metricsPanel.Children.Add(BuildInfoRow("Add-ins", $"{result.Outlook.AddInCount}"));
                metricsPanel.Children.Add(BuildDivider());
            }

            // ── Browser ──
            if (result.Browser.Browsers.Count > 0)
            {
                metricsPanel.Children.Add(BuildSectionHeader("BROWSERS", IcoBrowser));
                foreach (var b in result.Browser.Browsers)
                {
                    metricsPanel.Children.Add(BuildInfoRow(b.Name, $"{b.OpenTabs} tabs, {FormatMB(b.CacheSizeMB)} cache, {b.ExtensionCount} extensions"));
                }
                metricsPanel.Children.Add(BuildDivider());
            }

            // ── User Profile ──
            metricsPanel.Children.Add(BuildSectionHeader("USER PROFILE", IcoContact));
            metricsPanel.Children.Add(BuildInfoRow("Profile Size", FormatMB(result.UserProfile.ProfileSizeMB)));
            metricsPanel.Children.Add(BuildInfoRow("Profile Age", result.UserProfile.ProfileAge));
            metricsPanel.Children.Add(BuildInfoRow("Desktop Items", $"{result.UserProfile.DesktopItemCount}"));
            metricsPanel.Children.Add(BuildDivider());

            // ── Office ──
            if (result.Office.OfficeInstalled)
            {
                metricsPanel.Children.Add(BuildSectionHeader("OFFICE", IcoOffice));
                metricsPanel.Children.Add(BuildInfoRow("Version", result.Office.OfficeVersion));
                metricsPanel.Children.Add(BuildInfoRow("Repair Needed", result.Office.RepairNeeded ? "Yes (!)" : "No"));
                metricsPanel.Children.Add(BuildDivider());
            }

            // ── Installed Software ──
            metricsPanel.Children.Add(BuildSectionHeader("INSTALLED SOFTWARE", IcoApps));
            metricsPanel.Children.Add(BuildInfoRow("Total Apps", $"{result.InstalledSoftware.TotalCount}"));
            metricsPanel.Children.Add(BuildInfoRow("VC++ Runtimes", $"{result.InstalledSoftware.RuntimeCount}"));
            foreach (var rt in result.InstalledSoftware.RuntimeApps)
                metricsPanel.Children.Add(BuildInfoRow($"  {rt.Name}", rt.Version));
            metricsPanel.Children.Add(BuildInfoRow("EOL Software", result.InstalledSoftware.EOLApps.Count > 0
                ? $"{result.InstalledSoftware.EOLApps.Count} found (!)"
                : "None detected"));
            metricsPanel.Children.Add(BuildInfoRow("Bloatware", result.InstalledSoftware.BloatwareApps.Count > 0
                ? $"{result.InstalledSoftware.BloatwareApps.Count} found"
                : "None detected"));
            foreach (var eol in result.InstalledSoftware.EOLApps.Take(5))
                metricsPanel.Children.Add(BuildInfoRow($"  ⚠ {eol.Name}", eol.Version));
            foreach (var bloat in result.InstalledSoftware.BloatwareApps.Take(5))
                metricsPanel.Children.Add(BuildInfoRow($"  ⓘ {bloat.Name}", bloat.Version));
            metricsPanel.Children.Add(BuildDivider());

            // ── Event Log ──
            {
                DateTime? lastRemediation = Optimizer.GetRemediationTimestamp();
                bool hasEvents = result.EventLog.TotalEventCount > 0;
                bool wasRemediated = lastRemediation.HasValue;

                metricsPanel.Children.Add(BuildSectionHeader("EVENT LOG ANALYSIS (30 DAYS)", IcoPage));

                if (hasEvents)
                {
                    metricsPanel.Children.Add(BuildColoredInfoRow("BSODs", $"{result.EventLog.BSODs.Count}",
                        result.EventLog.BSODs.Count > 0 ? 2 : 0));
                    metricsPanel.Children.Add(BuildColoredInfoRow("Unexpected Shutdowns", $"{result.EventLog.UnexpectedShutdowns.Count}",
                        result.EventLog.UnexpectedShutdowns.Count >= 3 ? 2 : result.EventLog.UnexpectedShutdowns.Count > 0 ? 1 : 0));
                    metricsPanel.Children.Add(BuildColoredInfoRow("Disk Errors", $"{result.EventLog.DiskErrors.Count}",
                        result.EventLog.DiskErrors.Count >= 5 ? 2 : result.EventLog.DiskErrors.Count > 0 ? 1 : 0));
                    metricsPanel.Children.Add(BuildColoredInfoRow("App Crashes", $"{result.EventLog.AppCrashes.Count}",
                        result.EventLog.AppCrashes.Count >= 5 ? 1 : 0));

                    // Show recent entries with timestamps
                    void ShowRecentEntries(string label, List<EventLogEntry> entries, int max)
                    {
                        foreach (var entry in entries.Take(max))
                        {
                            string msg = !string.IsNullOrEmpty(entry.Message) ? entry.Message : entry.Source;
                            metricsPanel.Children.Add(BuildInfoRow(
                                $"  {entry.Timestamp:MM/dd h:mm tt}", msg));
                        }
                    }

                    if (result.EventLog.BSODs.Count > 0)
                        ShowRecentEntries("BSOD", result.EventLog.BSODs, 3);
                    if (result.EventLog.UnexpectedShutdowns.Count > 0)
                        ShowRecentEntries("Shutdown", result.EventLog.UnexpectedShutdowns, 3);
                    if (result.EventLog.DiskErrors.Count > 0)
                        ShowRecentEntries("Disk Error", result.EventLog.DiskErrors, 3);
                    if (result.EventLog.AppCrashes.Count > 0)
                        ShowRecentEntries("App Crash", result.EventLog.AppCrashes, 3);
                }
                else
                {
                    metricsPanel.Children.Add(BuildColoredInfoRow("BSODs", "0", 0));
                    metricsPanel.Children.Add(BuildColoredInfoRow("Unexpected Shutdowns", "0", 0));
                    metricsPanel.Children.Add(BuildColoredInfoRow("Disk Errors", "0", 0));
                    metricsPanel.Children.Add(BuildColoredInfoRow("App Crashes", "0", 0));

                    if (wasRemediated)
                        metricsPanel.Children.Add(BuildColoredInfoRow("Status",
                            $"All clear — remediated on {lastRemediation.Value.ToLocalTime():MMM dd} at {lastRemediation.Value.ToLocalTime():h:mm tt}", 0));
                    else
                        metricsPanel.Children.Add(BuildColoredInfoRow("Status", "No issues found", 0));
                }

                metricsPanel.Children.Add(BuildDivider());
            }

            // ── Flagged Issues ──
            if (result.FlaggedIssues.Count > 0)
            {
                metricsPanel.Children.Add(BuildSectionHeader("FLAGGED ISSUES", IcoWarn));
                foreach (var issue in result.FlaggedIssues)
                {
                    string iconGlyph = issue.Severity switch
                    {
                        Severity.Critical => IcoError,
                        Severity.Warning => IcoWarn,
                        _ => IcoInfo
                    };
                    Color fg = issue.Severity switch
                    {
                        Severity.Critical => ThemeManager.CritFg,
                        Severity.Warning => ThemeManager.WarnFg,
                        _ => ThemeManager.AccentBlue
                    };
                    Color bg = issue.Severity switch
                    {
                        Severity.Critical => ThemeManager.CritBg,
                        Severity.Warning => ThemeManager.WarnBg,
                        _ => ThemeManager.SurfaceAlt
                    };

                    var card = new Border
                    {
                        CornerRadius = new CornerRadius(8),
                        Background = new SolidColorBrush(bg),
                        BorderBrush = new SolidColorBrush(Color.FromArgb(25, fg.R, fg.G, fg.B)),
                        BorderThickness = new Thickness(1),
                        Padding = new Thickness(0),
                        Margin = new Thickness(0, 3, 0, 3),
                        ClipToBounds = true
                    };
                    var cardGrid = new Grid();
                    cardGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(4) });
                    cardGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    // Left accent bar
                    var accentBar = new Border
                    {
                        Background = new SolidColorBrush(fg),
                        CornerRadius = new CornerRadius(8, 0, 0, 8)
                    };
                    Grid.SetColumn(accentBar, 0);
                    cardGrid.Children.Add(accentBar);

                    var sp = new StackPanel { Margin = new Thickness(12, 8, 12, 8) };
                    Grid.SetColumn(sp, 1);

                    var header = new StackPanel { Orientation = Orientation.Horizontal };
                    header.Children.Add(MdlIcon(iconGlyph, 12, fg));
                    header.Children.Add(new TextBlock
                    {
                        Text = $" [{issue.Category}] {issue.Description}",
                        FontFamily = new FontFamily("Segoe UI"), FontSize = 12,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = new SolidColorBrush(fg),
                        TextWrapping = TextWrapping.Wrap,
                        VerticalAlignment = VerticalAlignment.Center
                    });
                    sp.Children.Add(header);

                    if (!string.IsNullOrEmpty(issue.Recommendation))
                    {
                        sp.Children.Add(new TextBlock
                        {
                            Text = $"→ {issue.Recommendation}",
                            FontFamily = new FontFamily("Segoe UI"), FontSize = 11,
                            Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                            TextWrapping = TextWrapping.Wrap,
                            Margin = new Thickness(22, 3, 0, 0)
                        });
                    }
                    cardGrid.Children.Add(sp);
                    card.Child = cardGrid;
                    metricsPanel.Children.Add(card);
                }
            }
            else
            {
                metricsPanel.Children.Add(BuildSectionHeader("SYSTEM STATUS", IcoCheck));
                metricsPanel.Children.Add(BuildFlagBadge(IcoCheck, " All Clear — No issues detected", ThemeManager.GoodFg, ThemeManager.GoodBg));
            }

            ShowRichMetrics();

            // Log flagged issues
            Log("\r\n=== FLAGGED ISSUES ===");
            foreach (var issue in result.FlaggedIssues)
                Log($"  [{issue.Severity}] {issue.Category}: {issue.Description}");
        }

        // ═══════════════════════════════════════════════════════════════
        //  RICH METRICS BUILDERS
        // ═══════════════════════════════════════════════════════════════

        private static FrameworkElement BuildSectionHeader(string title, string icon)
        {
            var container = new Border
            {
                CornerRadius = new CornerRadius(8),
                Background = new SolidColorBrush(ThemeManager.SurfaceAlt),
                Padding = new Thickness(0),
                Margin = new Thickness(0, 10, 0, 8),
                ClipToBounds = true
            };
            var outerGrid = new Grid();
            outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(3) });
            outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left brand accent stripe
            var stripe = new Border
            {
                Background = new SolidColorBrush(ThemeManager.AccentBlue),
                CornerRadius = new CornerRadius(8, 0, 0, 8)
            };
            Grid.SetColumn(stripe, 0);
            outerGrid.Children.Add(stripe);

            var sp = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(12, 8, 12, 8)
            };
            sp.Children.Add(MdlIcon(icon, 14, ThemeManager.AccentBlue));
            sp.Children.Add(new TextBlock
            {
                Text = title,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(ThemeManager.TextPrimary),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 0, 0)
            });
            Grid.SetColumn(sp, 1);
            outerGrid.Children.Add(sp);

            container.Child = outerGrid;
            return container;
        }

        private static FrameworkElement BuildDivider()
        {
            return new Border
            {
                Height = 1,
                Background = new SolidColorBrush(ThemeManager.Border),
                Opacity = 0.5,
                Margin = new Thickness(0, 4, 0, 4)
            };
        }

        private static FrameworkElement BuildFlagBadge(string iconGlyph, string text, Color fg, Color bg)
        {
            var badge = new Border
            {
                CornerRadius = new CornerRadius(12),
                Background = new SolidColorBrush(bg),
                Padding = new Thickness(10, 5, 12, 5),
                Margin = new Thickness(0, 3, 8, 3),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            sp.Children.Add(MdlIcon(iconGlyph, 10, fg));
            sp.Children.Add(new TextBlock
            {
                Text = text,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(fg),
                VerticalAlignment = VerticalAlignment.Center
            });
            badge.Child = sp;
            return badge;
        }

        private static int RateLevel(double value, double warnThreshold, double critThreshold, bool lowerIsBetter)
        {
            bool isCrit = lowerIsBetter ? value >= critThreshold : value <= critThreshold;
            bool isWarn = lowerIsBetter ? value >= warnThreshold : value <= warnThreshold;
            if (isCrit) return 2;
            if (isWarn) return 1;
            return 0;
        }

        private static FrameworkElement BuildGaugeRow(string label, double value, string unit, double max, int level)
        {
            Color barColor = level switch
            {
                2 => ThemeManager.CritFg,
                1 => ThemeManager.WarnFg,
                _ => ThemeManager.GoodFg
            };
            Color barBg = level switch
            {
                2 => ThemeManager.CritBg,
                1 => ThemeManager.WarnBg,
                _ => ThemeManager.GoodBg
            };
            string badgeText = level switch { 2 => "CRITICAL", 1 => "WARNING", _ => "OK" };
            Color badgeFg = barColor;
            Color badgeBgColor = barBg;

            var row = new Grid { Margin = new Thickness(4, 5, 4, 5) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var lbl = new TextBlock
            {
                Text = label,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(lbl, 0);
            row.Children.Add(lbl);

            var trackBorder = new Border
            {
                Height = 8,
                CornerRadius = new CornerRadius(4),
                Background = new SolidColorBrush(barBg),
                ClipToBounds = true,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };
            double fraction = Math.Clamp(value / max, 0, 1);
            var fillBorder = new Border
            {
                Height = 8,
                CornerRadius = new CornerRadius(4),
                HorizontalAlignment = HorizontalAlignment.Left,
                Background = new SolidColorBrush(barColor)
            };
            fillBorder.Loaded += (s, e) =>
            {
                var parent = fillBorder.Parent as FrameworkElement;
                double targetW = (parent?.ActualWidth ?? 200) * fraction;
                var anim = new DoubleAnimation
                {
                    From = 0,
                    To = targetW,
                    Duration = TimeSpan.FromMilliseconds(700),
                    EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut },
                    BeginTime = TimeSpan.FromMilliseconds(150)
                };
                fillBorder.BeginAnimation(WidthProperty, anim);
            };
            trackBorder.SizeChanged += (s, e) =>
            {
                double targetW = e.NewSize.Width * fraction;
                fillBorder.BeginAnimation(WidthProperty, null);
                fillBorder.Width = targetW;
            };
            trackBorder.Child = fillBorder;
            Grid.SetColumn(trackBorder, 1);
            row.Children.Add(trackBorder);

            var valText = new TextBlock
            {
                Text = $"{value:F1}{unit}",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(ThemeManager.TextPrimary),
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Right
            };
            Grid.SetColumn(valText, 2);
            row.Children.Add(valText);

            var badge = new Border
            {
                CornerRadius = new CornerRadius(4),
                Background = new SolidColorBrush(badgeBgColor),
                Padding = new Thickness(8, 2, 8, 2),
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            badge.Child = new TextBlock
            {
                Text = badgeText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(badgeFg)
            };
            Grid.SetColumn(badge, 3);
            row.Children.Add(badge);

            return row;
        }

        private static FrameworkElement BuildInfoRow(string label, string value)
        {
            var row = new Grid { Margin = new Thickness(4, 3, 4, 3) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var lbl = new TextBlock
            {
                Text = label,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(lbl, 0);
            row.Children.Add(lbl);

            var val = new TextBlock
            {
                Text = value,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(ThemeManager.TextPrimary),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetColumn(val, 1);
            row.Children.Add(val);

            return row;
        }

        private static FrameworkElement BuildColoredInfoRow(string label, string value, int level)
        {
            Color fg = level switch
            {
                2 => ThemeManager.CritFg,
                1 => ThemeManager.WarnFg,
                _ => ThemeManager.GoodFg
            };
            Color bg = level switch
            {
                2 => ThemeManager.CritBg,
                1 => ThemeManager.WarnBg,
                _ => ThemeManager.GoodBg
            };
            string badgeText = level switch { 2 => "POOR", 1 => "FAIR", _ => "GOOD" };

            var row = new Grid { Margin = new Thickness(4, 3, 4, 3) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var lbl = new TextBlock
            {
                Text = label,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemeManager.TextMuted),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(lbl, 0);
            row.Children.Add(lbl);

            var val = new TextBlock
            {
                Text = value,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(fg),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(val, 1);
            row.Children.Add(val);

            var badge = new Border
            {
                CornerRadius = new CornerRadius(4),
                Background = new SolidColorBrush(bg),
                Padding = new Thickness(8, 2, 8, 2),
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            badge.Child = new TextBlock
            {
                Text = badgeText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(fg)
            };
            Grid.SetColumn(badge, 2);
            row.Children.Add(badge);

            return row;
        }

        private static int RateDns(string dnsResponse) => DiagnosticFormatters.RateDns(dnsResponse);
        private static int RatePing(string pingLatency) => DiagnosticFormatters.RatePing(pingLatency);
        private static string FormatUptime(TimeSpan uptime) => DiagnosticFormatters.FormatUptime(uptime);
        private static string FormatMB(long mb) => DiagnosticFormatters.FormatMB(mb);

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

        private static Path BuildOrbitArc(double cx, double cy, double r, double startDeg, double sweepDeg, double stroke, Color color)
        {
            double startRad = (startDeg - 90) * Math.PI / 180.0;
            double endRad = (startDeg + sweepDeg - 90) * Math.PI / 180.0;
            var start = new Point(cx + r * Math.Cos(startRad), cy + r * Math.Sin(startRad));
            var end = new Point(cx + r * Math.Cos(endRad), cy + r * Math.Sin(endRad));

            var figure = new PathFigure { StartPoint = start, IsClosed = false };
            figure.Segments.Add(new ArcSegment
            {
                Point = end,
                Size = new Size(r, r),
                SweepDirection = SweepDirection.Clockwise,
                IsLargeArc = sweepDeg > 180
            });

            var geo = new PathGeometry();
            geo.Figures.Add(figure);
            return new Path
            {
                Data = geo,
                Stroke = new SolidColorBrush(color),
                StrokeThickness = stroke,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round
            };
        }

        private void SetScanStepCard(ScanProgress progress)
        {
            txtLiveScan.Children.Clear();
            txtLiveScan.VerticalAlignment = VerticalAlignment.Center;

            var accent = ThemeManager.AccentBlue;

            // ── Clean spinner matching reference design ──
            const double r = 58;
            const double stroke = 3.5;
            const double size = (r + stroke + 2) * 2;
            double cx = size / 2, cy = size / 2;

            var ring = new Grid
            {
                Width = size, Height = size,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 22)
            };

            // Full circle track — very light
            ring.Children.Add(new Ellipse
            {
                Width = r * 2, Height = r * 2,
                Stroke = new SolidColorBrush(Color.FromArgb(28, accent.R, accent.G, accent.B)),
                StrokeThickness = stroke,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            });

            // Comet-tail arcs — 4 segments at decreasing opacity, spinning together
            var spinCanvas = new Canvas
            {
                Width = size, Height = size,
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new RotateTransform(0)
            };

            // Head: 90° at full color
            var head = BuildOrbitArc(cx, cy, r, 0, 90, stroke, accent);
            head.Opacity = 0.85;
            spinCanvas.Children.Add(head);

            // Body: 80° continuing after head
            var body = BuildOrbitArc(cx, cy, r, 95, 80, stroke, accent);
            body.Opacity = 0.45;
            spinCanvas.Children.Add(body);

            // Tail: 70° fading out
            var tail = BuildOrbitArc(cx, cy, r, 180, 70, stroke, accent);
            tail.Opacity = 0.20;
            spinCanvas.Children.Add(tail);

            // Wisp: 40° barely visible
            var wisp = BuildOrbitArc(cx, cy, r, 256, 40, stroke, accent);
            wisp.Opacity = 0.08;
            spinCanvas.Children.Add(wisp);

            var spin = new DoubleAnimation
            {
                From = 0, To = 360,
                Duration = TimeSpan.FromSeconds(1.8),
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = new PowerEase { EasingMode = EasingMode.EaseInOut, Power = 1.3 }
            };
            spinCanvas.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, spin);
            ring.Children.Add(spinCanvas);

            // Laptop icon + scan line — centered inside ring
            const double iconArea = 48;
            var iconContainer = new Grid
            {
                Width = iconArea, Height = iconArea,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                ClipToBounds = true
            };

            // Laptop icon — larger, centered
            var icon = MdlIcon(IcoDevice, 44, ThemeManager.TextPrimary);
            icon.HorizontalAlignment = HorizontalAlignment.Center;
            icon.VerticalAlignment = VerticalAlignment.Center;
            icon.Opacity = 0.60;
            iconContainer.Children.Add(icon);

            // Horizontal scan line — sweeps up and down over the laptop
            var scanLine = new Border
            {
                Width = iconArea * 0.85,
                Height = 3.5,
                CornerRadius = new CornerRadius(2),
                Background = new SolidColorBrush(accent),
                Opacity = 0.75,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
                RenderTransform = new TranslateTransform(0, 0)
            };
            var scanSweep = new DoubleAnimation
            {
                From = 6,
                To = iconArea - 10,
                Duration = TimeSpan.FromMilliseconds(1400),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
            };
            ((TranslateTransform)scanLine.RenderTransform).BeginAnimation(TranslateTransform.YProperty, scanSweep);
            iconContainer.Children.Add(scanLine);

            ring.Children.Add(iconContainer);

            txtLiveScan.Children.Add(ring);

            // Phase name
            txtLiveScan.Children.Add(new TextBlock
            {
                Text = progress.CurrentPhase,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = (SolidColorBrush)Application.Current.FindResource("TextDark"),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 4)
            });

            // Phase description
            if (!string.IsNullOrEmpty(progress.PhaseDescription))
            {
                txtLiveScan.Children.Add(new TextBlock
                {
                    Text = progress.PhaseDescription,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 12,
                    Foreground = (SolidColorBrush)Application.Current.FindResource("TextMuted"),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 16)
                });
            }

            // Step dots — connected by thin lines, current pulses
            var dotsPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 12)
            };
            for (int i = 1; i <= progress.Total; i++)
            {
                // Connector line before each dot (except first)
                if (i > 1)
                {
                    Color lineColor = i <= progress.PhaseIndex ? ThemeManager.GoodFg : ThemeManager.Border;
                    dotsPanel.Children.Add(new Border
                    {
                        Width = 12, Height = 2,
                        Background = new SolidColorBrush(lineColor),
                        VerticalAlignment = VerticalAlignment.Center,
                        CornerRadius = new CornerRadius(1),
                        Opacity = i <= progress.PhaseIndex ? 0.6 : 0.3
                    });
                }

                Color dotColor;
                double dotSize;
                if (i < progress.PhaseIndex)
                {
                    dotColor = ThemeManager.GoodFg;
                    dotSize = 8;
                }
                else if (i == progress.PhaseIndex)
                {
                    dotColor = accent;
                    dotSize = 10;
                }
                else
                {
                    dotColor = ThemeManager.Border;
                    dotSize = 6;
                }

                var dot = new Ellipse
                {
                    Width = dotSize, Height = dotSize,
                    Fill = new SolidColorBrush(dotColor),
                    Margin = new Thickness(2, 0, 2, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    RenderTransformOrigin = new Point(0.5, 0.5),
                    RenderTransform = new ScaleTransform(1, 1)
                };

                if (i == progress.PhaseIndex)
                {
                    var pulse = new DoubleAnimation
                    {
                        From = 0.8, To = 1.3,
                        Duration = TimeSpan.FromMilliseconds(900),
                        AutoReverse = true,
                        RepeatBehavior = RepeatBehavior.Forever,
                        EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
                    };
                    ((ScaleTransform)dot.RenderTransform).BeginAnimation(ScaleTransform.ScaleXProperty, pulse);
                    ((ScaleTransform)dot.RenderTransform).BeginAnimation(ScaleTransform.ScaleYProperty, pulse);
                }

                dotsPanel.Children.Add(dot);
            }
            txtLiveScan.Children.Add(dotsPanel);

            // Counter + estimate
            string timeHint = progress.EstimatedSeconds > 3 ? $" • ~{progress.EstimatedSeconds}s" : "";
            txtLiveScan.Children.Add(new TextBlock
            {
                Text = $"{progress.PhaseIndex} of {progress.Total} checks{timeHint}",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 11,
                Foreground = (SolidColorBrush)Application.Current.FindResource("TextMuted"),
                HorizontalAlignment = HorizontalAlignment.Center,
                Opacity = 0.7
            });

            txtLiveScan.Visibility = Visibility.Visible;
            metricsPanel.Visibility = Visibility.Collapsed;
        }

        private void ShowRichMetrics()
        {
            txtLiveScan.Visibility = Visibility.Collapsed;
            metricsPanel.Visibility = Visibility.Visible;
        }

        private void SetButtonStates(bool scan, bool export, bool cancel, bool optimize = false)
        {
            btnScan.IsEnabled = scan;
            btnExport.Visibility = export ? Visibility.Visible : Visibility.Collapsed;
            btnCancel.IsEnabled = cancel;
            btnOptimize.Visibility = optimize ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ShowProgress(bool visible)
        {
            progressPanel.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;

            if (visible)
            {
                healthDashboard.Visibility = Visibility.Collapsed;
            }
        }

        private void SetProgressValue(int percent)
        {
            _lastProgressPercent = percent;
            double maxWidth = progressFill.Parent is FrameworkElement parent ? parent.ActualWidth : 600;
            double targetWidth = Math.Max(0, maxWidth * percent / 100.0);

            var animation = new DoubleAnimation
            {
                To = targetWidth,
                Duration = TimeSpan.FromMilliseconds(300),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };
            progressFill.BeginAnimation(WidthProperty, animation);
        }

        private void ProgressTrack_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (progressPanel.Visibility != Visibility.Visible) return;
            double targetWidth = Math.Max(0, e.NewSize.Width * _lastProgressPercent / 100.0);
            progressFill.BeginAnimation(WidthProperty, null);
            progressFill.Width = targetWidth;
        }

        private void UpdateStatus(string text, string icon, SolidColorBrush brush)
        {
            lblStatus.Text = text;
            lblStatus.Foreground = brush;
            statusIcon.Text = icon;
            statusIcon.Foreground = brush;
        }

        private void HandleError(Exception ex)
        {
            ShowProgress(false);
            Log($"✗ ERROR: {ex.Message}");
            UpdateStatus("Error", IcoError, ThemeManager.BrushDanger);
            SetButtonStates(scan: true, export: lastResult != null, cancel: false, optimize: lastResult != null);
        }

        /// <summary>
        /// Shows a dialog for the operator to add optional notes before PDF export.
        /// Returns the notes string (may be empty), or null if the user cancelled.
        /// </summary>
        private string PromptOperatorNotes()
        {
            var dlg = new Window
            {
                Title = "Operator Notes",
                Width = 480, Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize,
                Background = (SolidColorBrush)FindResource("BgCard")
            };
            var sp = new StackPanel { Margin = new Thickness(20) };
            sp.Children.Add(new TextBlock
            {
                Text = "Add notes for this report (optional):",
                FontFamily = new FontFamily("Segoe UI"), FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = (SolidColorBrush)FindResource("TextDark"),
                Margin = new Thickness(0, 0, 0, 8)
            });
            var textBox = new TextBox
            {
                Height = 140,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                FontFamily = new FontFamily("Segoe UI"), FontSize = 12,
                Padding = new Thickness(8),
                Background = (SolidColorBrush)FindResource("LogBg"),
                Foreground = (SolidColorBrush)FindResource("LogFg"),
                BorderBrush = (SolidColorBrush)FindResource("BorderBrush")
            };
            sp.Children.Add(textBox);
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var btnSkip = new Button
            {
                Content = "Skip", Width = 80, Height = 32, Margin = new Thickness(0, 0, 8, 0),
                Style = (Style)FindResource("ModernButton"),
                Background = (Brush)FindResource("NeutralBtnBg"),
                Foreground = (SolidColorBrush)FindResource("NeutralBtnForeground")
            };
            btnSkip.Click += (_, _) => { dlg.Tag = ""; dlg.DialogResult = true; };
            var btnOk = new Button
            {
                Content = "Add Notes", Width = 100, Height = 32,
                Style = (Style)FindResource("BluePrimaryButton")
            };
            btnOk.Click += (_, _) => { dlg.Tag = textBox.Text ?? ""; dlg.DialogResult = true; };
            var btnCancel = new Button
            {
                Content = "Cancel", Width = 80, Height = 32, Margin = new Thickness(8, 0, 0, 0),
                Style = (Style)FindResource("ModernButton"),
                Background = (Brush)FindResource("NeutralBtnBg"),
                Foreground = (SolidColorBrush)FindResource("NeutralBtnForeground")
            };
            btnCancel.Click += (_, _) => { dlg.DialogResult = false; };
            btnPanel.Children.Add(btnSkip);
            btnPanel.Children.Add(btnOk);
            btnPanel.Children.Add(btnCancel);
            sp.Children.Add(btnPanel);
            dlg.Content = sp;

            return dlg.ShowDialog() == true ? (string)dlg.Tag : null;
        }

        private string GeneratePDFReport(DiagnosticResult result)
        {
            string computerName = Environment.MachineName;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string filename = $"EDR_Diagnostic_{computerName}_{timestamp}.pdf";

            string fullPath = null;
            Dispatcher.Invoke(() =>
            {
                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    FileName = filename,
                    DefaultExt = ".pdf",
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };
                fullPath = dlg.ShowDialog() == true ? dlg.FileName : null;
            });

            if (fullPath == null)
                throw new OperationCanceledException("Export cancelled by user.");

            new PDFReportGenerator().Generate(fullPath, result);
            return fullPath;
        }

        private void Log(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => Log(message));
                return;
            }

            string timestamp = $"[{DateTime.Now:h:mm:ss tt}] ";

            // Use dictionary-backed brushes so log entries update on theme switch
            var tsRun = new System.Windows.Documents.Run(timestamp);
            tsRun.SetResourceReference(System.Windows.Documents.TextElement.ForegroundProperty, "LogTimestamp");

            // Message with semantic coloring — all via resource references
            string brushKey;
            if (message.Contains("✓") || message.Contains("complete", StringComparison.OrdinalIgnoreCase))
                brushKey = "SuccessFg";
            else if (message.Contains("✗") || message.Contains("ERROR", StringComparison.OrdinalIgnoreCase)
                     || message.Contains("failed", StringComparison.OrdinalIgnoreCase))
                brushKey = "DangerFg";
            else if (message.Contains("⚠") || message.Contains("Warning", StringComparison.OrdinalIgnoreCase)
                     || message.Contains("timed out", StringComparison.OrdinalIgnoreCase))
                brushKey = "WarnFg";
            else if (message.StartsWith("==="))
                brushKey = "BluePrimary";
            else
                brushKey = "LogFg";

            var msgRun = new System.Windows.Documents.Run(message);
            msgRun.SetResourceReference(System.Windows.Documents.TextElement.ForegroundProperty, brushKey);

            _logParagraph.Inlines.Add(tsRun);
            _logParagraph.Inlines.Add(msgRun);
            _logParagraph.Inlines.Add(new System.Windows.Documents.LineBreak());

            txtLog.ScrollToEnd();
        }

        private void CopyLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var range = new System.Windows.Documents.TextRange(
                    txtLog.Document.ContentStart, txtLog.Document.ContentEnd);
                if (!string.IsNullOrWhiteSpace(range.Text))
                    Clipboard.SetText(range.Text);
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        private void ClearLog_Click(object sender, RoutedEventArgs e)
        {
            _logParagraph.Inlines.Clear();
            Log("Log cleared.");
        }

        private string _statusLabel;
        private string _currentScanPhase;

        private void StartElapsedTimer(string label = "Scanning")
        {
            _scanStopwatch = System.Diagnostics.Stopwatch.StartNew();
            _currentScanPhase = null;
            _statusLabel = label;
            elapsedPanel.Visibility = Visibility.Visible;
            _elapsedTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _elapsedTimer.Tick += (_, _) =>
            {
                var elapsed = _scanStopwatch.Elapsed;
                lblElapsed.Text = $"{elapsed.Minutes}:{elapsed.Seconds:D2}";
                string phase = _currentScanPhase != null ? $" — {_currentScanPhase}" : "";
                lblStatus.Text = $"{_statusLabel}...{phase}";
            };
            _elapsedTimer.Start();
        }

        private void StopElapsedTimer()
        {
            _elapsedTimer?.Stop();
            _elapsedTimer = null;
            _scanStopwatch?.Stop();
            _currentScanPhase = null;
            elapsedPanel.Visibility = Visibility.Collapsed;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_isScanning)
            {
                var result = MessageBox.Show(
                    "A diagnostic scan is currently running.\n\nAre you sure you want to close?",
                    "Scan In Progress", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result != MessageBoxResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
            }
            base.OnClosing(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            ThemeManager.ThemeChanged -= OnThemeChanged;
            if (scanner != null)
            {
                scanner.OnProgress -= Scanner_OnProgress;
                scanner.OnLog -= Log;
                scanner.CancelScan();
                scanner.Dispose();
            }
            base.OnClosed(e);
        }
    }
}
