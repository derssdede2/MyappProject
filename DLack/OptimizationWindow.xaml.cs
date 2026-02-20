using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace DLack
{
    public partial class OptimizationWindow : Window
    {
        private readonly List<OptimizationAction> _actions;
        private readonly Optimizer _optimizer;
        private readonly Action<string> _logCallback;
        private readonly DiagnosticResult _diagnosticResult;
        private CancellationTokenSource _cts;
        private bool _isRunning;
        private bool _didRun;
        private DispatcherTimer _countdownTimer;
        private int _countdownRemaining;
        private string _currentActionBase;
        private bool _rebuildPending;

        private static readonly FontFamily MdlFont = new("Segoe MDL2 Assets");
        private static readonly FontFamily SegoeFont = new("Segoe UI");

        /// <summary>Creates a frozen (immutable, thread-safe) SolidColorBrush.</summary>
        private static SolidColorBrush FrozenBrush(Color c)
        {
            var b = new SolidColorBrush(c);
            b.Freeze();
            return b;
        }

        public bool ShouldRescan { get; private set; }

        /// <summary>
        /// True after optimization has completed at least once in this window.
        /// </summary>
        public bool DidRun => _didRun;

        public OptimizationWindow(
            List<OptimizationAction> actions,
            Optimizer optimizer,
            Action<string> logCallback,
            DiagnosticResult diagnosticResult)
        {
            InitializeComponent();
            _actions = actions;
            _optimizer = optimizer;
            _logCallback = logCallback;
            _diagnosticResult = diagnosticResult;

            ThemeManager.ThemeChanged += OnThemeChanged;
            _optimizer.OnActionStatusChanged += OnActionStatusChanged;
            BuildActionCards();
            UpdateSummary();
        }

        private void OnThemeChanged()
        {
            BuildActionCards();
        }

        private void OnActionStatusChanged(OptimizationAction action)
        {
            // Coalesce rapid-fire status changes into a single UI rebuild per frame.
            // During optimization, OnActionStatusChanged fires 2× per action (Running → result),
            // and each BuildActionCards call destroys/recreates ALL WPF elements.
            if (_rebuildPending) return;
            _rebuildPending = true;

            Dispatcher.BeginInvoke(() =>
            {
                _rebuildPending = false;
                BuildActionCards();
                UpdateSummary();
            }, DispatcherPriority.Background);
        }

        // ═══════════════════════════════════════════════════════════════
        //  BUILD ACTION CARDS
        // ═══════════════════════════════════════════════════════════════

        private void BuildActionCards()
        {
            actionPanel.Children.Clear();

            var autoFixes = _actions.Where(a => a.Type == ActionType.AutoFix).ToList();
            var shortcuts = _actions.Where(a => a.Type == ActionType.Shortcut).ToList();
            var manualOnly = _actions.Where(a => a.Type == ActionType.ManualOnly).ToList();

            // ── AUTO FIXES SECTION ──
            if (autoFixes.Count > 0)
            {
                AddSectionHeader("\uE713", "AUTOMATIC FIXES", ThemeManager.AccentBlue);
                var groups = autoFixes
                    .GroupBy(a => a.Category)
                    .OrderByDescending(g => g.Max(a => a.Risk))
                    .ThenBy(g => g.Key);

                foreach (var group in groups)
                {
                    AddCategoryHeader(group.Key);
                    foreach (var action in group)
                        actionPanel.Children.Add(BuildActionCard(action));
                }
            }

            // ── MANUAL ACTIONS SECTION ──
            if (shortcuts.Count + manualOnly.Count > 0)
            {
                AddSectionHeader("\uE8A7", "MANUAL ACTIONS", ThemeManager.TextMuted);

                var allManual = shortcuts.Concat(manualOnly).ToList();
                var groups = allManual
                    .GroupBy(a => a.Category)
                    .OrderBy(g => g.Key);

                foreach (var group in groups)
                {
                    AddCategoryHeader(group.Key);
                    foreach (var action in group)
                        actionPanel.Children.Add(BuildActionCard(action));
                }
            }
        }

        private void AddSectionHeader(string icon, string text, Color color)
        {
            var border = new Border
            {
                Background = FrozenBrush(Color.FromArgb(15, color.R, color.G, color.B)),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 6, 10, 6),
                Margin = new Thickness(0, 10, 0, 4)
            };
            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            sp.Children.Add(new TextBlock
            {
                Text = icon,
                FontFamily = MdlFont,
                FontSize = 12,
                Foreground = FrozenBrush(color),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            });
            sp.Children.Add(new TextBlock
            {
                Text = text,
                FontFamily = SegoeFont,
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = FrozenBrush(color)
            });
            border.Child = sp;
            actionPanel.Children.Add(border);
        }

        private void AddCategoryHeader(string category)
        {
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(4, 8, 0, 4) };
            headerPanel.Children.Add(new TextBlock
            {
                Text = GetCategoryIcon(category),
                FontFamily = MdlFont,
                FontSize = 11,
                Foreground = FrozenBrush(ThemeManager.TextMuted),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            });
            headerPanel.Children.Add(new TextBlock
            {
                Text = category.ToUpperInvariant(),
                FontFamily = SegoeFont,
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = FrozenBrush(ThemeManager.TextMuted)
            });
            actionPanel.Children.Add(headerPanel);
        }

        private UIElement BuildActionCard(OptimizationAction action)
        {
            Color riskColor = action.Risk switch
            {
                ActionRisk.RequiresReboot => ThemeManager.CritFg,
                ActionRisk.Moderate => ThemeManager.WarnFg,
                _ => ThemeManager.GoodFg
            };
            Color riskBg = action.Risk switch
            {
                ActionRisk.RequiresReboot => ThemeManager.CritBg,
                ActionRisk.Moderate => ThemeManager.WarnBg,
                _ => ThemeManager.GoodBg
            };
            string riskLabel = action.Risk switch
            {
                ActionRisk.RequiresReboot => "REBOOT",
                ActionRisk.Moderate => "MODERATE",
                _ => "SAFE"
            };

            Color statusColor = action.Status switch
            {
                ActionStatus.Success => ThemeManager.GoodFg,
                ActionStatus.PartialSuccess => ThemeManager.WarnFg,
                ActionStatus.NoChange => ThemeManager.TextMuted,
                ActionStatus.Failed => ThemeManager.CritFg,
                ActionStatus.Running => ThemeManager.AccentBlue,
                ActionStatus.Skipped => ThemeManager.TextMuted,
                _ => Colors.Transparent
            };

            // Accent bar: type-based before run, status-based after run
            Color accentColor = action.Status switch
            {
                ActionStatus.Success => ThemeManager.GoodFg,
                ActionStatus.PartialSuccess => ThemeManager.WarnFg,
                ActionStatus.NoChange => ThemeManager.TextMuted,
                ActionStatus.Failed => ThemeManager.CritFg,
                _ => action.Type switch
                {
                    ActionType.AutoFix => ThemeManager.GoodFg,
                    ActionType.Shortcut => ThemeManager.WarnFg,
                    _ => ThemeManager.TextMuted
                }
            };

            var card = new Border
            {
                CornerRadius = new CornerRadius(10),
                Background = FrozenBrush(ThemeManager.SurfaceAlt),
                BorderBrush = FrozenBrush(
                    action.Status is ActionStatus.Success or ActionStatus.PartialSuccess ? ThemeManager.GoodFg :
                    action.Status == ActionStatus.Failed ? ThemeManager.CritFg :
                    action.Status == ActionStatus.NoChange ? ThemeManager.TextSubtle :
                    ThemeManager.Border),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(0),
                Margin = new Thickness(0, 2, 0, 2),
                ClipToBounds = true,
                Opacity = _isRunning ? 1 : 0,
                RenderTransform = _isRunning ? Transform.Identity : new TranslateTransform(0, 14)
            };

            // Fade + slide entrance — skip during optimization to avoid flicker on rapid rebuilds
            if (!_isRunning)
            {
                card.Loaded += (_, _) =>
                {
                    var fadeIn = new DoubleAnimation
                    {
                        From = 0, To = 1,
                        Duration = TimeSpan.FromMilliseconds(400),
                        EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
                    };
                    card.BeginAnimation(OpacityProperty, fadeIn);

                    var slideUp = new DoubleAnimation
                    {
                        From = 14, To = 0,
                        Duration = TimeSpan.FromMilliseconds(450),
                        EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut }
                    };
                    ((TranslateTransform)card.RenderTransform).BeginAnimation(TranslateTransform.YProperty, slideUp);
                };
            }

            var outerGrid = new Grid();
            outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(4) });
            outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left accent bar
            var accentBar = new Border
            {
                Background = FrozenBrush(accentColor),
                CornerRadius = new CornerRadius(10, 0, 0, 10)
            };
            Grid.SetColumn(accentBar, 0);
            outerGrid.Children.Add(accentBar);

            var grid = new Grid { Margin = new Thickness(14, 10, 14, 10) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(grid, 1);
            outerGrid.Children.Add(grid);

            // Control: checkbox for AutoFix, "Open" button for Shortcuts, info icon for ManualOnly
            // After completion, show a status icon instead of the checkbox
            bool isCompleted = action.Status is ActionStatus.Success or ActionStatus.PartialSuccess
                or ActionStatus.NoChange or ActionStatus.Failed or ActionStatus.Skipped;

            if (isCompleted)
            {
                string glyph = action.Status switch
                {
                    ActionStatus.Success => "\uE73E",        // checkmark
                    ActionStatus.PartialSuccess => "\uE7BA",  // warning
                    ActionStatus.NoChange => "\uE738",        // dash/minus
                    ActionStatus.Failed => "\uE711",          // X
                    ActionStatus.Skipped => "\uE72B",         // forward
                    _ => ""
                };
                var statusIcon = new TextBlock
                {
                    Text = glyph,
                    FontFamily = MdlFont,
                    FontSize = 16,
                    Foreground = FrozenBrush(statusColor),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, 0, 12, 0)
                };
                Grid.SetColumn(statusIcon, 0);
                grid.Children.Add(statusIcon);
            }
            else if (action.Type == ActionType.AutoFix)
            {
                var cb = new CheckBox
                {
                    IsChecked = action.IsSelected,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 10, 0),
                    IsEnabled = !_isRunning && action.Status == ActionStatus.Pending
                };
                cb.Checked += (s, e) => { action.IsSelected = true; UpdateSummary(); };
                cb.Unchecked += (s, e) => { action.IsSelected = false; UpdateSummary(); };
                Grid.SetColumn(cb, 0);
                grid.Children.Add(cb);
            }
            else if (action.Type == ActionType.Shortcut && action.Status == ActionStatus.Pending)
            {
                var openBtn = new Button
                {
                    Content = "Open",
                    FontFamily = SegoeFont,
                    FontSize = 10,
                    FontWeight = FontWeights.SemiBold,
                    Padding = new Thickness(8, 2, 8, 2),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 10, 0),
                    Background = FrozenBrush(ThemeManager.SurfaceAlt),
                    Foreground = FrozenBrush(ThemeManager.AccentBlue),
                    BorderBrush = FrozenBrush(ThemeManager.AccentBlue),
                    BorderThickness = new Thickness(1),
                    IsEnabled = !_isRunning
                };
                openBtn.Click += (s, e) =>
                {
                    action.IsSelected = true;
                    _ = RunSingleShortcutAsync(action);
                };
                Grid.SetColumn(openBtn, 0);
                grid.Children.Add(openBtn);
            }
            else
            {
                var icon = new TextBlock
                {
                    Text = action.Status == ActionStatus.Pending ? "\uE946" : "", // info icon for pending
                    FontFamily = MdlFont,
                    FontSize = 14,
                    Foreground = FrozenBrush(ThemeManager.TextMuted),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, 0, 10, 0)
                };
                Grid.SetColumn(icon, 0);
                grid.Children.Add(icon);
            }

            // Content
            var content = new StackPanel { VerticalAlignment = VerticalAlignment.Center };

            // Title — always clean, no inline result
            var titleBlock = new TextBlock
            {
                Text = action.Title,
                FontFamily = SegoeFont,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = FrozenBrush(ThemeManager.TextPrimary),
                TextWrapping = TextWrapping.Wrap
            };

            // Running state: add an animated ellipsis
            if (action.Status == ActionStatus.Running)
            {
                var titleRow = new StackPanel { Orientation = Orientation.Horizontal };
                titleRow.Children.Add(titleBlock);
                titleRow.Children.Add(new TextBlock
                {
                    Text = " — Running...",
                    FontFamily = SegoeFont,
                    FontSize = 12,
                    Foreground = FrozenBrush(ThemeManager.AccentBlue),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(4, 0, 0, 0)
                });
                content.Children.Add(titleRow);
            }
            else
            {
                content.Children.Add(titleBlock);
            }

            // Result message — own line with wrapping (only shown after completion)
            if (isCompleted && !string.IsNullOrEmpty(action.ResultMessage))
            {
                var resultBorder = new Border
                {
                    Background = FrozenBrush(Color.FromArgb(18, statusColor.R, statusColor.G, statusColor.B)),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 4, 8, 4),
                    Margin = new Thickness(0, 4, 0, 0)
                };
                resultBorder.Child = new TextBlock
                {
                    Text = action.ResultMessage,
                    FontFamily = SegoeFont,
                    FontSize = 11.5,
                    Foreground = FrozenBrush(statusColor),
                    TextWrapping = TextWrapping.Wrap
                };
                content.Children.Add(resultBorder);
            }

            // Verification badge — shown after verification pass
            if (action.Verification is not VerificationStatus.None)
            {
                (string vGlyph, Color vColor, string vLabel) = action.Verification switch
                {
                    VerificationStatus.Verifying => ("\uE895", ThemeManager.AccentBlue, "Verifying..."),
                    VerificationStatus.Verified => ("\uE73E", ThemeManager.GoodFg, "Verified"),
                    VerificationStatus.PartiallyVerified => ("\uE7BA", ThemeManager.WarnFg, "Partially Verified"),
                    VerificationStatus.NotVerified => ("\uE711", ThemeManager.CritFg, "Not Verified"),
                    VerificationStatus.RequiresReboot => ("\uE777", ThemeManager.TextMuted, "Requires Reboot"),
                    _ => ("", ThemeManager.TextMuted, "")
                };

                var verifyRow = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 4, 0, 0)
                };
                verifyRow.Children.Add(new TextBlock
                {
                    Text = vGlyph,
                    FontFamily = MdlFont,
                    FontSize = 11,
                    Foreground = FrozenBrush(vColor),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 6, 0)
                });
                verifyRow.Children.Add(new TextBlock
                {
                    Text = vLabel,
                    FontFamily = SegoeFont,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = FrozenBrush(vColor),
                    VerticalAlignment = VerticalAlignment.Center
                });
                if (!string.IsNullOrEmpty(action.VerificationMessage)
                    && action.Verification is not VerificationStatus.Verifying)
                {
                    verifyRow.Children.Add(new TextBlock
                    {
                        Text = $" — {action.VerificationMessage}",
                        FontFamily = SegoeFont,
                        FontSize = 10.5,
                        Foreground = FrozenBrush(ThemeManager.TextMuted),
                        VerticalAlignment = VerticalAlignment.Center,
                        TextTrimming = TextTrimming.CharacterEllipsis
                    });
                }
                content.Children.Add(verifyRow);
            }

            // Description — shown when pending or running (hide after run to reduce noise)
            if (!isCompleted)
            {
                content.Children.Add(new TextBlock
                {
                    Text = action.Description,
                    FontFamily = SegoeFont,
                    FontSize = 11,
                    Foreground = FrozenBrush(ThemeManager.TextMuted),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 2, 0, 0)
                });
            }
            Grid.SetColumn(content, 1);
            grid.Children.Add(content);

            // Risk badge
            var badge = new Border
            {
                CornerRadius = new CornerRadius(4),
                Background = FrozenBrush(riskBg),
                Padding = new Thickness(8, 2, 8, 2),
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            badge.Child = new TextBlock
            {
                Text = riskLabel,
                FontFamily = SegoeFont,
                FontSize = 9,
                FontWeight = FontWeights.Bold,
                Foreground = FrozenBrush(riskColor)
            };
            Grid.SetColumn(badge, 2);
            grid.Children.Add(badge);

            // Estimated space
            if (action.EstimatedFreeMB > 0)
            {
                var est = new TextBlock
                {
                    Text = $"~{FormatMB(action.EstimatedFreeMB)}",
                    FontFamily = SegoeFont,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = FrozenBrush(ThemeManager.AccentGreen),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(10, 0, 0, 0),
                    MinWidth = 60,
                    TextAlignment = TextAlignment.Right
                };
                Grid.SetColumn(est, 3);
                grid.Children.Add(est);
            }
            else if (action.Status == ActionStatus.Pending && !string.IsNullOrEmpty(action.EstimatedDuration))
            {
                var dur = new TextBlock
                {
                    Text = action.EstimatedDuration,
                    FontFamily = SegoeFont,
                    FontSize = 10,
                    Foreground = FrozenBrush(ThemeManager.TextSubtle),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(10, 0, 0, 0),
                    MinWidth = 60,
                    TextAlignment = TextAlignment.Right
                };
                Grid.SetColumn(dur, 3);
                grid.Children.Add(dur);
            }

            card.Child = outerGrid;
            return card;
        }

        // ═══════════════════════════════════════════════════════════════
        //  SUMMARY
        // ═══════════════════════════════════════════════════════════════

        private void UpdateSummary()
        {
            if (_didRun)
            {
                // After completion: single pass over actions for all counters
                int succeeded = 0, partial = 0, noChange = 0, failed = 0, skipped = 0;
                long freedMB = 0, estimatedMB = 0;
                int vVerified = 0, vPartial = 0, vNotVerified = 0, vReboot = 0;
                foreach (var a in _actions)
                {
                    switch (a.Status)
                    {
                        case ActionStatus.Success: succeeded++; break;
                        case ActionStatus.PartialSuccess: partial++; break;
                        case ActionStatus.NoChange: noChange++; break;
                        case ActionStatus.Failed: failed++; break;
                        case ActionStatus.Skipped: skipped++; break;
                    }
                    switch (a.Verification)
                    {
                        case VerificationStatus.Verified: vVerified++; break;
                        case VerificationStatus.PartiallyVerified: vPartial++; break;
                        case VerificationStatus.NotVerified: vNotVerified++; break;
                        case VerificationStatus.RequiresReboot: vReboot++; break;
                    }
                    freedMB += a.ActualFreedMB;
                    if (a.IsSelected) estimatedMB += a.EstimatedFreeMB;
                }

                // Action count — show verification summary if verification has run
                bool hasVerification = vVerified + vPartial + vNotVerified + vReboot > 0;
                if (hasVerification)
                {
                    var vParts = new List<string>();
                    if (vVerified > 0) vParts.Add($"{vVerified} verified");
                    if (vPartial > 0) vParts.Add($"{vPartial} partial");
                    if (vNotVerified > 0) vParts.Add($"{vNotVerified} not verified");
                    if (vReboot > 0) vParts.Add($"{vReboot} needs reboot");
                    lblActionCount.Text = string.Join("  ·  ", vParts);
                }
                else
                {
                    var parts = new List<string>();
                    if (succeeded > 0) parts.Add($"{succeeded} ✓");
                    if (partial > 0) parts.Add($"{partial} partial");
                    if (noChange > 0) parts.Add($"{noChange} unchanged");
                    if (failed > 0) parts.Add($"{failed} failed");
                    if (skipped > 0) parts.Add($"{skipped} skipped");
                    lblActionCount.Text = parts.Count > 0 ? string.Join("  ·  ", parts) : "Done";
                }

                // Space: show actual vs estimated when they differ significantly
                if (freedMB > 0 && estimatedMB > 0 && freedMB < estimatedMB / 2)
                    lblEstimatedFree.Text = $"{FormatMB(freedMB)} of ~{FormatMB(estimatedMB)}";
                else if (freedMB > 0)
                    lblEstimatedFree.Text = FormatMB(freedMB);
                else
                    lblEstimatedFree.Text = noChange > 0 ? "No space freed" : "Done";

                // Color the freed label based on outcome
                lblEstimatedFree.Foreground = FrozenBrush(
                    (noChange > 0 || partial > 0) && failed == 0 ? ThemeManager.WarnFg :
                    failed > 0 ? ThemeManager.CritFg : ThemeManager.GoodFg);

                btnRunOptimize.Visibility = Visibility.Collapsed;
            }
            else
            {
                // Before run: show estimates (only AutoFix actions)
                var selected = _actions.Where(a => a.IsSelected && a.Type == ActionType.AutoFix).ToList();
                int count = selected.Count;
                long totalMB = selected.Sum(a => a.EstimatedFreeMB);

                lblActionCount.Text = $"{count} action{(count == 1 ? "" : "s")}";
                lblEstimatedFree.Text = $"~{FormatMB(totalMB)}";
                lblEstimatedFree.Foreground = (Brush)FindResource("GreenSuccess");

                btnRunOptimize.IsEnabled = count > 0 && !_isRunning;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  BUTTON HANDLERS
        // ═══════════════════════════════════════════════════════════════

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var a in _actions.Where(a => a.Type == ActionType.AutoFix && a.Status == ActionStatus.Pending))
                a.IsSelected = true;
            BuildActionCards();
            UpdateSummary();
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var a in _actions.Where(a => a.Type == ActionType.AutoFix))
                a.IsSelected = false;
            BuildActionCards();
            UpdateSummary();
        }

        private async void BtnRunOptimize_Click(object sender, RoutedEventArgs e)
        {
            // Confirm Moderate/Reboot actions
            var risky = _actions.Where(a =>
                a.IsSelected && a.IsAutomatable &&
                a.Risk is ActionRisk.Moderate or ActionRisk.RequiresReboot).ToList();

            if (risky.Count > 0)
            {
                string riskList = string.Join("\n", risky.Select(a => $"  • {a.Title} ({a.Risk})"));
                var confirm = MessageBox.Show(
                    $"The following actions have elevated risk:\n\n{riskList}\n\nProceed?",
                    "Confirm Optimization",
                    MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (confirm != MessageBoxResult.Yes) return;
            }

            _isRunning = true;
            _cts = new CancellationTokenSource();
            btnRunOptimize.IsEnabled = false;
            btnSelectAll.IsEnabled = false;
            btnSelectNone.IsEnabled = false;
            optProgressPanel.Visibility = Visibility.Visible;

            // Show "Optimizing your system" overlay
            ShowOverlay("\uE713", "Optimizing your system\u2026",
                "Applying selected fixes — this may take a few minutes");
            var overlayStopwatch = System.Diagnostics.Stopwatch.StartNew();

            _logCallback?.Invoke("\r\n=== STARTING OPTIMIZATION ===");

            var progress = new Progress<OptimizationProgress>(p =>
            {
                // Current is 1-based and means "starting action N", so use (N-1)/Total
                // to avoid filling the bar before the last action finishes.
                int pct = p.Total > 0 ? (int)((float)(p.Current - 1) / p.Total * 100) : 0;
                SetProgressValue(pct);

                _currentActionBase = $"{p.Current}/{p.Total} — {p.CurrentAction}";
                _countdownRemaining = p.EstimatedSeconds;

                // Show initial countdown text
                lblOptProgress.Text = _countdownRemaining > 0
                    ? $"{_currentActionBase} ({FormatCountdown(_countdownRemaining)})"
                    : _currentActionBase;

                // Start/stop the countdown timer based on whether this action is long-running
                if (p.IsLongRunning && _countdownRemaining > 0)
                {
                    if (_countdownTimer == null)
                    {
                        _countdownTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
                            _countdownTimer.Tick += (_, _) =>
                            {
                                _countdownRemaining--;
                                if (_countdownRemaining > 0)
                                {
                                    lblOptProgress.Text = $"{_currentActionBase} ({FormatCountdown(_countdownRemaining)})";
                                }
                                else
                                {
                                    int overtime = -_countdownRemaining;
                                    string overStr = overtime >= 60
                                        ? $"{overtime / 60}m {overtime % 60:D2}s"
                                        : $"{overtime}s";
                                    lblOptProgress.Text = $"{_currentActionBase} (still working... +{overStr})";
                                }
                            };
                    }
                    _countdownTimer.Start();

                    var pulse = new DoubleAnimation(0.45, 1.0, TimeSpan.FromMilliseconds(900))
                    {
                        AutoReverse = true,
                        RepeatBehavior = RepeatBehavior.Forever,
                        EasingFunction = new SineEase()
                    };
                    optProgressFill.BeginAnimation(OpacityProperty, pulse);
                }
                else
                {
                    _countdownTimer?.Stop();
                    optProgressFill.BeginAnimation(OpacityProperty, null);
                    optProgressFill.Opacity = 1.0;
                }
            });

            try
            {
                var summary = await _optimizer.ExecuteAsync(_actions, progress, _cts.Token);

                // Store results on the diagnostic result so the PDF report can include them
                _diagnosticResult.OptimizationActions = _actions;
                _diagnosticResult.OptimizationSummary = summary;

                _didRun = true;

                _logCallback?.Invoke($"✓ Optimization complete — {summary.SuccessCount} succeeded, {summary.FailureCount} failed");
                foreach (var detail in summary.ActionDetails)
                    _logCallback?.Invoke($"  {detail}");

                // Phase 1 minimum: keep "Optimizing…" overlay visible for at least 5s
                const int MinOptimizeMs = 5_000;
                long remaining = MinOptimizeMs - overlayStopwatch.ElapsedMilliseconds;
                if (remaining > 0)
                    await Task.Delay((int)remaining, _cts.Token);

                // ── Phase 2: Verify only the affected items ──
                _logCallback?.Invoke("\r\n=== VERIFYING IMPROVEMENTS ===");

                // Swap overlay to verification phase
                ShowOverlay("\uE73E", "Verifying changes\u2026",
                    "Confirming each fix was applied correctly");
                var verifyStopwatch = System.Diagnostics.Stopwatch.StartNew();

                lblOptProgress.Text = "Verifying improvements...";
                SetProgressValue(0);

                var verifyProgress = new Progress<OptimizationProgress>(vp =>
                {
                    int pct = vp.Total > 0 ? (int)((float)vp.Current / vp.Total * 100) : 0;
                    SetProgressValue(pct);
                    lblOptProgress.Text = vp.CurrentAction;
                });

                await _optimizer.VerifyActionsAsync(_actions, verifyProgress, _cts.Token);

                int verified = _actions.Count(a => a.Verification == VerificationStatus.Verified);
                int partial = _actions.Count(a => a.Verification == VerificationStatus.PartiallyVerified);
                int notVerified = _actions.Count(a => a.Verification == VerificationStatus.NotVerified);
                int reboot = _actions.Count(a => a.Verification == VerificationStatus.RequiresReboot);

                var vParts = new List<string>();
                if (verified > 0) vParts.Add($"{verified} verified");
                if (partial > 0) vParts.Add($"{partial} partial");
                if (notVerified > 0) vParts.Add($"{notVerified} not verified");
                if (reboot > 0) vParts.Add($"{reboot} needs reboot");
                _logCallback?.Invoke($"✓ Verification: {string.Join(", ", vParts)}");

                SetProgressValue(100);

                // Phase 2 minimum: keep "Verifying…" overlay visible for at least 5s
                const int MinVerifyMs = 5_000;
                long verifyRemaining = MinVerifyMs - verifyStopwatch.ElapsedMilliseconds;
                if (verifyRemaining > 0)
                    await Task.Delay((int)verifyRemaining, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                _logCallback?.Invoke("✗ Optimization cancelled.");
            }
            catch (Exception ex)
            {
                _logCallback?.Invoke($"✗ Optimization error: {ex.Message}");
            }
            finally
            {
                _isRunning = false;
                _countdownTimer?.Stop();
                optProgressFill.BeginAnimation(OpacityProperty, null);
                optProgressFill.Opacity = 1.0;
                optProgressPanel.Visibility = Visibility.Collapsed;
                HideOverlay();

                // Always refresh cards and stay open — results and verification
                // are shown inline on the action cards
                BuildActionCards();
                UpdateSummary();
            }
        }

        private async Task RunSingleShortcutAsync(OptimizationAction action)
        {
            try
            {
                action.Status = ActionStatus.Running;
                BuildActionCards();
                await Task.Run(() => _optimizer.DispatchShortcut(action.ActionKey));
                action.Status = ActionStatus.Success;
                action.ResultMessage = "Opened";
                _logCallback?.Invoke($"  ✓ {action.Title} — Opened");
            }
            catch (Exception ex)
            {
                action.Status = ActionStatus.Failed;
                action.ResultMessage = ex.Message;
                _logCallback?.Invoke($"  ✗ {action.Title} — {ex.Message}");
            }
            BuildActionCards();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning)
            {
                var result = MessageBox.Show(
                    "Optimization is in progress. Cancel and close?",
                    "Cancel Optimization", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result != MessageBoxResult.Yes) return;
                _cts?.Cancel();
            }
            Close();
        }

        // ═══════════════════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════════════════

        private static string GetCategoryIcon(string category) => category switch
        {
            "Disk Cleanup" => "\uE74D",
            "Browser" => "\uE774",
            "CPU" => "\uE9D9",
            "Performance" => "\uE9D9",
            "Visual Settings" => "\uE771",
            "Network" => "\uE968",
            "RAM" => "\uE7F4",
            "Startup" => "\uE7E8",
            "Windows Update" => "\uE895",
            "System" => "\uE770",
            "GPU" => "\uE7F4",
            "Office" => "\uE8A5",
            "Security" => "\uE72E",
            "Event Log" => "\uE7BA",
            "Antivirus" => "\uE83D",
            _ => "\uE946"
        };

        private void SetProgressValue(int percent)
        {
            double maxWidth = optProgressFill.Parent is FrameworkElement parent ? parent.ActualWidth : 400;
            double targetWidth = Math.Max(0, maxWidth * percent / 100.0);

            var animation = new DoubleAnimation
            {
                To = targetWidth,
                Duration = TimeSpan.FromMilliseconds(250),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };
            optProgressFill.BeginAnimation(WidthProperty, animation);
        }

        private static string FormatMB(long mb) => DiagnosticFormatters.FormatMB(mb);

        private static string FormatCountdown(int totalSeconds)
        {
            return totalSeconds >= 60
                ? $"~{totalSeconds / 60}m {totalSeconds % 60:D2}s remaining"
                : $"~{totalSeconds}s remaining";
        }

        // ═══════════════════════════════════════════════════════════════
        //  ANIMATED OVERLAY
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Builds the animated spinner overlay with a phase label and description.
        /// Uses the same comet-tail orbit spinner as the scan screen.
        /// </summary>
        private void ShowOverlay(string icon, string phase, string description)
        {
            optOverlayContent.Children.Clear();

            var accent = ThemeManager.AccentBlue;

            // ── Orbit spinner ──
            const double r = 52;
            const double stroke = 3.5;
            const double size = (r + stroke + 2) * 2;
            double cx = size / 2, cy = size / 2;

            var ring = new Grid
            {
                Width = size, Height = size,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 22)
            };

            // Track circle
            ring.Children.Add(new Ellipse
            {
                Width = r * 2, Height = r * 2,
                Stroke = FrozenBrush(Color.FromArgb(28, accent.R, accent.G, accent.B)),
                StrokeThickness = stroke,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            });

            // Comet-tail arcs spinning together
            var spinCanvas = new Canvas
            {
                Width = size, Height = size,
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new RotateTransform(0)
            };

            var head = BuildOrbitArc(cx, cy, r, 0, 90, stroke, accent);
            head.Opacity = 0.85;
            spinCanvas.Children.Add(head);

            var body = BuildOrbitArc(cx, cy, r, 95, 80, stroke, accent);
            body.Opacity = 0.45;
            spinCanvas.Children.Add(body);

            var tail = BuildOrbitArc(cx, cy, r, 180, 70, stroke, accent);
            tail.Opacity = 0.20;
            spinCanvas.Children.Add(tail);

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

            // Center icon inside spinner
            var centerIcon = new TextBlock
            {
                Text = icon,
                FontFamily = MdlFont,
                FontSize = 36,
                Foreground = FrozenBrush(ThemeManager.TextPrimary),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Opacity = 0.55,
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new ScaleTransform(1, 1)
            };

            // Gentle breathing pulse on the center icon
            var breathe = new DoubleAnimation
            {
                From = 0.9, To = 1.1,
                Duration = TimeSpan.FromMilliseconds(1600),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
            };
            ((ScaleTransform)centerIcon.RenderTransform).BeginAnimation(ScaleTransform.ScaleXProperty, breathe);
            ((ScaleTransform)centerIcon.RenderTransform).BeginAnimation(ScaleTransform.ScaleYProperty, breathe);
            ring.Children.Add(centerIcon);

            optOverlayContent.Children.Add(ring);

            // Phase title
            optOverlayContent.Children.Add(new TextBlock
            {
                Text = phase,
                FontFamily = SegoeFont,
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = FrozenBrush(ThemeManager.TextPrimary),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 4)
            });

            // Description
            optOverlayContent.Children.Add(new TextBlock
            {
                Text = description,
                FontFamily = SegoeFont,
                FontSize = 12,
                Foreground = FrozenBrush(ThemeManager.TextMuted),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 0)
            });

            optOverlay.Opacity = 0;
            optOverlay.Visibility = Visibility.Visible;
            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(450))
            {
                EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut }
            };
            optOverlay.BeginAnimation(OpacityProperty, fadeIn);
        }

        private void HideOverlay()
        {
            var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(350))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
            };
            fadeOut.Completed += (_, _) =>
            {
                optOverlay.Visibility = Visibility.Collapsed;
            };
            optOverlay.BeginAnimation(OpacityProperty, fadeOut);
        }

        private static Path BuildOrbitArc(
            double cx, double cy, double r,
            double startDeg, double sweepDeg, double stroke, Color color)
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
                Stroke = FrozenBrush(color),
                StrokeThickness = stroke,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round
            };
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_isRunning)
            {
                _cts?.Cancel();
            }
            _cts?.Dispose();
            _optimizer.OnActionStatusChanged -= OnActionStatusChanged;
            ThemeManager.ThemeChanged -= OnThemeChanged;
            base.OnClosing(e);
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape && !_isRunning)
            {
                Close();
                e.Handled = true;
            }
        }
    }
}
