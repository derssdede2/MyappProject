using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace DLack
{
    public partial class OptimizationResultsWindow : Window
    {
        private static readonly FontFamily MdlFont = new("Segoe MDL2 Assets");
        private static readonly FontFamily SegoeFont = new("Segoe UI");

        private static readonly CubicEase EaseOut = new() { EasingMode = EasingMode.EaseOut };
        private static readonly QuarticEase EaseOutQuart = new() { EasingMode = EasingMode.EaseOut };
        private static readonly ExponentialEase EaseOutExpo = new() { EasingMode = EasingMode.EaseOut, Exponent = 5 };
        private static readonly ElasticEase EaseOutElastic = new() { EasingMode = EasingMode.EaseOut, Oscillations = 1, Springiness = 5 };
        private static readonly BackEase EaseOutBack = new() { EasingMode = EasingMode.EaseOut, Amplitude = 0.4 };

        private readonly long _targetFreedMB;
        private readonly bool _reviewOnly;
        private DispatcherTimer _autoVerifyTimer;
        private int _autoVerifyCountdown = 5;

        private readonly List<OptimizationAction> _actions;
        private readonly OptimizationSummary _summary;

        public bool ShouldRescan { get; private set; }

        /// <param name="reviewOnly">When true, hides the verify button and skips the auto-rescan countdown (used when re-viewing past results).</param>
        public OptimizationResultsWindow(
            List<OptimizationAction> actions,
            OptimizationSummary summary,
            bool reviewOnly = false)
        {
            InitializeComponent();

            _actions = actions;
            _summary = summary;
            _targetFreedMB = summary.TotalFreedMB;
            _reviewOnly = reviewOnly;

            // Set initial state — everything hidden for orchestrated entrance
            lblActionCount.Opacity = 0;
            lblActionBreakdown.Opacity = 0;
            lblFreedSpace.Opacity = 0;

            BuildHeroBadge(summary);
            BuildSummary(actions, summary);
            BuildDetailCards(actions);

            if (_reviewOnly)
                btnVerify.Visibility = Visibility.Collapsed;

            ThemeManager.ThemeChanged += OnThemeChanged;
            Loaded += async (_, _) => await PlayEntranceAsync();
        }

        // ═══════════════════════════════════════════════════════════════
        //  ORCHESTRATED ENTRANCE (single async sequence, no timer spam)
        // ═══════════════════════════════════════════════════════════════

        private async Task PlayEntranceAsync()
        {
            // Hero badge — elastic pop from 0 → 1 with dramatic timing
            AnimateScale(heroBadgeScale, 0, 1.0, 700, EaseOutElastic);
            await Task.Delay(180);

            // Summary text — staggered fade + slide
            AnimateEntrance(lblActionCount, 450);
            await Task.Delay(150);
            AnimateEntrance(lblActionBreakdown, 400);
            await Task.Delay(150);
            AnimateEntrance(lblFreedSpace, 450);

            // Count-up the freed space number
            if (_targetFreedMB > 0)
                _ = AnimateCountUpAsync(300);

            await Task.Delay(120);

            // Detail cards — cascading waterfall with wider spacing
            for (int i = 0; i < resultPanel.Children.Count; i++)
            {
                if (resultPanel.Children[i] is FrameworkElement card)
                    AnimateCardEntrance(card);
                await Task.Delay(80);
            }

            // Start auto-verify countdown after animations complete (skip in review mode)
            if (!_reviewOnly)
                StartAutoVerifyCountdown();
        }

        // ═══════════════════════════════════════════════════════════════
        //  HERO BADGE
        // ═══════════════════════════════════════════════════════════════

        private void BuildHeroBadge(OptimizationSummary summary)
        {
            bool allGood = summary.FailureCount == 0;
            Color color = allGood ? ThemeManager.GoodFg : ThemeManager.WarnFg;
            string icon = allGood ? "\uE73E" : "\uE7BA"; // checkmark vs warning

            heroRing.Stroke = Frozen(color);
            heroFill.Fill = Frozen(color);
            heroIcon.Text = icon;
            heroIcon.Foreground = Frozen(color);
        }

        // ═══════════════════════════════════════════════════════════════
        //  ANIMATION PRIMITIVES
        // ═══════════════════════════════════════════════════════════════

        private static void AnimateScale(ScaleTransform target, double from, double to, int durationMs, IEasingFunction ease)
        {
            var anim = new DoubleAnimation(from, to, TimeSpan.FromMilliseconds(durationMs)) { EasingFunction = ease };
            target.BeginAnimation(ScaleTransform.ScaleXProperty, anim);
            target.BeginAnimation(ScaleTransform.ScaleYProperty, anim);
        }

        private static void AnimateEntrance(FrameworkElement el, int durationMs)
        {
            el.Opacity = 0;
            el.RenderTransform = new TranslateTransform(0, 18);

            var fade = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(durationMs)) { EasingFunction = EaseOut };
            el.BeginAnimation(UIElement.OpacityProperty, fade);

            var slide = new DoubleAnimation(18, 0, TimeSpan.FromMilliseconds(durationMs + 100)) { EasingFunction = EaseOutQuart };
            ((TranslateTransform)el.RenderTransform).BeginAnimation(TranslateTransform.YProperty, slide);
        }

        private static void AnimateCardEntrance(FrameworkElement card)
        {
            card.Opacity = 0;
            card.RenderTransform = new TranslateTransform(0, 24);

            var fade = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(400)) { EasingFunction = EaseOut };
            card.BeginAnimation(UIElement.OpacityProperty, fade);

            var slide = new DoubleAnimation(24, 0, TimeSpan.FromMilliseconds(500)) { EasingFunction = EaseOutExpo };
            ((TranslateTransform)card.RenderTransform).BeginAnimation(TranslateTransform.YProperty, slide);
        }

        private async Task AnimateCountUpAsync(int delayMs)
        {
            await Task.Delay(delayMs);

            lblFreedSpace.Text = "0 MB";
            const int totalFrames = 55;

            for (int frame = 1; frame <= totalFrames; frame++)
            {
                double t = (double)frame / totalFrames;
                double eased = 1 - Math.Pow(1 - t, 4);
                lblFreedSpace.Text = DiagnosticFormatters.FormatMB((long)(_targetFreedMB * eased));
                await Task.Delay(16);
            }
            lblFreedSpace.Text = DiagnosticFormatters.FormatMB(_targetFreedMB);
        }

        // ═══════════════════════════════════════════════════════════════
        //  SUMMARY BANNER
        // ═══════════════════════════════════════════════════════════════

        private void BuildSummary(List<OptimizationAction> actions, OptimizationSummary summary)
        {
            int succeeded = 0, partial = 0, noChange = 0, failed = 0, skipped = 0;
            foreach (var a in actions)
            {
                switch (a.Status)
                {
                    case ActionStatus.Success: succeeded++; break;
                    case ActionStatus.PartialSuccess: partial++; break;
                    case ActionStatus.NoChange: noChange++; break;
                    case ActionStatus.Failed: failed++; break;
                    case ActionStatus.Skipped: skipped++; break;
                }
            }

            lblActionCount.Text = $"{summary.ActionsRun} action{(summary.ActionsRun == 1 ? "" : "s")} completed";

            var parts = new List<string>();
            if (succeeded > 0) parts.Add($"{succeeded} succeeded");
            if (partial > 0) parts.Add($"{partial} partial");
            if (noChange > 0) parts.Add($"{noChange} no change");
            if (failed > 0) parts.Add($"{failed} failed");
            if (skipped > 0) parts.Add($"{skipped} skipped");
            lblActionBreakdown.Text = string.Join("  ·  ", parts);

            lblFreedSpace.Text = summary.TotalFreedMB > 0
                ? DiagnosticFormatters.FormatMB(summary.TotalFreedMB)
                : "—";
        }

        // ═══════════════════════════════════════════════════════════════
        //  DETAIL CARDS
        // ═══════════════════════════════════════════════════════════════

        private void BuildDetailCards(List<OptimizationAction> actions)
        {
            var ordered = actions
                .Where(a => a.Status != ActionStatus.Pending)
                .OrderBy(a => a.Status switch
                {
                    ActionStatus.Failed => 0,
                    ActionStatus.PartialSuccess => 1,
                    ActionStatus.NoChange => 2,
                    ActionStatus.Success => 3,
                    _ => 4
                });

            foreach (var action in ordered)
                resultPanel.Children.Add(BuildResultCard(action));
        }

        private UIElement BuildResultCard(OptimizationAction action)
        {
            (string glyph, Color fg) = action.Status switch
            {
                ActionStatus.Success => ("\uE73E", ThemeManager.GoodFg),
                ActionStatus.PartialSuccess => ("\uE7BA", ThemeManager.WarnFg),
                ActionStatus.NoChange => ("\uE738", ThemeManager.TextMuted),
                ActionStatus.Failed => ("\uE711", ThemeManager.CritFg),
                ActionStatus.Skipped => ("\uE72B", ThemeManager.TextMuted),
                _ => ("", ThemeManager.TextMuted)
            };

            Color borderColor = action.Status switch
            {
                ActionStatus.Success or ActionStatus.PartialSuccess => ThemeManager.GoodFg,
                ActionStatus.Failed => ThemeManager.CritFg,
                _ => ThemeManager.Border
            };

            var wrapper = new Border
            {
                Margin = new Thickness(0, 2, 0, 2),
                Background = Brushes.Transparent,
                RenderTransform = new TranslateTransform(0, 0)
            };

            var card = new Border
            {
                CornerRadius = new CornerRadius(10),
                Background = Frozen(ThemeManager.SurfaceAlt),
                BorderBrush = Frozen(borderColor),
                BorderThickness = new Thickness(1),
                ClipToBounds = true,
                Effect = new DropShadowEffect
                {
                    BlurRadius = 4,
                    ShadowDepth = 1,
                    Opacity = 0.06,
                    Color = Colors.Black,
                    Direction = 270
                }
            };

            // Hover — lift + deeper shadow
            card.MouseEnter += (_, _) =>
            {
                Animate(wrapper.RenderTransform, TranslateTransform.YProperty, -3, 200);
                var s = (DropShadowEffect)card.Effect;
                Animate(s, DropShadowEffect.BlurRadiusProperty, 16, 200);
                Animate(s, DropShadowEffect.ShadowDepthProperty, 5, 200);
                Animate(s, DropShadowEffect.OpacityProperty, 0.15, 200);
            };
            card.MouseLeave += (_, _) =>
            {
                Animate(wrapper.RenderTransform, TranslateTransform.YProperty, 0, 300);
                var s = (DropShadowEffect)card.Effect;
                Animate(s, DropShadowEffect.BlurRadiusProperty, 4, 300);
                Animate(s, DropShadowEffect.ShadowDepthProperty, 1, 300);
                Animate(s, DropShadowEffect.OpacityProperty, 0.06, 300);
            };

            var outerGrid = new Grid();
            outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(4) });
            outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left accent bar — animates width from 0 → 4 on load
            var accent = new Border
            {
                Background = Frozen(fg),
                CornerRadius = new CornerRadius(10, 0, 0, 10),
                Width = 0
            };
            accent.Loaded += (_, _) =>
            {
                var grow = new DoubleAnimation(0, 4, TimeSpan.FromMilliseconds(400))
                {
                    EasingFunction = EaseOutQuart,
                    BeginTime = TimeSpan.FromMilliseconds(250)
                };
                accent.BeginAnimation(WidthProperty, grow);
            };
            Grid.SetColumn(accent, 0);
            outerGrid.Children.Add(accent);

            // Content area
            var content = new Grid { Margin = new Thickness(14, 10, 14, 10) };
            content.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            content.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            content.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(content, 1);
            outerGrid.Children.Add(content);

            // Status icon — pops in with overshoot
            var icon = new TextBlock
            {
                Text = glyph,
                FontFamily = MdlFont,
                FontSize = 16,
                Foreground = Frozen(fg),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(2, 0, 12, 0),
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new ScaleTransform(0, 0)
            };
            icon.Loaded += (_, _) =>
            {
                var pop = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(500))
                {
                    EasingFunction = EaseOutBack,
                    BeginTime = TimeSpan.FromMilliseconds(300)
                };
                ((ScaleTransform)icon.RenderTransform).BeginAnimation(ScaleTransform.ScaleXProperty, pop);
                ((ScaleTransform)icon.RenderTransform).BeginAnimation(ScaleTransform.ScaleYProperty, pop);
            };
            Grid.SetColumn(icon, 0);
            content.Children.Add(icon);

            // Title + result message
            var textStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };

            textStack.Children.Add(new TextBlock
            {
                Text = action.Title,
                FontFamily = SegoeFont,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = Frozen(ThemeManager.TextPrimary),
                TextWrapping = TextWrapping.Wrap
            });

            if (!string.IsNullOrEmpty(action.ResultMessage))
            {
                var resultBorder = new Border
                {
                    Background = Frozen(Color.FromArgb(18, fg.R, fg.G, fg.B)),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 4, 8, 4),
                    Margin = new Thickness(0, 4, 0, 0)
                };
                resultBorder.Child = new TextBlock
                {
                    Text = action.ResultMessage,
                    FontFamily = SegoeFont,
                    FontSize = 11.5,
                    Foreground = Frozen(fg),
                    TextWrapping = TextWrapping.Wrap
                };
                textStack.Children.Add(resultBorder);
            }

            Grid.SetColumn(textStack, 1);
            content.Children.Add(textStack);

            // Freed space (right side)
            if (action.ActualFreedMB > 0)
            {
                var freed = new TextBlock
                {
                    Text = DiagnosticFormatters.FormatMB(action.ActualFreedMB),
                    FontFamily = SegoeFont,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Frozen(ThemeManager.AccentGreen),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(10, 0, 0, 0),
                    MinWidth = 60,
                    TextAlignment = TextAlignment.Right
                };
                Grid.SetColumn(freed, 2);
                content.Children.Add(freed);
            }

            card.Child = outerGrid;
            wrapper.Child = card;
            return wrapper;
        }

        // ═══════════════════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════════════════

        private static void Animate(Animatable target, DependencyProperty prop, double to, int durationMs)
        {
            var anim = new DoubleAnimation(to, TimeSpan.FromMilliseconds(durationMs)) { EasingFunction = EaseOut };
            target.BeginAnimation(prop, anim);
        }

        private static SolidColorBrush Frozen(Color c)
        {
            var b = new SolidColorBrush(c);
            b.Freeze();
            return b;
        }

        // ═══════════════════════════════════════════════════════════════
        //  AUTO-VERIFY COUNTDOWN
        // ═══════════════════════════════════════════════════════════════

        private void StartAutoVerifyCountdown()
        {
            _autoVerifyCountdown = 5;
            UpdateVerifyText();

            _autoVerifyTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _autoVerifyTimer.Tick += (_, _) =>
            {
                _autoVerifyCountdown--;
                if (_autoVerifyCountdown > 0)
                {
                    UpdateVerifyText();
                }
                else
                {
                    _autoVerifyTimer.Stop();
                    ShouldRescan = true;
                    Close();
                }
            };
            _autoVerifyTimer.Start();
        }

        private void UpdateVerifyText()
        {
            if (btnVerify.Content is StackPanel sp &&
                sp.Children.OfType<TextBlock>().LastOrDefault() is { } tb)
            {
                tb.Text = $"Verifying in {_autoVerifyCountdown}s...";
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  EVENT HANDLERS
        // ═══════════════════════════════════════════════════════════════

        private void CleanupWindowState()
        {
            _autoVerifyTimer?.Stop();
            _autoVerifyTimer = null;
            ThemeManager.ThemeChanged -= OnThemeChanged;
        }

        private void BtnVerify_Click(object sender, RoutedEventArgs e)
        {
            ShouldRescan = true;
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                Close();
                e.Handled = true;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            CleanupWindowState();
            base.OnClosed(e);
        }

        private void OnThemeChanged()
        {
            resultPanel.Children.Clear();
            BuildHeroBadge(_summary);
            BuildSummary(_actions, _summary);
            BuildDetailCards(_actions);
        }
    }
}
