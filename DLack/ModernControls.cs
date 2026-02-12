using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using FontAwesome.Sharp;

namespace DLack
{
    // ═══════════════════════════════════════════════════════════════
    //  SHARED HELPERS
    // ═══════════════════════════════════════════════════════════════

    internal static class DrawingHelpers
    {
        /// <summary>
        /// Creates a rounded rectangle path. Caller is responsible for disposal.
        /// </summary>
        public static GraphicsPath CreateRoundedRect(RectangleF rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;

            // Clamp diameter so it never exceeds the rect dimensions
            d = Math.Min(d, (int)Math.Min(rect.Width, rect.Height));

            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        /// <summary>
        /// Configures graphics for high-quality anti-aliased rendering.
        /// </summary>
        public static void SetHighQuality(Graphics g)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  MODERN BUTTON  –  Rounded with hover/press/disabled states
    // ═══════════════════════════════════════════════════════════════

    public class ModernButton : Button
    {
        private bool _isHovering;
        private bool _isPressed;

        public int CornerRadius { get; set; } = 8;
        public int HoverLighten { get; set; } = 20;
        public int PressDarken { get; set; } = 15;

        /// <summary>Custom disabled background. If not set, auto-derives from BackColor.</summary>
        public Color DisabledBackColor { get; set; } = Color.Empty;

        /// <summary>Custom disabled foreground. If not set, auto-derives from ForeColor.</summary>
        public Color DisabledForeColor { get; set; } = Color.Empty;

        public ModernButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Font = new Font("Segoe UI", 10, FontStyle.Bold);
            Cursor = Cursors.Hand;
            Size = new Size(150, 50);

            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.SupportsTransparentBackColor,
                true);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _isHovering = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _isHovering = false;
            _isPressed = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            _isPressed = true;
            Invalidate();
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            _isPressed = false;
            Invalidate();
            base.OnMouseUp(e);
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            Cursor = Enabled ? Cursors.Hand : Cursors.Default;
            Invalidate();
            base.OnEnabledChanged(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            DrawingHelpers.SetHighQuality(e.Graphics);

            bool disabled = !Enabled;
            Color bg = disabled ? GetDisabledBackColor() : GetStateColor();
            Color fg = disabled ? GetDisabledForeColor() : ForeColor;

            var rect = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);

            using (var path = DrawingHelpers.CreateRoundedRect(rect, CornerRadius))
            {
                using (var brush = new SolidBrush(bg))
                    e.Graphics.FillPath(brush, path);

                e.Graphics.SetClip(path);
            }

            // Let subclasses draw icon + text
            DrawContent(e.Graphics, fg, disabled);

            e.Graphics.ResetClip();
        }

        /// <summary>Override in subclasses to draw icon + text. Base draws centered text.</summary>
        protected virtual void DrawContent(Graphics g, Color foreColor, bool disabled)
        {
            TextRenderer.DrawText(
                g, Text, Font, ClientRectangle, foreColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private Color GetStateColor()
        {
            if (_isPressed)
                return ControlPaint.Dark(BackColor, PressDarken / 100f);
            if (_isHovering)
                return ControlPaint.Light(BackColor, HoverLighten / 100f);
            return BackColor;
        }

        private Color GetDisabledBackColor()
            => DisabledBackColor != Color.Empty
                ? DisabledBackColor
                : Color.FromArgb(226, 232, 240); // slate-200

        private Color GetDisabledForeColor()
            => DisabledForeColor != Color.Empty
                ? DisabledForeColor
                : Color.FromArgb(100, 116, 139); // slate-500
    }

    // ═══════════════════════════════════════════════════════════════
    //  ICON MODERN BUTTON  –  ModernButton + FontAwesome icon
    // ═══════════════════════════════════════════════════════════════

    public class IconModernButton : ModernButton
    {
        private IconChar _iconChar = IconChar.None;
        private Color _iconColor = Color.White;
        private int _iconSize = 20;
        private int _iconTextGap = 6;
        private Bitmap _iconCache;
        private IconChar _cachedChar;
        private Color _cachedColor;
        private int _cachedSize;

        public IconChar IconChar
        {
            get => _iconChar;
            set { _iconChar = value; InvalidateIconCache(); Invalidate(); }
        }

        public Color IconColor
        {
            get => _iconColor;
            set { _iconColor = value; InvalidateIconCache(); Invalidate(); }
        }

        public int IconSize
        {
            get => _iconSize;
            set { _iconSize = value; InvalidateIconCache(); Invalidate(); }
        }

        /// <summary>Pixel gap between icon and text.</summary>
        public int IconTextGap
        {
            get => _iconTextGap;
            set { _iconTextGap = value; Invalidate(); }
        }

        /// <summary>Custom disabled icon color. If not set, uses DisabledForeColor.</summary>
        public Color DisabledIconColor { get; set; } = Color.Empty;

        protected override void DrawContent(Graphics g, Color foreColor, bool disabled)
        {
            Color iconClr = disabled
                ? (DisabledIconColor != Color.Empty ? DisabledIconColor : foreColor)
                : _iconColor;

            var iconBmp = GetOrCreateIcon(iconClr);

            // Measure text without padding for accurate sizing
            string label = Text?.Trim() ?? "";
            var textSize = TextRenderer.MeasureText(g, label, Font,
                new Size(int.MaxValue, Height),
                TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);

            // Total content width: icon + gap + text
            int iconW = iconBmp?.Width ?? 0;
            int gap = iconBmp != null ? _iconTextGap : 0;
            int totalWidth = iconW + gap + textSize.Width;

            // Center the whole block horizontally
            int x = (Width - totalWidth) / 2;

            // Shared vertical center
            int centerY = Height / 2;

            // Draw icon — vertically centered
            if (iconBmp != null)
            {
                int iconY = centerY - iconBmp.Height / 2;
                g.DrawImage(iconBmp, x, iconY, iconBmp.Width, iconBmp.Height);
                x += iconBmp.Width + _iconTextGap;
            }

            // Draw text — vertically centered using the same center point
            int textY = centerY - textSize.Height / 2;
            var textRect = new Rectangle(x, textY, textSize.Width + 2, textSize.Height);
            TextRenderer.DrawText(g, label, Font, textRect, foreColor,
                TextFormatFlags.Left | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
        }

        private Bitmap GetOrCreateIcon(Color color)
        {
            if (_iconChar == IconChar.None) return null;

            // Return cached if still valid
            if (_iconCache != null &&
                _cachedChar == _iconChar &&
                _cachedColor.ToArgb() == color.ToArgb() &&
                _cachedSize == _iconSize)
            {
                return _iconCache;
            }

            _iconCache?.Dispose();
            _iconCache = RenderIconBitmap(_iconChar, _iconSize, color);
            _cachedChar = _iconChar;
            _cachedColor = color;
            _cachedSize = _iconSize;

            return _iconCache;
        }

        /// <summary>
        /// Renders a FontAwesome icon to a Bitmap using IconPictureBox,
        /// which handles all the font lookup internally.
        /// </summary>
        private static Bitmap RenderIconBitmap(IconChar icon, int size, Color color)
        {
            using var pic = new IconPictureBox
            {
                IconChar = icon,
                IconSize = size,
                IconColor = color,
                BackColor = Color.Transparent,
                Size = new Size(size, size)
            };

            return pic.Image != null
                ? new Bitmap(pic.Image)
                : null;
        }

        private void InvalidateIconCache()
        {
            _iconCache?.Dispose();
            _iconCache = null;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _iconCache?.Dispose();
                _iconCache = null;
            }
            base.Dispose(disposing);
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  MODERN CARD  –  Rounded white panel with soft drop shadow
    // ═══════════════════════════════════════════════════════════════

    public class ModernCard : Panel
    {
        public int CornerRadius { get; set; } = 12;
        public int ShadowDepth { get; set; } = 4;
        public Color ShadowColor { get; set; } = Color.FromArgb(30, 0, 0, 0);

        public ModernCard()
        {
            BackColor = Color.White;
            BorderStyle = BorderStyle.None;

            // Required for transparent shadow painting
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.SupportsTransparentBackColor |
                ControlStyles.ResizeRedraw,
                true);

            // Add internal padding so child controls don't overlap shadow area
            Padding = new Padding(ShadowDepth);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            DrawingHelpers.SetHighQuality(e.Graphics);

            // ── Shadow layers (offset, progressively fading) ──
            for (int i = ShadowDepth; i >= 1; i--)
            {
                int alpha = (int)(ShadowColor.A * ((float)i / ShadowDepth) * 0.5f);
                var shadowRect = new RectangleF(
                    i, i,
                    Width - i * 2f, Height - i * 2f);

                using (var path = DrawingHelpers.CreateRoundedRect(shadowRect, CornerRadius))
                using (var pen = new Pen(Color.FromArgb(alpha, ShadowColor), 1.5f))
                    e.Graphics.DrawPath(pen, path);
            }

            // ── Card body ──
            var bodyRect = new RectangleF(
                ShadowDepth, ShadowDepth,
                Width - ShadowDepth * 2f, Height - ShadowDepth * 2f);

            using (var path = DrawingHelpers.CreateRoundedRect(bodyRect, CornerRadius))
            {
                using (var brush = new SolidBrush(BackColor))
                    e.Graphics.FillPath(brush, path);

                // Subtle border
                using (var pen = new Pen(Color.FromArgb(20, 0, 0, 0), 1f))
                    e.Graphics.DrawPath(pen, path);
            }
        }

        /// <summary>
        /// Prevent the default background erase to avoid flicker.
        /// </summary>
        protected override void OnPaintBackground(PaintEventArgs e) { }
    }

    // ═══════════════════════════════════════════════════════════════
    //  MODERN PROGRESS BAR  –  Rounded with gradient fill + animation
    // ═══════════════════════════════════════════════════════════════

    public class ModernProgressBar : Panel
    {
        // ── Appearance ──
        public Color BarColor { get; set; } = Color.FromArgb(59, 130, 246);
        public Color BarEndColor { get; set; } = Color.FromArgb(99, 160, 255);
        public Color TrackColor { get; set; } = Color.FromArgb(226, 232, 240);
        public int CornerRadius { get; set; } = 3;

        // ── Value ──
        private int _value;
        private float _displayValue;
        private Timer _animTimer;

        public int Value
        {
            get => _value;
            set
            {
                _value = Math.Clamp(value, 0, 100);
                EnsureAnimationTimer();
                _animTimer.Start();
            }
        }

        public ModernProgressBar()
        {
            Height = 6;
            BackColor = Color.Transparent;

            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.ResizeRedraw,
                true);
        }

        private void EnsureAnimationTimer()
        {
            if (_animTimer != null) return;

            _animTimer = new Timer { Interval = 16 }; // ~60 fps
            _animTimer.Tick += (s, e) =>
            {
                float diff = _value - _displayValue;

                if (Math.Abs(diff) < 0.5f)
                {
                    _displayValue = _value;
                    _animTimer.Stop();
                }
                else
                {
                    // Ease towards target
                    _displayValue += diff * 0.15f;
                }

                Invalidate();
            };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            DrawingHelpers.SetHighQuality(e.Graphics);

            var trackRect = new RectangleF(0, 0, Width, Height);

            // ── Track ──
            using (var path = DrawingHelpers.CreateRoundedRect(trackRect, CornerRadius))
            using (var brush = new SolidBrush(TrackColor))
                e.Graphics.FillPath(brush, path);

            // ── Fill ──
            if (_displayValue > 0.5f)
            {
                float fillWidth = Math.Max(Height, _displayValue / 100f * Width);
                var fillRect = new RectangleF(0, 0, fillWidth, Height);

                using (var path = DrawingHelpers.CreateRoundedRect(fillRect, CornerRadius))
                using (var brush = new LinearGradientBrush(
                    fillRect, BarColor, BarEndColor, LinearGradientMode.Horizontal))
                {
                    e.Graphics.FillPath(brush, path);
                }
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e) { }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _animTimer?.Stop();
                _animTimer?.Dispose();
                _animTimer = null;
            }
            base.Dispose(disposing);
        }
    }
}