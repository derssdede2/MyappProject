using System;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;

namespace DLack
{
    public static class ThemeManager
    {
        public static event Action ThemeChanged;

        private static bool _isDark;

        public static bool IsDark
        {
            get => _isDark;
            set
            {
                if (_isDark == value) return;
                _isDark = value;
                Apply();
                Save();
                ThemeChanged?.Invoke();
            }
        }

        // ── Text Colors ──
        public static Color TextPrimary { get; private set; }
        public static Color TextMuted { get; private set; }
        public static Color TextSubtle { get; private set; }

        // ── Surface Colors ──
        public static Color Surface { get; private set; }
        public static Color SurfaceAlt { get; private set; }
        public static Color Border { get; private set; }

        // ── Accent Colors ──
        public static Color AccentBlue { get; private set; }
        public static Color AccentGreen { get; private set; }
        public static Color AccentRed { get; private set; }
        public static Color AccentAmber { get; private set; }

        // ── Severity Colors ──
        public static Color GoodFg { get; private set; }
        public static Color GoodBg { get; private set; }
        public static Color WarnFg { get; private set; }
        public static Color WarnBg { get; private set; }
        public static Color CritFg { get; private set; }
        public static Color CritBg { get; private set; }

        // ── Status Brushes ──
        public static SolidColorBrush BrushPrimary { get; private set; }
        public static SolidColorBrush BrushSuccess { get; private set; }
        public static SolidColorBrush BrushDanger { get; private set; }

        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DLack", "theme.json");

        public static void Initialize()
        {
            Load();
            Apply();
        }

        private static void Apply()
        {
            var app = Application.Current;
            if (app == null) return;

            // Swap XAML resource dictionary
            var dicts = app.Resources.MergedDictionaries;
            for (int i = dicts.Count - 1; i >= 0; i--)
            {
                var src = dicts[i].Source?.ToString() ?? "";
                if (src.Contains("Professional"))
                    dicts.RemoveAt(i);
            }

            string themeUri = _isDark
                ? "Themes/ProfessionalDark.xaml"
                : "Themes/ProfessionalLight.xaml";
            dicts.Insert(0, new ResourceDictionary
            {
                Source = new Uri(themeUri, UriKind.Relative)
            });

            if (_isDark)
                ApplyDarkColors();
            else
                ApplyLightColors();

            BrushPrimary = new SolidColorBrush(AccentBlue);
            BrushSuccess = new SolidColorBrush(AccentGreen);
            BrushDanger = new SolidColorBrush(AccentRed);
            BrushPrimary.Freeze();
            BrushSuccess.Freeze();
            BrushDanger.Freeze();
        }

        private static void ApplyLightColors()
        {
            TextPrimary = Color.FromRgb(17, 24, 39);     // #111827
            TextMuted = Color.FromRgb(107, 114, 128);    // #6B7280
            TextSubtle = Color.FromRgb(156, 163, 175);   // #9CA3AF
            Surface = Color.FromRgb(255, 255, 255);      // #FFFFFF
            SurfaceAlt = Color.FromRgb(249, 250, 251);   // #F9FAFB
            Border = Color.FromRgb(229, 231, 235);       // #E5E7EB

            AccentBlue = Color.FromRgb(155, 77, 43);     // #9B4D2B brand sienna
            AccentGreen = Color.FromRgb(27, 122, 61);    // #1B7A3D
            AccentRed = Color.FromRgb(185, 28, 28);      // #B91C1C
            AccentAmber = Color.FromRgb(180, 130, 10);

            GoodFg = Color.FromRgb(27, 122, 61);
            GoodBg = Color.FromRgb(220, 252, 231);
            WarnFg = Color.FromRgb(180, 130, 10);
            WarnBg = Color.FromRgb(254, 249, 195);
            CritFg = Color.FromRgb(185, 28, 28);
            CritBg = Color.FromRgb(254, 226, 226);
        }

        private static void ApplyDarkColors()
        {
            TextPrimary = Color.FromRgb(241, 245, 249);  // #F1F5F9
            TextMuted = Color.FromRgb(148, 163, 184);    // #94A3B8
            TextSubtle = Color.FromRgb(100, 116, 139);   // #64748B
            Surface = Color.FromRgb(28, 28, 28);         // #1C1C1C
            SurfaceAlt = Color.FromRgb(20, 20, 20);      // #141414
            Border = Color.FromRgb(46, 46, 46);          // #2E2E2E

            AccentBlue = Color.FromRgb(212, 145, 90);    // #D4915A warm sienna
            AccentGreen = Color.FromRgb(74, 222, 128);   // #4ADE80
            AccentRed = Color.FromRgb(248, 113, 113);    // #F87171
            AccentAmber = Color.FromRgb(251, 191, 36);

            GoodFg = Color.FromRgb(74, 222, 128);        // #4ADE80
            GoodBg = Color.FromRgb(20, 50, 30);
            WarnFg = Color.FromRgb(251, 191, 36);
            WarnBg = Color.FromRgb(55, 42, 12);
            CritFg = Color.FromRgb(248, 113, 113);
            CritBg = Color.FromRgb(60, 24, 24);
        }

        private static void Save()
        {
            try
            {
                string dir = Path.GetDirectoryName(SettingsPath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(SettingsPath,
                    JsonSerializer.Serialize(new { DarkMode = _isDark }));
            }
            catch { }
        }

        private static void Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var json = JsonSerializer.Deserialize<JsonElement>(
                        File.ReadAllText(SettingsPath));
                    _isDark = json.GetProperty("DarkMode").GetBoolean();
                }
            }
            catch { _isDark = false; }
        }
    }
}
