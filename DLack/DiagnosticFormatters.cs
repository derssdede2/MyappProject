using System;
using System.Linq;

namespace DLack
{
    /// <summary>
    /// Shared formatting and rating helpers used by both the UI and PDF report.
    /// </summary>
    internal static class DiagnosticFormatters
    {
        /// <summary>Application name and version displayed in the UI and PDF reports.</summary>
        public const string AppVersion = "Endpoint Diagnostics & Remediation v2.0";
        /// <summary>Formats a <see cref="TimeSpan"/> uptime into a compact "Xd Xh Xm" string.</summary>
        public static string FormatUptime(TimeSpan uptime)
        {
            if (uptime.TotalDays >= 1)
                return $"{uptime.Days}d {uptime.Hours}h {uptime.Minutes}m";
            if (uptime.TotalHours >= 1)
                return $"{uptime.Hours}h {uptime.Minutes}m";
            return $"{uptime.Minutes}m";
        }

        /// <summary>Rate DNS response: 0=Good (&lt;50 ms), 1=Fair (50-150 ms), 2=Poor (&gt;150 ms).</summary>
        public static int RateDns(string dnsResponse)
        {
            if (dnsResponse == "Failed" || dnsResponse == "Unknown") return 2;
            string digits = new string(dnsResponse.Where(char.IsDigit).ToArray());
            if (!int.TryParse(digits, out int ms)) return 1;
            if (ms > 150) return 2;
            if (ms > 50) return 1;
            return 0;
        }

        /// <summary>Rate ping latency: 0=Good (&lt;30 ms), 1=Fair (30-100 ms), 2=Poor (&gt;100 ms).</summary>
        public static int RatePing(string pingLatency)
        {
            if (pingLatency == "Failed" || pingLatency == "Timed out") return 2;
            if (pingLatency == "Unknown") return 1;
            string digits = new string(pingLatency.Where(char.IsDigit).ToArray());
            if (!int.TryParse(digits, out int ms)) return 1;
            if (ms > 100) return 2;
            if (ms > 30) return 1;
            return 0;
        }

        /// <summary>Formats a megabyte value as "X.X GB" when ≥1024 MB, otherwise "X MB".</summary>
        public static string FormatMB(long mb) => mb >= 1024 ? $"{mb / 1024.0:F1} GB" : $"{mb:N0} MB";
    }
}
