using System;

namespace DLack
{
    // ═══════════════════════════════════════════════════════════════
    //  ACTION TYPE
    // ═══════════════════════════════════════════════════════════════

    public enum ActionType
    {
        /// <summary>Runs automatically — deletes files, changes settings, runs commands.</summary>
        AutoFix,
        /// <summary>Opens a settings page or external tool — requires user action.</summary>
        Shortcut,
        /// <summary>Informational only — no automation possible.</summary>
        ManualOnly
    }

    // ═══════════════════════════════════════════════════════════════
    //  RISK LEVELS
    // ═══════════════════════════════════════════════════════════════

    public enum ActionRisk
    {
        Safe,
        Moderate,
        RequiresReboot
    }

    // ═══════════════════════════════════════════════════════════════
    //  IMPACT LEVEL (how likely this action is to help)
    // ═══════════════════════════════════════════════════════════════

    public enum ImpactLevel
    {
        /// <summary>Low confidence this will noticeably help.</summary>
        Low,
        /// <summary>Moderate confidence — common fix for this class of issue.</summary>
        Medium,
        /// <summary>High confidence — directly addresses detected problem.</summary>
        High
    }

    // ═══════════════════════════════════════════════════════════════
    //  ACTION STATUS
    // ═══════════════════════════════════════════════════════════════

    public enum ActionStatus
    {
        Pending,
        Running,
        Success,
        PartialSuccess,
        NoChange,
        Failed,
        Skipped
    }

    // ═══════════════════════════════════════════════════════════════
    //  VERIFICATION STATUS
    // ═══════════════════════════════════════════════════════════════

    public enum VerificationStatus
    {
        /// <summary>Not yet verified.</summary>
        None,
        /// <summary>Verification is running.</summary>
        Verifying,
        /// <summary>The fix was confirmed effective.</summary>
        Verified,
        /// <summary>Partial improvement detected.</summary>
        PartiallyVerified,
        /// <summary>The fix did not resolve the issue.</summary>
        NotVerified,
        /// <summary>Verification requires a reboot to complete.</summary>
        RequiresReboot
    }

    // ═══════════════════════════════════════════════════════════════
    //  OPTIMIZATION ACTION
    // ═══════════════════════════════════════════════════════════════

    public class OptimizationAction
    {
        public string Category { get; set; } = "";
        public string Title { get; set; } = "";
        public string Description { get; set; } = "";
        public bool IsAutomatable { get; set; }
        public ActionType Type { get; set; } = ActionType.AutoFix;
        public ActionRisk Risk { get; set; } = ActionRisk.Safe;
        public long EstimatedFreeMB { get; set; }
        public ActionStatus Status { get; set; } = ActionStatus.Pending;
        public string ResultMessage { get; set; } = "";
        public long ActualFreedMB { get; set; }
        public bool IsSelected { get; set; }

        /// <summary>
        /// Approximate wall-clock time shown to the user (e.g. "~1-5 min").
        /// </summary>
        public string EstimatedDuration { get; set; } = "";

        /// <summary>
        /// The key used by the Optimizer to dispatch execution.
        /// </summary>
        public string ActionKey { get; set; } = "";

        /// <summary>
        /// Post-optimization verification result.
        /// </summary>
        public VerificationStatus Verification { get; set; } = VerificationStatus.None;

        /// <summary>
        /// Human-readable verification detail (e.g. "Temp folder: 12 MB (was 1.2 GB)").
        /// </summary>
        public string VerificationMessage { get; set; } = "";

        /// <summary>
        /// How likely this action is to produce a visible improvement.
        /// </summary>
        public ImpactLevel Impact { get; set; } = ImpactLevel.Medium;

        /// <summary>
        /// Short rationale for the impact assessment (shown to operator).
        /// </summary>
        public string ImpactRationale { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    //  OPTIMIZATION PROGRESS
    // ═══════════════════════════════════════════════════════════════

    public class OptimizationProgress
    {
        public int Current { get; set; }
        public int Total { get; set; }
        public string CurrentAction { get; set; } = "";
        public bool IsLongRunning { get; set; }

        /// <summary>
        /// Approximate wall-clock seconds for countdown display. Zero means no estimate available.
        /// </summary>
        public int EstimatedSeconds { get; set; }
    }
}
