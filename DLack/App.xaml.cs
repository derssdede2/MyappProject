using System;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Windows;

namespace DLack
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                if (!IsRunningAsAdmin())
                {
                    MessageBox.Show(
                        "Endpoint Diagnostics & Remediation requires Administrator privileges to scan system health.\n\n" +
                        "Please run this application as Administrator.",
                        "Admin Rights Required",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    Shutdown();
                    return;
                }

                base.OnStartup(e);
                ThemeManager.Initialize();

                // Silent mode: /scan /export — run scan and export PDF without UI
                var args = e.Args.Select(a => a.ToLowerInvariant()).ToArray();
                if (args.Contains("/scan") && args.Contains("/export"))
                {
                    RunSilentAsync();
                    return;
                }

                // Normal interactive mode
                var mainWindow = new MainWindow();
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"STARTUP ERROR:\n\n{ex.Message}\n\n{ex.StackTrace}",
                    "DLack Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Shutdown();
            }
        }

        /// <summary>
        /// Silent scan + PDF export for scheduled/automated runs.
        /// Usage: DLack.exe /scan /export
        /// </summary>
        private async void RunSilentAsync()
        {
            try
            {
                var scanner = new Scanner();
                var result = await scanner.RunScan();
                if (result == null)
                {
                    Environment.ExitCode = 1;
                    Shutdown();
                    return;
                }

                // Save scan history
                ScanHistoryManager.Save(ScanHistoryManager.CreateSnapshot(result));

                // Export PDF to Desktop
                string filename = $"EDR_Diagnostic_{Environment.MachineName}_{DateTime.Now:yyyyMMdd_HHmm}.pdf";
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fullPath = Path.Combine(desktopPath, filename);

                new PDFReportGenerator().Generate(fullPath, result);

                Environment.ExitCode = 0;
            }
            catch
            {
                Environment.ExitCode = 2;
            }
            finally
            {
                Shutdown();
            }
        }

        private static bool IsRunningAsAdmin()
        {
            try
            {
                var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }
    }
}
