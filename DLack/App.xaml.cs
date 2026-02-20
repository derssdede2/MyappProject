using System;
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

                // Load/apply theme before StartupUri creates MainWindow so controls
                // (including theme-toggle icon) initialize against the persisted mode.
                ThemeManager.Initialize();
                base.OnStartup(e);
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
