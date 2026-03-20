using System;
using System.Threading;
using System.Windows.Forms;

namespace FocusShield
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            // Single-instance guard
            using var mutex = new Mutex(true, "FocusShield_SingleInstance", out bool first);
            if (!first)
            {
                MessageBox.Show(
                    "FocusShield is already running.\nCheck your system tray.",
                    "FocusShield", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Environment.OSVersion.Version.Major < 6)
            {
                MessageBox.Show(
                    "FocusShield requires Windows Vista or later.",
                    "Unsupported OS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.ApplicationExit += (_, _) =>
                FocusShieldForm.Instance.TrayIconVisible = false;

            Application.Run(FocusShieldForm.Instance);
        }
    }
}
