using System;
using System.Drawing;
using System.Windows.Forms;

namespace FocusShield
{
    /// <summary>
    /// Invisible message window that owns the tray icon and receives shell hook messages.
    /// </summary>
    internal class FocusShieldForm : Form
    {
        // ─── singleton ───────────────────────────────────────────────────────────
        private static FocusShieldForm _instance;
        public static FocusShieldForm Instance => _instance ??= new FocusShieldForm();

        // ─── tray ────────────────────────────────────────────────────────────────
        public bool TrayIconVisible
        {
            get => _trayIcon.Visible;
            set => _trayIcon.Visible = value;
        }

        // ─── fields ──────────────────────────────────────────────────────────────
        private readonly NotifyIcon       _trayIcon;
        private readonly ContextMenuStrip _menu;
        private readonly ToolStripMenuItem _menuEnabled;

        private uint   _shellMsg;              // registered WM_SHELLHOOKMESSAGE id
        private IntPtr _lastUserForeground;    // hwnd the user was actually using
        private bool   _enabled = true;
        private uint   _originalTimeout;

        private Icon _iconActive;
        private Icon _iconBlocked;
        private Icon _iconPaused;

        private readonly System.Windows.Forms.Timer _resetTimer;

        // ─── ctor ────────────────────────────────────────────────────────────────
        private FocusShieldForm()
        {
            // Make the form invisible and excluded from Alt+Tab / taskbar
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar   = false;
            WindowState     = FormWindowState.Minimized;
            Opacity         = 0;
            Size            = new Size(1, 1);

            _iconActive  = IconRenderer.CreateActiveIcon();
            _iconBlocked = IconRenderer.CreateBlockedIcon();
            _iconPaused  = IconRenderer.CreatePausedIcon();

            // ── context menu ──
            _menuEnabled = new ToolStripMenuItem("Protection Enabled", null, OnToggleEnabled)
                { Checked = true };
            var menuExit = new ToolStripMenuItem("Exit", null, OnExit);

            _menu = new ContextMenuStrip();
            _menu.Items.Add(_menuEnabled);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(menuExit);

            // ── tray icon ──
            _trayIcon = new NotifyIcon
            {
                Icon             = _iconActive,
                Text             = "FocusShield \u2014 protecting focus",
                Visible          = true,
                ContextMenuStrip = _menu
            };

            // ── timer to reset icon after a block event ──
            _resetTimer = new System.Windows.Forms.Timer { Interval = 2500 };
            _resetTimer.Tick += (_, _) =>
            {
                _resetTimer.Stop();
                RefreshTrayIcon();
            };

            Load        += OnLoad;
            FormClosing += OnFormClosing;
        }

        // ─── lifecycle ───────────────────────────────────────────────────────────
        private void OnLoad(object sender, EventArgs e)
        {
            // Hide the form properly on the next pump cycle
            BeginInvoke((Action)Hide);

            // Save the current system timeout so we can restore it on exit
            NativeMethods.SystemParametersInfo(
                NativeMethods.SPI_GETFOREGROUNDLOCKTIMEOUT, 0, ref _originalTimeout, 0);

            // Raise it high so Windows itself blocks focus steals
            ApplyLockTimeout(30_000u);

            // Register this window to receive WM_SHELLHOOKMESSAGE notifications
            _shellMsg = NativeMethods.RegisterWindowMessage("SHELLHOOK");
            NativeMethods.RegisterShellHookWindow(Handle);
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            NativeMethods.DeregisterShellHookWindow(Handle);

            // Restore the original timeout
            ApplyLockTimeout(_originalTimeout);

            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _iconActive.Dispose();
            _iconBlocked.Dispose();
            _iconPaused.Dispose();
        }

        // ─── shell hook message pump ─────────────────────────────────────────────
        protected override void WndProc(ref Message m)
        {
            if (_shellMsg != 0 && m.Msg == (int)_shellMsg)
            {
                int code = m.WParam.ToInt32();

                if (code == NativeMethods.HSHELL_WINDOWACTIVATED && _enabled)
                {
                    // Normal (user-initiated) activation: remember this window
                    if (m.LParam != IntPtr.Zero)
                        _lastUserForeground = m.LParam;
                }
                else if (code == NativeMethods.HSHELL_RUDEAPPACTIVATED && _enabled)
                {
                    OnRudeActivation(m.LParam);
                }
            }

            base.WndProc(ref m);
        }

        private void OnRudeActivation(IntPtr rudeHwnd)
        {
            if (rudeHwnd == IntPtr.Zero) return;

            // Ignore our own process
            NativeMethods.GetWindowThreadProcessId(rudeHwnd, out uint pid);
            if (pid == NativeMethods.GetCurrentProcessId()) return;

            // 1. Flash the rude window's taskbar button (amber, until user clicks it)
            NativeMethods.FlashTaskbar(rudeHwnd);

            // 2. Return focus to whatever the user was doing
            if (_lastUserForeground != IntPtr.Zero && _lastUserForeground != rudeHwnd)
                NativeMethods.ForceForeground(_lastUserForeground);

            // 3. Update tray to amber + balloon
            string title = NativeMethods.GetWindowTitle(rudeHwnd);
            string label = string.IsNullOrWhiteSpace(title) ? "a background app" : $"\"{title}\"";

            _trayIcon.Icon = _iconBlocked;
            _trayIcon.Text = TruncateTip($"Blocked: {label}");
            _trayIcon.ShowBalloonTip(
                timeout: 3000,
                tipTitle: "FocusShield blocked a pop-up",
                tipText:  $"{label} wants your attention \u2014 click its taskbar button when ready.",
                tipIcon:  ToolTipIcon.Info);

            _resetTimer.Stop();
            _resetTimer.Start();
        }

        // ─── menu handlers ───────────────────────────────────────────────────────
        private void OnToggleEnabled(object sender, EventArgs e)
        {
            _enabled = !_enabled;
            _menuEnabled.Checked = _enabled;

            ApplyLockTimeout(_enabled ? 30_000u : _originalTimeout);
            RefreshTrayIcon();
        }

        private void OnExit(object sender, EventArgs e) => Application.Exit();

        // ─── helpers ─────────────────────────────────────────────────────────────
        private static void ApplyLockTimeout(uint ms)
        {
            NativeMethods.SystemParametersInfo(
                NativeMethods.SPI_SETFOREGROUNDLOCKTIMEOUT, 0,
                (IntPtr)ms, NativeMethods.SPIF_SENDCHANGE);
        }

        private void RefreshTrayIcon()
        {
            _trayIcon.Icon = _enabled ? _iconActive : _iconPaused;
            _trayIcon.Text = _enabled
                ? "FocusShield \u2014 protecting focus"
                : "FocusShield \u2014 paused";
        }

        // NotifyIcon.Text has a 64-char limit
        private static string TruncateTip(string s) =>
            s.Length > 63 ? s.Substring(0, 60) + "\u2026" : s;

        // Exclude from Alt+Tab switcher
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x80; // WS_EX_TOOLWINDOW
                return cp;
            }
        }
    }
}
