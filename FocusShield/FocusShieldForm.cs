using System;
using System.Collections.Generic;
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
        private readonly AppListManager    _appListManager;

        private uint   _shellMsg;              // registered WM_SHELLHOOKMESSAGE id
        private IntPtr _lastUserForeground;    // hwnd the user was actually using
        private bool   _enabled = true;
        private uint   _originalTimeout;

        // If the user hasn't touched the keyboard/mouse in this many ms,
        // any window activation is treated as a focus steal.
        private const uint IdleThresholdMs = 300;

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

            _appListManager = new AppListManager();

            // ── context menu ──
            _menuEnabled = new ToolStripMenuItem("Protection Enabled", null, OnToggleEnabled)
                { Checked = true };
            var menuWhitelist = new ToolStripMenuItem("Whitelist...", null, OnEditWhitelist);
            var menuBlacklist = new ToolStripMenuItem("Blacklist...", null, OnEditBlacklist);
            var menuExit = new ToolStripMenuItem("Exit", null, OnExit);

            _menu = new ContextMenuStrip();
            _menu.Items.Add(_menuEnabled);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(menuWhitelist);
            _menu.Items.Add(menuBlacklist);
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
            if (_shellMsg != 0 && m.Msg == (int)_shellMsg && _enabled)
            {
                int code = m.WParam.ToInt32();

                if (code == NativeMethods.HSHELL_RUDEAPPACTIVATED)
                {
                    // Always block forced activations
                    OnFocusSteal(m.LParam);
                }
                else if (code == NativeMethods.HSHELL_WINDOWACTIVATED)
                {
                    if (m.LParam == IntPtr.Zero) goto done;

                    // If user was idle, this is likely an app stealing focus
                    uint idle = NativeMethods.GetIdleTime();
                    if (idle > IdleThresholdMs && IsDifferentProcess(m.LParam))
                    {
                        OnFocusSteal(m.LParam);
                    }
                    else
                    {
                        // Genuine user-initiated activation: remember it
                        _lastUserForeground = m.LParam;
                    }
                }
            }

            done:
            base.WndProc(ref m);
        }

        private bool IsDifferentProcess(IntPtr hwnd)
        {
            if (_lastUserForeground == IntPtr.Zero) return false;
            NativeMethods.GetWindowThreadProcessId(hwnd, out uint newPid);
            NativeMethods.GetWindowThreadProcessId(_lastUserForeground, out uint oldPid);
            return newPid != oldPid;
        }

        private void OnFocusSteal(IntPtr rudeHwnd)
        {
            if (rudeHwnd == IntPtr.Zero) return;

            // Ignore our own process
            NativeMethods.GetWindowThreadProcessId(rudeHwnd, out uint pid);
            if (pid == NativeMethods.GetCurrentProcessId()) return;

            // Check whitelist: if app is whitelisted, allow it to take focus
            if (_appListManager.IsWhitelisted(pid))
            {
                _lastUserForeground = rudeHwnd;
                return;
            }

            // Blacklist always blocks, even if user initiated
            // (useful for known problematic apps)
            bool isBlacklisted = _appListManager.IsBlacklisted(pid);

            // 1. Flash the rude window's taskbar button (amber, until user clicks it)
            NativeMethods.FlashTaskbar(rudeHwnd);

            // 2. Return focus to whatever the user was doing
            if (_lastUserForeground != IntPtr.Zero && _lastUserForeground != rudeHwnd)
                NativeMethods.ForceForeground(_lastUserForeground);

            // 3. Update tray to amber + notification
            string title = NativeMethods.GetWindowTitle(rudeHwnd);
            string appName = string.IsNullOrWhiteSpace(title) ? "An app" : $"{title}";

            _trayIcon.Icon = _iconBlocked;
            _trayIcon.Text = TruncateTip($"Ready: {appName}");

            string tipText = isBlacklisted
                ? $"{appName} is blacklisted and blocked."
                : "Click its taskbar button when you\u2019re ready to switch.";

            _trayIcon.ShowBalloonTip(
                timeout: 3000,
                tipTitle: $"{appName} is ready",
                tipText:  tipText,
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

        private void OnEditWhitelist(object sender, EventArgs e)
        {
            ShowListEditor("Whitelist", _appListManager.GetWhitelist(),
                _appListManager.AddToWhitelist, _appListManager.RemoveFromWhitelist);
        }

        private void OnEditBlacklist(object sender, EventArgs e)
        {
            ShowListEditor("Blacklist", _appListManager.GetBlacklist(),
                _appListManager.AddToBlacklist, _appListManager.RemoveFromBlacklist);
        }

        private void ShowListEditor(string listName, IEnumerable<string> items,
            Action<string> onAdd, Action<string> onRemove)
        {
            var form = new Form
            {
                Text = $"FocusShield - {listName}",
                Width = 400,
                Height = 300,
                StartPosition = FormStartPosition.CenterScreen,
                ShowIcon = false
            };

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                DataSource = new List<string>(items)
            };
            form.Controls.Add(listBox);

            var panel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            var btnAdd = new Button
            {
                Text = "Add",
                Width = 80,
                Left = 10,
                Top = 5,
                DialogResult = DialogResult.OK
            };
            btnAdd.Click += (_, _) =>
            {
                string app = PromptForAppName();
                if (!string.IsNullOrWhiteSpace(app))
                {
                    onAdd(app);
                    listBox.DataSource = new List<string>(items);
                }
            };

            var btnRemove = new Button
            {
                Text = "Remove",
                Width = 80,
                Left = 100,
                Top = 5,
                DialogResult = DialogResult.Cancel
            };
            btnRemove.Click += (_, _) =>
            {
                if (listBox.SelectedItem is string selected)
                {
                    onRemove(selected);
                    listBox.DataSource = new List<string>(items);
                }
            };

            var btnClose = new Button
            {
                Text = "Close",
                Width = 80,
                Left = 190,
                Top = 5,
                DialogResult = DialogResult.Cancel
            };
            btnClose.Click += (_, _) => form.Close();

            panel.Controls.Add(btnAdd);
            panel.Controls.Add(btnRemove);
            panel.Controls.Add(btnClose);
            form.Controls.Add(panel);

            form.ShowDialog();
        }

        private string PromptForAppName()
        {
            var form = new Form
            {
                Text = "Add Application",
                Width = 300,
                Height = 150,
                StartPosition = FormStartPosition.CenterParent,
                ShowIcon = false
            };

            var label = new Label
            {
                Text = "Application name (without .exe):",
                Left = 10,
                Top = 10,
                Width = 260
            };

            var textBox = new TextBox
            {
                Left = 10,
                Top = 40,
                Width = 260
            };

            var btnOK = new Button
            {
                Text = "OK",
                Left = 110,
                Top = 70,
                Width = 80,
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Left = 200,
                Top = 70,
                Width = 80,
                DialogResult = DialogResult.Cancel
            };

            form.Controls.Add(label);
            form.Controls.Add(textBox);
            form.Controls.Add(btnOK);
            form.Controls.Add(btnCancel);
            form.AcceptButton = btnOK;
            form.CancelButton = btnCancel;

            return form.ShowDialog() == DialogResult.OK ? textBox.Text : null;
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
