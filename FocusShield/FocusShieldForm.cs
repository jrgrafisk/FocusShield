using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace FocusShield
{
    /// <summary>
    /// Invisible message window that owns the tray icon, listens to shell hook
    /// notifications and decides whether a window is allowed to take over the
    /// screen. Programs keep loading normally — they just do it behind whatever
    /// the user is working on.
    /// </summary>
    internal class FocusShieldForm : Form
    {
        // ─── singleton ───────────────────────────────────────────────────────────
        private static FocusShieldForm _instance;
        public static FocusShieldForm Instance => _instance ??= new FocusShieldForm();

        // ─── tray ────────────────────────────────────────────────────────────────

        // ApplicationExit runs after FormClosing has already disposed the icon, so
        // the setter has to tolerate being called once the tray is gone.
        public bool TrayIconVisible
        {
            get => !_trayDisposed && _trayIcon.Visible;
            set
            {
                if (!_trayDisposed)
                    _trayIcon.Visible = value;
            }
        }

        // ─── tuning ──────────────────────────────────────────────────────────────

        // A window counts as "just finished loading" for this long after the shell
        // announced it. Slow starters put their window up well after the process
        // was created, so this is measured from the window, not the process.
        private const int NewWindowGraceMs = 4000;

        // How long the tray icon stays amber after an interception.
        private const int IconResetMs = 2500;

        // Backstop poll for focus changes the shell hook never reports.
        private const int PollIntervalMs = 200;

        // At most one balloon per window per this many ms.
        private const int NotifyCooldownMs = 8000;

        // An app that keeps shoving itself forward is left alone for a while rather
        // than trading the foreground back and forth with us forever.
        private const int MaxInterceptions    = 8;
        private const int InterceptionResetMs = 15000;
        private const int BackOffMs           = 30000;

        // What FOREGROUNDLOCKTIMEOUT is raised to while protection is on, so that
        // Windows itself turns most focus steals into a taskbar flash.
        private const uint LockTimeoutMs = 30000;

        private const string RunKeyPath   = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "FocusShield";

        // ─── fields ──────────────────────────────────────────────────────────────
        private readonly NotifyIcon        _trayIcon;
        private readonly ContextMenuStrip  _menu;
        private readonly ToolStripMenuItem _menuEnabled;
        private readonly ToolStripMenuItem _menuBackground;
        private readonly ToolStripMenuItem _menuNotify;
        private readonly ToolStripMenuItem _menuBootload;
        private readonly ConfigManager     _config;

        private readonly System.Windows.Forms.Timer _resetTimer;
        private readonly System.Windows.Forms.Timer _pollTimer;

        // hwnd → tick when the shell said the window was created
        private readonly Dictionary<IntPtr, int> _created = new();

        // hwnd → how often we have had to push this window back
        private readonly Dictionary<IntPtr, Interception> _seen = new();

        private uint   _shellMsg;           // registered WM_SHELLHOOKMESSAGE id
        private IntPtr _userForeground;     // the window the user is actually working in
        private IntPtr _lastForeground;     // last foreground window we looked at
        private uint   _originalTimeout;
        private bool   _ready;
        private bool   _trayDisposed;

        private readonly Icon _iconActive;
        private readonly Icon _iconBlocked;
        private readonly Icon _iconPaused;

        private sealed class Interception
        {
            public int Count;
            public int LastTick;
            public int NotifiedTick;
            public int BackOffUntil;
        }

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

            _config = new ConfigManager();

            // ── context menu ──
            _menuEnabled = new ToolStripMenuItem("Protection Enabled", null, OnToggleEnabled)
                { Checked = _config.ProtectionEnabled };
            _menuBackground = new ToolStripMenuItem(
                "Keep New Windows in Background", null, OnToggleBackgroundNewWindows)
                { Checked = _config.BackgroundNewWindows };
            _menuNotify = new ToolStripMenuItem("Show Notifications", null, OnToggleNotifications)
                { Checked = _config.ShowNotifications };
            var menuWhitelist = new ToolStripMenuItem("Whitelist...", null, OnEditWhitelist);
            var menuBlacklist = new ToolStripMenuItem("Blacklist...", null, OnEditBlacklist);
            _menuBootload = new ToolStripMenuItem("Load on Boot", null, OnToggleBootload)
                { Checked = IsBootloadEnabled() };
            var menuExit = new ToolStripMenuItem("Exit", null, OnExit);

            _menu = new ContextMenuStrip();
            _menu.Items.Add(_menuEnabled);
            _menu.Items.Add(_menuBackground);
            _menu.Items.Add(_menuNotify);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(menuWhitelist);
            _menu.Items.Add(menuBlacklist);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_menuBootload);
            _menu.Items.Add(menuExit);

            // ── tray icon ──
            _trayIcon = new NotifyIcon
            {
                Icon             = _config.ProtectionEnabled ? _iconActive : _iconPaused,
                Text             = TrayTooltip(),
                Visible          = true,
                ContextMenuStrip = _menu
            };
            _trayIcon.DoubleClick += OnToggleEnabled;

            // ── timer to reset icon after an interception ──
            _resetTimer = new System.Windows.Forms.Timer { Interval = IconResetMs };
            _resetTimer.Tick += (_, _) =>
            {
                _resetTimer.Stop();
                RefreshTrayIcon();
            };

            // ── backstop poll ──
            _pollTimer = new System.Windows.Forms.Timer { Interval = PollIntervalMs };
            _pollTimer.Tick += (_, _) => OnForegroundChanged(NativeMethods.GetForegroundWindow());

            Load        += OnLoad;
            FormClosing += OnFormClosing;
        }

        // ─── lifecycle ───────────────────────────────────────────────────────────

        // Registering here rather than in Load means the hook survives a handle
        // recreation, which would otherwise silently stop all notifications.
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            if (_shellMsg == 0)
                _shellMsg = NativeMethods.RegisterWindowMessage("SHELLHOOK");

            NativeMethods.RegisterShellHookWindow(Handle);
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            NativeMethods.DeregisterShellHookWindow(Handle);
            base.OnHandleDestroyed(e);
        }

        private void OnLoad(object sender, EventArgs e)
        {
            // Hide the form properly on the next pump cycle
            BeginInvoke((Action)Hide);

            // Save the current system timeout so we can restore it on exit
            _originalTimeout = NativeMethods.GetForegroundLockTimeout();
            ApplyLockTimeout();

            IntPtr foreground = NativeMethods.GetForegroundWindow();
            if (!IsOwnWindow(foreground) && NativeMethods.IsAppWindow(foreground))
                _userForeground = foreground;
            _lastForeground = foreground;

            _pollTimer.Start();
            _ready = true;
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            _ready = false;
            _pollTimer.Stop();
            _resetTimer.Stop();

            NativeMethods.DeregisterShellHookWindow(Handle);

            // Restore the original timeout
            NativeMethods.SetForegroundLockTimeout(_originalTimeout);

            _trayIcon.Visible = false;
            _trayDisposed = true;
            _trayIcon.Dispose();
            _iconActive.Dispose();
            _iconBlocked.Dispose();
            _iconPaused.Dispose();
        }

        // ─── shell hook message pump ─────────────────────────────────────────────
        protected override void WndProc(ref Message m)
        {
            if (_ready && _shellMsg != 0 && m.Msg == (int)_shellMsg)
                HandleShellNotification(unchecked((int)m.WParam.ToInt64()), m.LParam);

            base.WndProc(ref m);
        }

        private void HandleShellNotification(int code, IntPtr hwnd)
        {
            switch (code)
            {
                case NativeMethods.HSHELL_WINDOWCREATED:
                    NoteWindowCreated(hwnd);
                    break;

                case NativeMethods.HSHELL_WINDOWDESTROYED:
                    Forget(hwnd);
                    break;

                // HSHELL_RUDEAPPACTIVATED is the same notification for a full-screen
                // window, not proof of a forced activation, so it is judged by the
                // same rules instead of being blocked outright.
                case NativeMethods.HSHELL_WINDOWACTIVATED:
                case NativeMethods.HSHELL_RUDEAPPACTIVATED:
                    OnForegroundChanged(hwnd);
                    break;
            }
        }

        // ─── policy ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Single entry point for "the foreground window is now X", fed by both the
        /// shell hook and the backstop poll.
        /// </summary>
        private void OnForegroundChanged(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || hwnd == _lastForeground)
                return;

            _lastForeground = hwnd;

            if (IsOwnWindow(hwnd) || !NativeMethods.IsAppWindow(hwnd))
                return;

            if (!_config.ProtectionEnabled)
            {
                _userForeground = hwnd;
                return;
            }

            string app = NativeMethods.GetProcessName(NativeMethods.GetWindowProcessId(hwnd));

            // Whitelisted apps are always allowed through.
            if (_config.IsWhitelisted(app) || !ShouldKeepInBackground(hwnd, app))
            {
                _userForeground = hwnd;
                return;
            }

            // Only intercept when there is somewhere sensible to hand focus back to;
            // stranding the user on the desktop would be worse than the steal.
            if (hwnd == _userForeground || !IsUsableRestoreTarget(_userForeground))
            {
                _userForeground = hwnd;
                return;
            }

            KeepInBackground(hwnd, app);
        }

        private bool ShouldKeepInBackground(IntPtr hwnd, string app)
        {
            // We already tried and it would not stay put — leave it be for now.
            if (BackingOff(hwnd))
                return false;

            // Blacklisted apps never come forward, however they got activated.
            if (_config.IsBlacklisted(app))
                return true;

            // An app raising its own second window (a dialog, another document) is
            // doing what the user just asked it to.
            if (SharesProcessWithUserWindow(hwnd))
                return false;

            // The window only just appeared: a program that has finished loading.
            // That is the case this whole app exists for, so it goes to the back
            // regardless of what the user happened to be doing at that moment.
            if (JustCreated(hwnd))
                return _config.BackgroundNewWindows || !LooksUserInitiated(hwnd);

            // An existing window pushing itself forward on its own.
            return !LooksUserInitiated(hwnd);
        }

        /// <summary>
        /// Distinguishes a switch the user asked for from one an app helped itself
        /// to. Alt+Tab and Win+Tab are still held when the switch lands, and a click
        /// leaves the pointer on the window or on the taskbar; anything else is
        /// judged on how recently the user touched the machine at all.
        /// </summary>
        private bool LooksUserInitiated(IntPtr hwnd)
        {
            if (NativeMethods.IsKeyDown(NativeMethods.VK_MENU) ||
                NativeMethods.IsKeyDown(NativeMethods.VK_LWIN) ||
                NativeMethods.IsKeyDown(NativeMethods.VK_RWIN))
                return true;

            if (NativeMethods.PointerIsOn(hwnd))
                return true;

            return NativeMethods.GetIdleTime() <= (uint)_config.IdleThresholdMs;
        }

        /// <summary>
        /// Pushes the window to the bottom of the z-order and gives the user their
        /// window back. The app keeps running and finishes starting up — it just
        /// does it out of the way, with its taskbar button flashing.
        /// </summary>
        private void KeepInBackground(IntPtr hwnd, string app)
        {
            IntPtr restoreTo = _userForeground;

            NativeMethods.SendToBack(hwnd);
            NativeMethods.RestoreForeground(restoreTo);

            // Our own restore arrives as another activation — don't reprocess it.
            _lastForeground = restoreTo;

            NativeMethods.FlashTaskbar(hwnd);

            RecordInterception(hwnd);
            ShowReadyNotification(hwnd, app);
        }

        private void ShowReadyNotification(IntPtr hwnd, string app)
        {
            string name = FriendlyName(hwnd, app);

            _trayIcon.Icon = _iconBlocked;
            _trayIcon.Text = TruncateTip($"FocusShield \u2014 {name} is waiting");
            _resetTimer.Stop();
            _resetTimer.Start();

            if (!_config.ShowNotifications)
                return;

            // One balloon per window per cooldown, so a persistent app cannot spam.
            if (_seen.TryGetValue(hwnd, out Interception info))
            {
                if (info.NotifiedTick != 0 && !Elapsed(info.NotifiedTick, NotifyCooldownMs))
                    return;

                info.NotifiedTick = Environment.TickCount;
            }

            string tipText = _config.IsBlacklisted(app)
                ? $"{name} is on your blacklist \u2014 kept in the background."
                : "Finished loading in the background. Click its taskbar button when you\u2019re ready.";

            _trayIcon.ShowBalloonTip(
                timeout:  3000,
                tipTitle: $"{name} is ready",
                tipText:  tipText,
                tipIcon:  ToolTipIcon.Info);
        }

        // ─── window bookkeeping ──────────────────────────────────────────────────

        private void NoteWindowCreated(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || IsOwnWindow(hwnd))
                return;

            _created[hwnd] = Environment.TickCount;
            PruneCreated();
        }

        /// <summary>
        /// True the first time a freshly created window is activated. The entry is
        /// consumed so later activations of the same window are judged normally.
        /// </summary>
        private bool JustCreated(IntPtr hwnd)
        {
            if (!_created.TryGetValue(hwnd, out int tick))
                return false;

            _created.Remove(hwnd);
            return !Elapsed(tick, NewWindowGraceMs);
        }

        private void Forget(IntPtr hwnd)
        {
            _created.Remove(hwnd);
            _seen.Remove(hwnd);

            if (_userForeground == hwnd) _userForeground = IntPtr.Zero;
            if (_lastForeground == hwnd) _lastForeground = IntPtr.Zero;
        }

        private void RecordInterception(IntPtr hwnd)
        {
            if (!_seen.TryGetValue(hwnd, out Interception info))
            {
                info = new Interception();
                _seen[hwnd] = info;
                PruneSeen();
            }
            else if (Elapsed(info.LastTick, InterceptionResetMs))
            {
                info.Count = 0;
            }

            info.LastTick = Environment.TickCount;
            info.Count++;

            if (info.Count >= MaxInterceptions)
            {
                info.BackOffUntil = unchecked(Environment.TickCount + BackOffMs);
                info.Count = 0;
            }
        }

        private bool BackingOff(IntPtr hwnd) =>
            _seen.TryGetValue(hwnd, out Interception info) &&
            info.BackOffUntil != 0 &&
            !Elapsed(info.BackOffUntil, 0);

        private bool SharesProcessWithUserWindow(IntPtr hwnd)
        {
            if (!NativeMethods.IsWindow(_userForeground))
                return false;

            uint pid = NativeMethods.GetWindowProcessId(hwnd);
            return pid != 0 && pid == NativeMethods.GetWindowProcessId(_userForeground);
        }

        private static bool IsUsableRestoreTarget(IntPtr hwnd) =>
            hwnd != IntPtr.Zero &&
            NativeMethods.IsWindow(hwnd) &&
            NativeMethods.IsWindowVisible(hwnd) &&
            !NativeMethods.IsIconic(hwnd);

        private static bool IsOwnWindow(IntPtr hwnd) =>
            hwnd != IntPtr.Zero &&
            NativeMethods.GetWindowProcessId(hwnd) == NativeMethods.GetCurrentProcessId();

        private void PruneCreated()
        {
            if (_created.Count <= 128)
                return;

            var stale = new List<IntPtr>();
            foreach (var entry in _created)
            {
                if (Elapsed(entry.Value, NewWindowGraceMs) || !NativeMethods.IsWindow(entry.Key))
                    stale.Add(entry.Key);
            }

            foreach (IntPtr hwnd in stale)
                _created.Remove(hwnd);
        }

        private void PruneSeen()
        {
            if (_seen.Count <= 128)
                return;

            var stale = new List<IntPtr>();
            foreach (var entry in _seen)
            {
                if (!NativeMethods.IsWindow(entry.Key))
                    stale.Add(entry.Key);
            }

            foreach (IntPtr hwnd in stale)
                _seen.Remove(hwnd);
        }

        // Tick counts wrap every ~25 days; unchecked subtraction stays correct.
        private static bool Elapsed(int since, int ms) =>
            unchecked(Environment.TickCount - since) >= ms;

        /// <summary>
        /// Best available name for a window: its title, the process behind it, or a
        /// generic fallback. Window titles can be a whole sentence, so it is capped
        /// to keep the balloon and the tray tooltip readable.
        /// </summary>
        private static string FriendlyName(IntPtr hwnd, string app)
        {
            string title = NativeMethods.GetWindowTitle(hwnd).Trim();
            if (title.Length == 0)
                title = string.IsNullOrWhiteSpace(app) ? "An app" : app;

            return title.Length > 40 ? title.Substring(0, 39) + "\u2026" : title;
        }

        // ─── menu handlers ───────────────────────────────────────────────────────
        private void OnToggleEnabled(object sender, EventArgs e)
        {
            _config.ProtectionEnabled = !_config.ProtectionEnabled;
            _menuEnabled.Checked = _config.ProtectionEnabled;

            ApplyLockTimeout();
            RefreshTrayIcon();
        }

        private void OnToggleBackgroundNewWindows(object sender, EventArgs e)
        {
            _config.BackgroundNewWindows = !_config.BackgroundNewWindows;
            _menuBackground.Checked = _config.BackgroundNewWindows;
        }

        private void OnToggleNotifications(object sender, EventArgs e)
        {
            _config.ShowNotifications = !_config.ShowNotifications;
            _menuNotify.Checked = _config.ShowNotifications;
        }

        private void OnEditWhitelist(object sender, EventArgs e)
        {
            ShowListEditor(
                "Whitelist",
                "These apps may always come to the front when they finish loading.",
                _config.GetWhitelist, _config.AddToWhitelist, _config.RemoveFromWhitelist);
        }

        private void OnEditBlacklist(object sender, EventArgs e)
        {
            ShowListEditor(
                "Blacklist",
                "These apps are always kept in the background, even if you were busy.",
                _config.GetBlacklist, _config.AddToBlacklist, _config.RemoveFromBlacklist);
        }

        private void ShowListEditor(string listName, string hint,
            Func<IEnumerable<string>> read, Action<string> add, Action<string> remove)
        {
            using var form = new Form
            {
                Text            = $"FocusShield \u2014 {listName}",
                ClientSize      = new Size(380, 300),
                MinimumSize     = new Size(340, 260),
                StartPosition   = FormStartPosition.CenterScreen,
                ShowIcon        = false,
                MaximizeBox     = false
            };

            var label = new Label
            {
                Text     = hint,
                Location = new Point(12, 10),
                Size     = new Size(356, 32),
                Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var listBox = new ListBox
            {
                Location       = new Point(12, 46),
                Size           = new Size(356, 200),
                IntegralHeight = false,
                Anchor         = AnchorStyles.Top | AnchorStyles.Bottom |
                                 AnchorStyles.Left | AnchorStyles.Right
            };

            var btnAdd = new Button
            {
                Text     = "Add\u2026",
                Location = new Point(12, 258),
                Size     = new Size(84, 28),
                Anchor   = AnchorStyles.Bottom | AnchorStyles.Left
            };

            var btnRemove = new Button
            {
                Text     = "Remove",
                Location = new Point(104, 258),
                Size     = new Size(84, 28),
                Anchor   = AnchorStyles.Bottom | AnchorStyles.Left,
                Enabled  = false
            };

            var btnClose = new Button
            {
                Text         = "Close",
                Location     = new Point(284, 258),
                Size         = new Size(84, 28),
                Anchor       = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };

            void Reload()
            {
                listBox.BeginUpdate();
                listBox.Items.Clear();
                foreach (string app in read())
                    listBox.Items.Add(app);
                listBox.EndUpdate();

                btnRemove.Enabled = listBox.SelectedIndex >= 0;
            }

            // Neither Add nor Remove carries a DialogResult — they edit the list in
            // place instead of closing the editor.
            btnAdd.Click += (_, _) =>
            {
                string app = PromptForAppName(form);
                if (string.IsNullOrEmpty(app))
                    return;

                add(app);
                Reload();
                listBox.SelectedItem = app;
            };

            btnRemove.Click += (_, _) =>
            {
                if (listBox.SelectedItem is string selected)
                {
                    remove(selected);
                    Reload();
                }
            };

            listBox.SelectedIndexChanged += (_, _) => btnRemove.Enabled = listBox.SelectedIndex >= 0;

            form.Controls.AddRange(new Control[] { label, listBox, btnAdd, btnRemove, btnClose });
            form.CancelButton = btnClose;

            Reload();
            form.ShowDialog();
        }

        private static string PromptForAppName(Form owner)
        {
            using var form = new Form
            {
                Text            = "Add Application",
                ClientSize      = new Size(300, 122),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterParent,
                ShowIcon        = false,
                MinimizeBox     = false,
                MaximizeBox     = false
            };

            var label = new Label
            {
                Text     = "Process name, without .exe (Task Manager \u2192 Details):",
                Location = new Point(12, 12),
                Size     = new Size(276, 32)
            };

            var textBox = new TextBox
            {
                Location = new Point(12, 46),
                Size     = new Size(276, 23)
            };

            var btnOk = new Button
            {
                Text         = "OK",
                Location     = new Point(116, 84),
                Size         = new Size(84, 26),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text         = "Cancel",
                Location     = new Point(204, 84),
                Size         = new Size(84, 26),
                DialogResult = DialogResult.Cancel
            };

            form.Controls.AddRange(new Control[] { label, textBox, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            return form.ShowDialog(owner) == DialogResult.OK
                ? ConfigManager.Normalize(textBox.Text)
                : null;
        }

        private void OnToggleBootload(object sender, EventArgs e)
        {
            SetBootload(!IsBootloadEnabled());
            _menuBootload.Checked = IsBootloadEnabled();
        }

        private void OnExit(object sender, EventArgs e) => Application.Exit();

        // ─── helpers ─────────────────────────────────────────────────────────────

        /// <summary>
        /// Path of the running executable. Assembly.Location points at the managed
        /// dll (and is empty for single-file builds), which Windows cannot start.
        /// </summary>
        private static string ExecutablePath()
        {
            string path = Environment.ProcessPath;
            return string.IsNullOrEmpty(path) ? Application.ExecutablePath : path;
        }

        private static bool IsBootloadEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
                return key?.GetValue(RunValueName) != null;
            }
            catch
            {
                return false;
            }
        }

        private static void SetBootload(bool enable)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
                if (key == null)
                    return;

                if (enable)
                    key.SetValue(RunValueName, $"\"{ExecutablePath()}\"");   // quoted: the path may contain spaces
                else
                    key.DeleteValue(RunValueName, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to update the startup entry: {ex.Message}",
                    "FocusShield", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyLockTimeout() =>
            NativeMethods.SetForegroundLockTimeout(
                _config.ProtectionEnabled ? LockTimeoutMs : _originalTimeout);

        private void RefreshTrayIcon()
        {
            _trayIcon.Icon = _config.ProtectionEnabled ? _iconActive : _iconPaused;
            _trayIcon.Text = TrayTooltip();
        }

        private string TrayTooltip() =>
            _config.ProtectionEnabled
                ? "FocusShield \u2014 protecting focus"
                : "FocusShield \u2014 paused";

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
