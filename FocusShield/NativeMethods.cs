using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace FocusShield
{
    internal static class NativeMethods
    {
        // ─── SystemParametersInfo ───────────────────────────────────────
        public const int  SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000;
        public const int  SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001;
        public const uint SPIF_SENDCHANGE = 0x0002;

        // ─── shell hook codes delivered via WM_SHELLHOOKMESSAGE ─────────
        public const int HSHELL_WINDOWCREATED   = 1;
        public const int HSHELL_WINDOWDESTROYED = 2;
        public const int HSHELL_WINDOWACTIVATED = 4;
        public const int HSHELL_REDRAW          = 6;
        public const int HSHELL_HIGHBIT         = 0x8000;

        // Same notification as HSHELL_WINDOWACTIVATED — the high bit only means
        // the activated window is running full-screen, NOT that it forced itself
        // to the front, so it gets judged by exactly the same rules.
        public const int HSHELL_RUDEAPPACTIVATED = HSHELL_WINDOWACTIVATED | HSHELL_HIGHBIT;

        // An app politely asking for attention (taskbar flash) — never a steal.
        public const int HSHELL_FLASH = HSHELL_REDRAW | HSHELL_HIGHBIT;

        // ─── window styles ──────────────────────────────────────────────
        public const int GWL_STYLE   = -16;
        public const int GWL_EXSTYLE = -20;

        public const long WS_CHILD         = 0x40000000;
        public const long WS_EX_TOOLWINDOW = 0x00000080;
        public const long WS_EX_APPWINDOW  = 0x00040000;
        public const long WS_EX_NOACTIVATE = 0x08000000;

        // ─── SetWindowPos ───────────────────────────────────────────────
        public static readonly IntPtr HWND_TOP    = new IntPtr(0);
        public static readonly IntPtr HWND_BOTTOM = new IntPtr(1);

        public const uint SWP_NOSIZE         = 0x0001;
        public const uint SWP_NOMOVE         = 0x0002;
        public const uint SWP_NOACTIVATE     = 0x0010;
        public const uint SWP_NOOWNERZORDER  = 0x0200;
        public const uint SWP_ASYNCWINDOWPOS = 0x4000;

        // ─── misc ───────────────────────────────────────────────────────
        public const uint GA_ROOT = 2;

        public const uint WM_NULL          = 0x0000;
        public const uint SMTO_ABORTIFHUNG = 0x0002;

        public const int DWMWA_CLOAKED = 14;

        public const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

        // Virtual keys used to recognise a deliberate window switch
        public const int VK_MENU = 0x12;   // Alt — held during Alt+Tab
        public const int VK_LWIN = 0x5B;
        public const int VK_RWIN = 0x5C;

        // FlashWindowEx flags
        public const uint FLASHW_TRAY      = 2;
        public const uint FLASHW_TIMERNOFG = 12; // flash until the window comes to the foreground

        [StructLayout(LayoutKind.Sequential)]
        public struct FLASHWINFO
        {
            public uint   cbSize;
            public IntPtr hwnd;
            public uint   dwFlags;
            public uint   uCount;
            public uint   dwTimeout;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X, Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        // ── shell hook ──────────────────────────────────────────────────
        [DllImport("user32.dll")]
        public static extern bool RegisterShellHookWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool DeregisterShellHookWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern uint RegisterWindowMessage(string lpString);

        // ── SPI — GET (pvParam is a pointer to DWORD) ───────────────────
        [DllImport("user32.dll")]
        public static extern bool SystemParametersInfo(
            int uiAction, uint uiParam, ref uint pvParam, uint fWinIni);

        // ── SPI — SET (pvParam carries the value as a pointer) ──────────
        [DllImport("user32.dll")]
        public static extern bool SystemParametersInfo(
            int uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

        // ── foreground / focus ──────────────────────────────────────────
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        // ── window inspection ───────────────────────────────────────────
        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll")]
        public static extern IntPtr GetShellWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg,
            IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute,
            out int pvAttribute, int cbAttribute);

        // ── input state ─────────────────────────────────────────────────
        [DllImport("user32.dll")]
        public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT Point);

        // ── flash ───────────────────────────────────────────────────────
        [DllImport("user32.dll")]
        public static extern bool FlashWindowEx(ref FLASHWINFO pfwi);

        // ── process ─────────────────────────────────────────────────────
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentProcessId();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool QueryFullProcessImageName(IntPtr hProcess, uint dwFlags,
            StringBuilder lpExeName, ref uint lpdwSize);

        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr hObject);

        // ─── helpers ────────────────────────────────────────────────────

        private static long GetWindowStyles(IntPtr hwnd, int index) =>
            IntPtr.Size == 8 ? GetWindowLongPtr64(hwnd, index).ToInt64() : GetWindowLong32(hwnd, index);

        /// <summary>
        /// Milliseconds since the user last pressed a key or moved the mouse.
        /// </summary>
        public static uint GetIdleTime()
        {
            var lii = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO)) };
            if (!GetLastInputInfo(ref lii))
                return 0;
            return unchecked((uint)Environment.TickCount - lii.dwTime);
        }

        public static bool IsKeyDown(int vKey) => (GetAsyncKeyState(vKey) & 0x8000) != 0;

        public static uint GetWindowProcessId(IntPtr hwnd)
        {
            GetWindowThreadProcessId(hwnd, out uint pid);
            return pid;
        }

        public static string GetWindowTitle(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            return GetWindowText(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;
        }

        public static string GetWindowClass(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            return GetClassName(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;
        }

        /// <summary>
        /// Executable name (without extension) that owns <paramref name="pid"/>.
        /// Uses PROCESS_QUERY_LIMITED_INFORMATION so it also works for elevated
        /// processes, where Process.GetProcessById would throw.
        /// </summary>
        public static string GetProcessName(uint pid)
        {
            if (pid == 0) return string.Empty;

            IntPtr handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
            if (handle == IntPtr.Zero) return string.Empty;

            try
            {
                var sb = new StringBuilder(1024);
                uint size = (uint)sb.Capacity;
                return QueryFullProcessImageName(handle, 0, sb, ref size)
                    ? Path.GetFileNameWithoutExtension(sb.ToString())
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                CloseHandle(handle);
            }
        }

        /// <summary>
        /// True when the window's thread answers a message within 200 ms. A window
        /// that is still loading often does not, and touching its input queue would
        /// drag FocusShield down with it.
        /// </summary>
        public static bool IsResponsive(IntPtr hwnd) =>
            SendMessageTimeout(hwnd, WM_NULL, IntPtr.Zero, IntPtr.Zero,
                SMTO_ABORTIFHUNG, 200, out _) != IntPtr.Zero;

        /// <summary>
        /// A window that has been cloaked by the desktop window manager is not on
        /// screen at all (suspended UWP apps leave plenty of these lying around).
        /// </summary>
        public static bool IsCloaked(IntPtr hwnd)
        {
            try
            {
                return DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, out int cloaked, sizeof(int)) == 0
                       && cloaked != 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// True for real, top-level application windows — the only kind that can
        /// take over the screen. Filters out the desktop, the shell, tool windows,
        /// child windows, cloaked windows and zero-sized helper windows.
        /// </summary>
        public static bool IsAppWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd) || !IsWindowVisible(hwnd))
                return false;

            if (hwnd == GetShellWindow() || hwnd == GetDesktopWindow())
                return false;

            if (GetAncestor(hwnd, GA_ROOT) != hwnd)
                return false;

            long style   = GetWindowStyles(hwnd, GWL_STYLE);
            long exStyle = GetWindowStyles(hwnd, GWL_EXSTYLE);

            if ((style & WS_CHILD) != 0)
                return false;

            if ((exStyle & WS_EX_NOACTIVATE) != 0)
                return false;

            if ((exStyle & WS_EX_TOOLWINDOW) != 0 && (exStyle & WS_EX_APPWINDOW) == 0)
                return false;

            if (IsCloaked(hwnd))
                return false;

            if (!GetWindowRect(hwnd, out RECT r) || r.Right - r.Left <= 1 || r.Bottom - r.Top <= 1)
                return false;

            switch (GetWindowClass(hwnd))
            {
                case "Progman":                     // desktop
                case "WorkerW":                     // desktop wallpaper host
                case "Shell_TrayWnd":               // taskbar
                case "Shell_SecondaryTrayWnd":
                case "Windows.UI.Core.CoreWindow":  // Start menu, search, action centre
                    return false;
                default:
                    return true;
            }
        }

        /// <summary>
        /// True when the mouse pointer is over <paramref name="hwnd"/> or over the
        /// taskbar — the two places a deliberate window switch is clicked from.
        /// </summary>
        public static bool PointerIsOn(IntPtr hwnd)
        {
            if (!GetCursorPos(out POINT pt))
                return false;

            IntPtr under = WindowFromPoint(pt);
            if (under == IntPtr.Zero)
                return false;

            IntPtr root = GetAncestor(under, GA_ROOT);
            if (root == hwnd)
                return true;

            string cls = GetWindowClass(root);
            return cls == "Shell_TrayWnd" || cls == "Shell_SecondaryTrayWnd";
        }

        public static uint GetForegroundLockTimeout()
        {
            uint value = 0;
            SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, ref value, 0);
            return value;
        }

        /// <summary>
        /// Changes the timeout in memory only — deliberately without
        /// SPIF_UPDATEINIFILE, so a crash can never leave the user's machine with
        /// FocusShield's setting baked into the registry.
        /// </summary>
        public static void SetForegroundLockTimeout(uint ms) =>
            SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, (IntPtr)ms, SPIF_SENDCHANGE);

        /// <summary>
        /// Drops <paramref name="hwnd"/> to the bottom of the z-order without
        /// activating it, so a program that has just finished loading keeps running
        /// but stops covering whatever the user is working on.
        /// SWP_ASYNCWINDOWPOS keeps us from blocking on a window that is still busy
        /// starting up and not yet pumping messages.
        /// </summary>
        public static bool SendToBack(IntPtr hwnd) =>
            SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_ASYNCWINDOWPOS);

        /// <summary>
        /// Puts <paramref name="target"/> back in front. The raised
        /// FOREGROUNDLOCKTIMEOUT that keeps other apps out would block this call
        /// too, so it is lifted for the duration and restored afterwards.
        /// </summary>
        public static bool RestoreForeground(IntPtr target)
        {
            if (target == IntPtr.Zero || !IsWindow(target))
                return false;

            uint saved = GetForegroundLockTimeout();
            SetForegroundLockTimeout(0);

            try
            {
                IntPtr fgWnd  = GetForegroundWindow();
                uint   fgTid  = fgWnd == IntPtr.Zero ? 0 : GetWindowThreadProcessId(fgWnd, out _);
                uint   ourTid = GetCurrentThreadId();

                bool attached = false;
                if (fgTid != 0 && fgTid != ourTid && IsResponsive(fgWnd))
                    attached = AttachThreadInput(ourTid, fgTid, true);

                SetWindowPos(target, HWND_TOP, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
                BringWindowToTop(target);
                bool ok = SetForegroundWindow(target);

                if (attached)
                    AttachThreadInput(ourTid, fgTid, false);

                return ok;
            }
            finally
            {
                SetForegroundLockTimeout(saved);
            }
        }

        /// <summary>
        /// Flash the taskbar button of <paramref name="hwnd"/> until the
        /// user brings it to the foreground.
        /// </summary>
        public static void FlashTaskbar(IntPtr hwnd)
        {
            var fi = new FLASHWINFO
            {
                cbSize    = (uint)Marshal.SizeOf(typeof(FLASHWINFO)),
                hwnd      = hwnd,
                dwFlags   = FLASHW_TRAY | FLASHW_TIMERNOFG,
                uCount    = 5,
                dwTimeout = 0
            };
            FlashWindowEx(ref fi);
        }
    }
}
