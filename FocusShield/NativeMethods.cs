using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FocusShield
{
    internal static class NativeMethods
    {
        // SystemParametersInfo actions
        public const int SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000;
        public const int SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001;
        public const uint SPIF_SENDCHANGE = 0x0002;

        // Shell hook codes delivered via WM_SHELLHOOKMESSAGE
        public const int HSHELL_WINDOWCREATED    = 1;
        public const int HSHELL_WINDOWDESTROYED  = 2;
        public const int HSHELL_WINDOWACTIVATED  = 4;     // normal activation
        public const int HSHELL_RUDEAPPACTIVATED = 0x8004; // forced activation

        // FlashWindowEx flags
        public const uint FLASHW_STOP      = 0;
        public const uint FLASHW_CAPTION   = 1;
        public const uint FLASHW_TRAY      = 2;
        public const uint FLASHW_ALL       = 3;
        public const uint FLASHW_TIMER     = 4;
        public const uint FLASHW_TIMERNOFG = 12; // flash until window comes to foreground

        [StructLayout(LayoutKind.Sequential)]
        public struct FLASHWINFO
        {
            public uint   cbSize;
            public IntPtr hwnd;
            public uint   dwFlags;
            public uint   uCount;
            public uint   dwTimeout;
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
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentProcessId();

        // ── flash / window text ─────────────────────────────────────────
        [DllImport("user32.dll")]
        public static extern bool FlashWindowEx(ref FLASHWINFO pfwi);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool DestroyIcon(IntPtr hIcon);

        // ── helpers ─────────────────────────────────────────────────────

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

        /// <summary>
        /// Try to bring <paramref name="target"/> to the foreground even
        /// if we are not the foreground process, by temporarily merging
        /// our input queue with the current foreground thread's.
        /// </summary>
        public static bool ForceForeground(IntPtr target)
        {
            if (target == IntPtr.Zero || !IsWindow(target))
                return false;

            IntPtr fgWnd  = GetForegroundWindow();
            uint   fgTid  = GetWindowThreadProcessId(fgWnd, out _);
            uint   ourTid = GetCurrentThreadId();

            bool attached = false;
            if (fgTid != ourTid)
                attached = AttachThreadInput(ourTid, fgTid, true);

            bool ok = SetForegroundWindow(target);

            if (attached)
                AttachThreadInput(ourTid, fgTid, false);

            return ok;
        }

        public static string GetWindowTitle(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
    }
}
