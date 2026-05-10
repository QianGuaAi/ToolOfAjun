using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MyTools.Services
{
    public static class HotkeyService
    {
        public const int ScreenshotHotkeyId = 9001;
        private const uint ModAlt = 0x0001;
        private const uint ModControl = 0x0002;
        private const uint ModShift = 0x0004;
        private const uint ModWin = 0x0008;
        private const uint ModNoRepeat = 0x4000;
        private const int ErrorInvalidParameter = 87;

        private static readonly object SyncRoot = new object();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private static IntPtr _handle = IntPtr.Zero;
        private static uint _currentModifiers;
        private static uint _currentKey;
        private static bool _isRegistered;
        private static int _lastWin32Error;

        public static void Initialize(IntPtr windowHandle)
        {
            lock (SyncRoot)
            {
                if (_handle != IntPtr.Zero && _handle != windowHandle && _isRegistered)
                {
                    UnregisterHotKey(_handle, ScreenshotHotkeyId);
                    _isRegistered = false;
                }

                _handle = windowHandle;
            }
        }

        public static bool Register(uint modifiers, uint key)
        {
            lock (SyncRoot)
            {
                if (_handle == IntPtr.Zero || key == 0)
                {
                    _lastWin32Error = 0;
                    return false;
                }

                if (_isRegistered)
                {
                    UnregisterHotKey(_handle, ScreenshotHotkeyId);
                    _isRegistered = false;
                }

                _currentModifiers = modifiers;
                _currentKey = key;

                var registered = RegisterHotKey(_handle, ScreenshotHotkeyId, modifiers | ModNoRepeat, key);
                if (!registered)
                {
                    var errorCode = Marshal.GetLastWin32Error();
                    if (errorCode == ErrorInvalidParameter)
                    {
                        // Win7+ supports MOD_NOREPEAT, but keep a fallback for older/quirky environments.
                        registered = RegisterHotKey(_handle, ScreenshotHotkeyId, modifiers, key);
                    }
                }

                _lastWin32Error = registered ? 0 : Marshal.GetLastWin32Error();
                _isRegistered = registered;
                return registered;
            }
        }

        public static void Unregister()
        {
            lock (SyncRoot)
            {
                if (_handle == IntPtr.Zero)
                {
                    return;
                }

                UnregisterHotKey(_handle, ScreenshotHotkeyId);
                _isRegistered = false;
            }
        }

        public static int LastWin32ErrorCode
        {
            get
            {
                lock (SyncRoot)
                {
                    return _lastWin32Error;
                }
            }
        }

        public static string BuildDisplayText(uint modifiers, uint key)
        {
            var parts = new System.Collections.Generic.List<string>();
            if ((modifiers & ModControl) != 0) parts.Add("Ctrl");
            if ((modifiers & ModShift) != 0) parts.Add("Shift");
            if ((modifiers & ModAlt) != 0) parts.Add("Alt");
            if ((modifiers & ModWin) != 0) parts.Add("Win");
            var keyName = ((Keys)key).ToString();
            parts.Add(keyName);
            return string.Join("+", parts);
        }
    }
}
