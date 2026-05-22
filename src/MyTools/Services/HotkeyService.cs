using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MyTools.Services
{
    public static class HotkeyService
    {
        public const int ScreenshotHotkeyId = 9001;
        public const int VideoRecordHotkeyId = 9002;
        public const int AudioRecordHotkeyId = 9003;
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
        private static readonly System.Collections.Generic.HashSet<int> _registeredIds = new System.Collections.Generic.HashSet<int>();
        private static int _lastWin32Error;

        public static void Initialize(IntPtr windowHandle)
        {
            lock (SyncRoot)
            {
                if (_handle != IntPtr.Zero && _handle != windowHandle && _registeredIds.Count > 0)
                {
                    foreach (var id in _registeredIds)
                        UnregisterHotKey(_handle, id);
                    _registeredIds.Clear();
                }

                _handle = windowHandle;
            }
        }

        public static bool Register(uint modifiers, uint key)
            => Register(ScreenshotHotkeyId, modifiers, key);

        public static bool Register(int id, uint modifiers, uint key)
        {
            lock (SyncRoot)
            {
                if (_handle == IntPtr.Zero || key == 0)
                {
                    _lastWin32Error = 0;
                    return false;
                }

                if (_registeredIds.Contains(id))
                {
                    UnregisterHotKey(_handle, id);
                    _registeredIds.Remove(id);
                }

                var registered = RegisterHotKey(_handle, id, modifiers | ModNoRepeat, key);
                if (!registered)
                {
                    var errorCode = Marshal.GetLastWin32Error();
                    if (errorCode == ErrorInvalidParameter)
                    {
                        // Win7+ supports MOD_NOREPEAT, but keep a fallback for older/quirky environments.
                        registered = RegisterHotKey(_handle, id, modifiers, key);
                    }
                }

                _lastWin32Error = registered ? 0 : Marshal.GetLastWin32Error();
                if (registered)
                    _registeredIds.Add(id);
                return registered;
            }
        }

        public static void Unregister() => UnregisterById(ScreenshotHotkeyId);

        public static void UnregisterById(int id)
        {
            lock (SyncRoot)
            {
                if (_handle == IntPtr.Zero) return;
                UnregisterHotKey(_handle, id);
                _registeredIds.Remove(id);
            }
        }

        public static void UnregisterAll()
        {
            lock (SyncRoot)
            {
                if (_handle == IntPtr.Zero) return;
                foreach (var id in _registeredIds)
                    UnregisterHotKey(_handle, id);
                _registeredIds.Clear();
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
