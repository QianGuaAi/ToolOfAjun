using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MyTools.Services
{
    public static class HotkeyService
    {
        public const int ScreenshotHotkeyId = 9001;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private static IntPtr _handle = IntPtr.Zero;
        private static uint _currentModifiers;
        private static uint _currentKey;

        public static void Initialize(IntPtr windowHandle)
        {
            _handle = windowHandle;
        }

        public static bool Register(uint modifiers, uint key)
        {
            if (_handle == IntPtr.Zero) return false;

            UnregisterHotKey(_handle, ScreenshotHotkeyId);
            _currentModifiers = modifiers;
            _currentKey = key;

            const uint MOD_NOREPEAT = 0x4000;
            return RegisterHotKey(_handle, ScreenshotHotkeyId, modifiers | MOD_NOREPEAT, key);
        }

        public static void Unregister()
        {
            if (_handle != IntPtr.Zero)
                UnregisterHotKey(_handle, ScreenshotHotkeyId);
        }

        public static string BuildDisplayText(uint modifiers, uint key)
        {
            var parts = new System.Collections.Generic.List<string>();
            if ((modifiers & 0x0002) != 0) parts.Add("Ctrl");
            if ((modifiers & 0x0004) != 0) parts.Add("Shift");
            if ((modifiers & 0x0001) != 0) parts.Add("Alt");
            if ((modifiers & 0x0008) != 0) parts.Add("Win");
            var keyName = ((Keys)key).ToString();
            parts.Add(keyName);
            return string.Join("+", parts);
        }
    }
}
