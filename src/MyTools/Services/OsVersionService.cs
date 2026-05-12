using System;
using System.Runtime.InteropServices;

namespace MyTools.Services
{
    /// <summary>
    /// 通过 RtlGetVersion 获取真实的 Windows 版本号，规避 GetVersionEx 在未声明 manifest 时被锁定到 6.2 的问题。
    /// </summary>
    public static class OsVersionService
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RTL_OSVERSIONINFOEX
        {
            public uint dwOSVersionInfoSize;
            public uint dwMajorVersion;
            public uint dwMinorVersion;
            public uint dwBuildNumber;
            public uint dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szCSDVersion;
            public ushort wServicePackMajor;
            public ushort wServicePackMinor;
            public ushort wSuiteMask;
            public byte wProductType;
            public byte wReserved;
        }

        [DllImport("ntdll.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int RtlGetVersion(ref RTL_OSVERSIONINFOEX versionInfo);

        private static readonly Lazy<RTL_OSVERSIONINFOEX> _info = new Lazy<RTL_OSVERSIONINFOEX>(QueryRtlVersion);

        private static RTL_OSVERSIONINFOEX QueryRtlVersion()
        {
            var info = new RTL_OSVERSIONINFOEX
            {
                dwOSVersionInfoSize = (uint)Marshal.SizeOf(typeof(RTL_OSVERSIONINFOEX))
            };
            try
            {
                RtlGetVersion(ref info);
            }
            catch
            {
                // ignore, fields will be 0
            }
            return info;
        }

        public static int MajorVersion => (int)_info.Value.dwMajorVersion;
        public static int MinorVersion => (int)_info.Value.dwMinorVersion;
        public static int BuildNumber => (int)_info.Value.dwBuildNumber;
        public static int ServicePackMajor => _info.Value.wServicePackMajor;

        /// <summary>Windows 7 = 6.1，含 Win7 SP1。</summary>
        public static bool IsWindows7 => MajorVersion == 6 && MinorVersion == 1;

        /// <summary>Windows 8 / 8.1 = 6.2 / 6.3。</summary>
        public static bool IsWindows8OrLater => (MajorVersion == 6 && MinorVersion >= 2) || MajorVersion >= 10;

        /// <summary>Windows 10 / 11 真实主版本号为 10。</summary>
        public static bool IsWindows10OrGreater => MajorVersion >= 10;

        /// <summary>Windows 11 起始构建号 22000。</summary>
        public static bool IsWindows11OrGreater => IsWindows10OrGreater && BuildNumber >= 22000;

        /// <summary>用户可读的版本字符串，例如 "Windows 10 22H2 (10.0.19045)"。</summary>
        public static string DisplayName
        {
            get
            {
                var name = "Windows";
                if (IsWindows11OrGreater) name = "Windows 11";
                else if (IsWindows10OrGreater) name = "Windows 10";
                else if (MajorVersion == 6 && MinorVersion == 3) name = "Windows 8.1";
                else if (MajorVersion == 6 && MinorVersion == 2) name = "Windows 8";
                else if (IsWindows7) name = "Windows 7";
                else if (MajorVersion == 6 && MinorVersion == 0) name = "Windows Vista";
                else if (MajorVersion == 5) name = "Windows XP/2003";

                var sp = ServicePackMajor > 0 ? $" SP{ServicePackMajor}" : string.Empty;
                return $"{name}{sp} ({MajorVersion}.{MinorVersion}.{BuildNumber})";
            }
        }
    }
}
