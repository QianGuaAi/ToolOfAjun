using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;

namespace MyTools.Services
{
    /// <summary>
    /// Windows 资源管理器/任务栏/桌面图标若干常用调整。
    /// 只写 HKCU，不需管理员权限。时钟秒/任务栏/托盘改动需重启 explorer.exe 生效；
    /// 桌面图标通过 SHChangeNotify 即时生效。
    /// </summary>
    public static class WindowsTweaksService
    {
        // ===== 注册表路径 =====
        private const string KeyAdvanced = @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced";
        private const string KeyExplorer = @"Software\Microsoft\Windows\CurrentVersion\Explorer";
        private const string KeyHideNew = @"Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel";
        private const string KeyHideClassic = @"Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu";

        // ===== 桌面图标 CLSID =====
        public const string ClsidComputer = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
        public const string ClsidRecycleBin = "{645FF040-5081-101B-9F08-00AA002F954E}";
        public const string ClsidControlPanel = "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}";
        public const string ClsidUserFiles = "{59031a47-3f72-44a7-89c5-5595fe6b30ee}";
        public const string ClsidNetwork = "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}";

        public enum TaskbarGlom
        {
            AlwaysCombine = 0,        // 总是合并并隐藏标签
            CombineWhenFull = 1,      // 任务栏满时才合并
            NeverCombine = 2          // 从不合并
        }

        // ===== 系统版本 =====
        /// <summary>Win11 = Build &gt;= 22000；时钟秒功能需 22H2(22621)+</summary>
        public static bool IsWindows11 => Environment.OSVersion.Version.Major >= 10
            && Environment.OSVersion.Version.Build >= 22000;

        // ===================== 时钟秒 =====================
        public static bool GetShowSecondsInClock()
        {
            return ReadDword(Registry.CurrentUser, KeyAdvanced, "ShowSecondsInSystemClock", 0) != 0;
        }
        public static void SetShowSecondsInClock(bool show)
        {
            WriteDword(Registry.CurrentUser, KeyAdvanced, "ShowSecondsInSystemClock", show ? 1 : 0);
        }

        // ===================== 托盘图标自动隐藏 =====================
        public static bool GetTrayShowAll()
        {
            // EnableAutoTray=0 → 全部显示；=1（默认）→ 隐藏不活跃
            return ReadDword(Registry.CurrentUser, KeyExplorer, "EnableAutoTray", 1) == 0;
        }
        public static void SetTrayShowAll(bool showAll)
        {
            WriteDword(Registry.CurrentUser, KeyExplorer, "EnableAutoTray", showAll ? 0 : 1);
        }

        // ===================== 任务栏合并 =====================
        public static TaskbarGlom GetTaskbarGlom()
        {
            var v = ReadDword(Registry.CurrentUser, KeyAdvanced, "TaskbarGlomLevel", 0);
            if (v < 0 || v > 2) v = 0;
            return (TaskbarGlom)v;
        }
        public static void SetTaskbarGlom(TaskbarGlom level)
        {
            WriteDword(Registry.CurrentUser, KeyAdvanced, "TaskbarGlomLevel", (int)level);
            // Win11/10 部分版本同时认 MMTaskbarGlomLevel（多显示器任务栏）
            WriteDword(Registry.CurrentUser, KeyAdvanced, "MMTaskbarGlomLevel", (int)level);
        }

        // ===================== 桌面图标 =====================
        public static bool GetDesktopIconVisible(string clsid)
        {
            // 0 = 显示, 1 = 隐藏；缺省按显示处理（视图标而异，这里统一缺省=显示）
            return ReadDword(Registry.CurrentUser, KeyHideNew, clsid, 0) == 0;
        }

        public static void SetDesktopIconVisible(string clsid, bool visible)
        {
            var v = visible ? 0 : 1;
            WriteDword(Registry.CurrentUser, KeyHideNew, clsid, v);
            WriteDword(Registry.CurrentUser, KeyHideClassic, clsid, v);
            // 即时刷新桌面图标
            try { SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero); }
            catch { /* swallow */ }
        }

        // ===================== 重启 Explorer =====================
        /// <summary>结束所有 explorer.exe 进程并重新启动一个；用于让任务栏/时钟/托盘改动生效。</summary>
        public static void RestartExplorer()
        {
            try
            {
                foreach (var p in Process.GetProcessesByName("explorer"))
                {
                    try { p.Kill(); } catch { /* 单个失败继续 */ }
                }
            }
            catch { /* swallow */ }

            // 等系统短暂稳定（部分 Windows 版本 explorer 会自动复活）
            Thread.Sleep(800);

            try
            {
                if (Process.GetProcessesByName("explorer").Length == 0)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        UseShellExecute = false
                    });
                }
            }
            catch { /* swallow */ }
        }

        // ===================== 私有：注册表读写 =====================
        private static int ReadDword(RegistryKey root, string subKey, string name, int defVal)
        {
            try
            {
                using (var key = root.OpenSubKey(subKey))
                {
                    if (key == null) return defVal;
                    var v = key.GetValue(name);
                    if (v is int i) return i;
                    if (v != null && int.TryParse(v.ToString(), out var p)) return p;
                    return defVal;
                }
            }
            catch { return defVal; }
        }

        private static void WriteDword(RegistryKey root, string subKey, string name, int value)
        {
            using (var key = root.CreateSubKey(subKey, writable: true))
            {
                if (key == null) throw new InvalidOperationException("打开注册表项失败：" + subKey);
                key.SetValue(name, value, RegistryValueKind.DWord);
            }
        }

        // ===================== Win32 互操作 =====================
        private const uint SHCNE_ASSOCCHANGED = 0x08000000;
        private const uint SHCNF_IDLIST = 0x0000;

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
    }
}
