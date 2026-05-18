using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace MyTools.Services
{
    /// <summary>
    /// 桌面背景管理：
    ///   - 图库目录：MyTools.exe 同目录下的 Wallpapers/
    ///   - 设置壁纸：Win7/8/10/11 通用 SystemParametersInfo + 注册表 WallpaperStyle/TileWallpaper
    ///   - 读取当前壁纸：SPI_GETDESKWALLPAPER（fallback：注册表 / TranscodedWallpaper 缓存）
    /// </summary>
    public static class WallpaperService
    {
        public enum WallpaperStyle
        {
            Fill = 10,    // 填充（裁剪铺满，Win7+）
            Fit = 6,      // 适应（保持宽高比，Win7+）
            Stretch = 2,  // 拉伸（变形）
            Tile = 0,     // 平铺（依赖 TileWallpaper=1）
            Center = 1    // 居中（实际 WallpaperStyle=0 / TileWallpaper=0；用 1 作为内部 enum 标识，写注册表前转换）
        }

        public static readonly string[] SupportedExtensions = new[] { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };

        public static string LibraryDirectory =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Wallpapers");

        public static void EnsureLibrary()
        {
            Directory.CreateDirectory(LibraryDirectory);
        }

        /// <summary>列出图库内所有支持的图片，按修改时间倒序。</summary>
        public static IReadOnlyList<string> ListLibrary()
        {
            EnsureLibrary();
            return Directory.EnumerateFiles(LibraryDirectory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(p => SupportedExtensions.Contains(Path.GetExtension(p)?.ToLowerInvariant()))
                .OrderByDescending(p =>
                {
                    try { return File.GetLastWriteTime(p); }
                    catch { return DateTime.MinValue; }
                })
                .ToList();
        }

        /// <summary>把外部图片拷贝进图库；返回新路径。重名自动加序号。</summary>
        public static async Task<List<string>> ImportImagesAsync(IEnumerable<string> sourcePaths)
        {
            EnsureLibrary();
            var imported = new List<string>();
            foreach (var src in sourcePaths ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(src) || !File.Exists(src)) continue;
                var ext = Path.GetExtension(src)?.ToLowerInvariant();
                if (!SupportedExtensions.Contains(ext)) continue;

                var dest = ResolveDestination(LibraryDirectory, Path.GetFileName(src));
                using (var inS = new FileStream(src, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
                using (var outS = new FileStream(dest, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    await inS.CopyToAsync(outS).ConfigureAwait(false);
                }
                imported.Add(dest);
            }
            return imported;
        }

        /// <summary>从图库删除图片（仅图库内的文件，路径越界拒绝）。</summary>
        public static void DeleteFromLibrary(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            var fullLib = Path.GetFullPath(LibraryDirectory).TrimEnd(Path.DirectorySeparatorChar);
            var fullFile = Path.GetFullPath(path);
            if (!fullFile.StartsWith(fullLib + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("仅允许删除图库目录内的文件。");
            if (File.Exists(fullFile)) File.Delete(fullFile);
        }

        /// <summary>读取当前桌面壁纸路径（多重 fallback）。</summary>
        public static string GetCurrentWallpaperPath()
        {
            // 1) SPI_GETDESKWALLPAPER（最权威）
            try
            {
                var sb = new StringBuilder(MAX_PATH);
                if (SystemParametersInfo(SPI_GETDESKWALLPAPER, (uint)sb.Capacity, sb, 0))
                {
                    var p = sb.ToString();
                    if (!string.IsNullOrEmpty(p) && File.Exists(p)) return p;
                }
            }
            catch { /* swallow */ }

            // 2) 注册表 HKCU\Control Panel\Desktop\Wallpaper
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop"))
                {
                    if (key != null)
                    {
                        var p = key.GetValue("Wallpaper") as string;
                        if (!string.IsNullOrEmpty(p) && File.Exists(p)) return p;
                    }
                }
            }
            catch { /* swallow */ }

            // 3) TranscodedWallpaper 缓存（Win7+，系统转码后存这里，本质是 JPG 流）
            try
            {
                var transcoded = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Windows", "Themes", "TranscodedWallpaper");
                if (File.Exists(transcoded)) return transcoded;
            }
            catch { /* swallow */ }

            return null;
        }

        /// <summary>把当前桌面壁纸保存到图库；返回保存后的图库内路径。</summary>
        public static async Task<string> SaveCurrentWallpaperToLibraryAsync()
        {
            var src = GetCurrentWallpaperPath();
            if (string.IsNullOrEmpty(src) || !File.Exists(src))
                throw new InvalidOperationException("无法读取当前桌面壁纸。");

            EnsureLibrary();

            var ext = Path.GetExtension(src)?.ToLowerInvariant();
            // TranscodedWallpaper 没有后缀，但内容是 JPG 流
            if (string.IsNullOrEmpty(ext) || !SupportedExtensions.Contains(ext))
                ext = ".jpg";

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var fileName = $"desktop_{stamp}{ext}";
            var dest = ResolveDestination(LibraryDirectory, fileName);

            using (var inS = new FileStream(src, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
            using (var outS = new FileStream(dest, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true))
            {
                await inS.CopyToAsync(outS).ConfigureAwait(false);
            }
            return dest;
        }

        /// <summary>把图库内某张图片设为桌面壁纸（含显示方式）。</summary>
        public static void SetWallpaper(string imagePath, WallpaperStyle style)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                throw new FileNotFoundException("图片不存在。", imagePath);

            // 1) 写注册表的 WallpaperStyle / TileWallpaper（必须先于 SPI_SETDESKWALLPAPER）
            string wallpaperStyleVal;
            string tileVal;
            switch (style)
            {
                case WallpaperStyle.Tile:
                    wallpaperStyleVal = "0"; tileVal = "1"; break;
                case WallpaperStyle.Center:
                    wallpaperStyleVal = "0"; tileVal = "0"; break;
                case WallpaperStyle.Stretch:
                    wallpaperStyleVal = "2"; tileVal = "0"; break;
                case WallpaperStyle.Fit:
                    wallpaperStyleVal = "6"; tileVal = "0"; break;
                case WallpaperStyle.Fill:
                default:
                    wallpaperStyleVal = "10"; tileVal = "0"; break;
            }

            using (var key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", writable: true))
            {
                if (key != null)
                {
                    key.SetValue("WallpaperStyle", wallpaperStyleVal, RegistryValueKind.String);
                    key.SetValue("TileWallpaper", tileVal, RegistryValueKind.String);
                }
            }

            // 2) SPI_SETDESKWALLPAPER（路径以 absolute 提供，Win7+ 自动转码非 BMP）
            var ok = SystemParametersInfo(
                SPI_SETDESKWALLPAPER,
                0,
                imagePath,
                SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);

            if (!ok)
            {
                var err = Marshal.GetLastWin32Error();
                throw new InvalidOperationException($"SystemParametersInfo 失败（Win32 错误码 {err}）。");
            }
        }

        // ===================== 私有辅助 =====================

        private static string ResolveDestination(string dir, string fileName)
        {
            var safeName = string.IsNullOrWhiteSpace(fileName) ? "wallpaper.jpg" : fileName;
            foreach (var c in Path.GetInvalidFileNameChars()) safeName = safeName.Replace(c, '_');
            var dest = Path.Combine(dir, safeName);
            if (!File.Exists(dest)) return dest;

            var name = Path.GetFileNameWithoutExtension(safeName);
            var ext = Path.GetExtension(safeName);
            for (int i = 1; i < 9999; i++)
            {
                var candidate = Path.Combine(dir, $"{name}_{i}{ext}");
                if (!File.Exists(candidate)) return candidate;
            }
            return Path.Combine(dir, $"{name}_{Guid.NewGuid():N}{ext}");
        }

        // ===================== Win32 互操作 =====================

        private const int MAX_PATH = 260;
        private const uint SPI_GETDESKWALLPAPER = 0x0073;
        private const uint SPI_SETDESKWALLPAPER = 0x0014;
        private const uint SPIF_UPDATEINIFILE = 0x01;
        private const uint SPIF_SENDCHANGE = 0x02;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SystemParametersInfo(uint uAction, uint uParam, StringBuilder lpvParam, uint fuWinIni);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SystemParametersInfo(uint uAction, uint uParam, string lpvParam, uint fuWinIni);
    }
}
