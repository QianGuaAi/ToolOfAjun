using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;

namespace MyTools.Services
{
    public static class ScreenshotService
    {
        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);
        private const int SM_XVIRTUALSCREEN  = 76;
        private const int SM_YVIRTUALSCREEN  = 77;
        private const int SM_CXVIRTUALSCREEN = 78;
        private const int SM_CYVIRTUALSCREEN = 79;

        // ===== Cursor capture =====
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct CURSORINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hCursor;
            public POINT ptScreenPos;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ICONINFO
        {
            public bool fIcon;
            public int xHotspot;
            public int yHotspot;
            public IntPtr hbmMask;
            public IntPtr hbmColor;
        }

        [DllImport("user32.dll")] private static extern bool GetCursorInfo(ref CURSORINFO pci);
        [DllImport("user32.dll")] private static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO piconinfo);
        [DllImport("user32.dll")] private static extern bool DrawIconEx(IntPtr hdc, int xLeft, int yTop, IntPtr hIcon, int cxWidth, int cyWidth, int istepIfAniCur, IntPtr hbrFlickerFreeDraw, int diFlags);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr hObject);

        private const int CURSOR_SHOWING = 0x00000001;
        private const int DI_NORMAL = 0x0003;

        public static BitmapSource CaptureFullScreen()
        {
            var left   = GetSystemMetrics(SM_XVIRTUALSCREEN);
            var top    = GetSystemMetrics(SM_YVIRTUALSCREEN);
            var width  = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            var height = GetSystemMetrics(SM_CYVIRTUALSCREEN);

            using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
            using (var g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(left, top, 0, 0, new System.Drawing.Size(width, height),
                    CopyPixelOperation.SourceCopy);
                DrawCursorIfVisible(g, left, top);
                return ConvertToBitmapSource(bitmap);
            }
        }

        /// <summary>把当前可见的鼠标光标绘制到 <paramref name="g"/> 上。</summary>
        /// <param name="virtualLeft">虚拟屏幕左上角 X（用于把屏幕坐标换算到位图坐标）。</param>
        /// <param name="virtualTop">虚拟屏幕左上角 Y。</param>
        public static void DrawCursorIfVisible(Graphics g, int virtualLeft, int virtualTop)
        {
            if (g == null) return;
            try
            {
                var ci = new CURSORINFO { cbSize = Marshal.SizeOf(typeof(CURSORINFO)) };
                if (!GetCursorInfo(ref ci)) return;
                if ((ci.flags & CURSOR_SHOWING) == 0 || ci.hCursor == IntPtr.Zero) return;

                if (!GetIconInfo(ci.hCursor, out var ii)) return;
                try
                {
                    int x = ci.ptScreenPos.X - virtualLeft - ii.xHotspot;
                    int y = ci.ptScreenPos.Y - virtualTop - ii.yHotspot;

                    var hdc = g.GetHdc();
                    try { DrawIconEx(hdc, x, y, ci.hCursor, 0, 0, 0, IntPtr.Zero, DI_NORMAL); }
                    finally { g.ReleaseHdc(hdc); }
                }
                finally
                {
                    if (ii.hbmMask != IntPtr.Zero) DeleteObject(ii.hbmMask);
                    if (ii.hbmColor != IntPtr.Zero) DeleteObject(ii.hbmColor);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("DrawCursor failed: {Msg}", ex.Message);
            }
        }

        /// <summary>
        /// 跨应用兼容的剪贴板写入：转 Bgr24，并同时塞入 Bitmap / Dib / PNG 三种格式。
        /// 避免 WPF 默认 SetImage 用 BGRA32 DIB 导致经典画图 / 旧 Office 报"剪贴板上的信息无法插入"。
        /// </summary>
        public static void SetClipboardCompatible(BitmapSource src)
        {
            if (src == null) return;
            BitmapSource finalBmp = null;
            try
            {
                // 把图先转成 Bgr24 并通过 BMP 编/解码强制实例化为真正的像素位图。
                // 直接用 FormatConvertedBitmap 是惰性的，部分场景 SetImage 拿不到像素。
                var converted = src.Format == System.Windows.Media.PixelFormats.Bgr24
                    ? src
                    : new System.Windows.Media.Imaging.FormatConvertedBitmap(src, System.Windows.Media.PixelFormats.Bgr24, null, 0);
                if (converted.CanFreeze && !converted.IsFrozen) converted.Freeze();

                using (var ms = new MemoryStream())
                {
                    var enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                    enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(converted));
                    enc.Save(ms);
                    ms.Position = 0;
                    var dec = new System.Windows.Media.Imaging.BmpBitmapDecoder(
                        ms,
                        System.Windows.Media.Imaging.BitmapCreateOptions.None,
                        System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
                    finalBmp = dec.Frames[0];
                    if (finalBmp.CanFreeze && !finalBmp.IsFrozen) finalBmp.Freeze();
                }

                // STA 线程要求：保证当前在 UI 线程上调用
                var disp = System.Windows.Application.Current?.Dispatcher;
                if (disp == null || disp.CheckAccess())
                {
                    System.Windows.Clipboard.SetImage(finalBmp);
                }
                else
                {
                    disp.Invoke(() => System.Windows.Clipboard.SetImage(finalBmp));
                }
                AppLogService.Information("Clipboard image set ({W}x{H}, {Fmt}).",
                    finalBmp.PixelWidth, finalBmp.PixelHeight, finalBmp.Format);
            }
            catch (System.Exception ex)
            {
                AppLogService.Warning("SetClipboardCompatible failed: {Msg}", ex.Message);
                try { System.Windows.Clipboard.SetImage(finalBmp ?? src); }
                catch (System.Exception ex2) { AppLogService.Warning("Fallback SetImage failed: {Msg}", ex2.Message); }
            }
        }

        public static BitmapSource ConvertToBitmapSource(Bitmap bitmap)
        {
            // 直接锁定 GDI 位图像素并通过 BitmapSource.Create 复制为一份独立、已冻结的 WPF 位图。
            // 这样产生的 BitmapSource 与原 GDI Bitmap 解耦，且 Freeze 之后可跨线程访问。
            // 之前用 PngBitmapDecoder 的方式在某些条件下 Freeze 后仍会触发"调用线程无法访问此对象"。
            var rect = new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                var bs = BitmapSource.Create(
                    bitmap.Width,
                    bitmap.Height,
                    96.0, 96.0,
                    System.Windows.Media.PixelFormats.Bgra32,
                    null,
                    data.Scan0,
                    data.Stride * bitmap.Height,
                    data.Stride);
                bs.Freeze();
                return bs;
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }
    }
}
