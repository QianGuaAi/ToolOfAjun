using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
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

            using (var bitmap = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
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
                // 把图先转成 32bpp BGRA 并复制像素，避免区域截图任意宽度在 24bpp DIB
                // 行对齐上被部分粘贴目标误读，表现为右侧缺失。
                var converted = src.Format == PixelFormats.Bgra32
                    ? src
                    : new FormatConvertedBitmap(src, PixelFormats.Bgra32, null, 0);
                finalBmp = CopyBitmapSource(converted);
                var dibBytes = CreateDibBytes(finalBmp);
                var pngBytes = CreatePngBytes(finalBmp);

                // STA 线程要求：保证当前在 UI 线程上调用
                var disp = System.Windows.Application.Current?.Dispatcher;
                if (disp == null || disp.CheckAccess())
                {
                    SetClipboardData(finalBmp, dibBytes, pngBytes);
                }
                else
                {
                    disp.Invoke(() => SetClipboardData(finalBmp, dibBytes, pngBytes));
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

        private static BitmapSource CopyBitmapSource(BitmapSource source)
        {
            var stride = checked(source.PixelWidth * 4);
            var pixels = new byte[checked(stride * source.PixelHeight)];
            source.CopyPixels(pixels, stride, 0);
            var copy = BitmapSource.Create(
                source.PixelWidth,
                source.PixelHeight,
                NormalizeDpi(source.DpiX),
                NormalizeDpi(source.DpiY),
                PixelFormats.Bgra32,
                null,
                pixels,
                stride);
            copy.Freeze();
            return copy;
        }

        private static void SetClipboardData(BitmapSource bitmap, byte[] dibBytes, byte[] pngBytes)
        {
            var data = new DataObject();
            data.SetImage(bitmap);
            data.SetData(DataFormats.Dib, new MemoryStream(dibBytes), false);
            data.SetData("PNG", new MemoryStream(pngBytes), false);
            data.SetData("image/png", new MemoryStream(pngBytes), false);
            Clipboard.SetDataObject(data, true);
        }

        private static byte[] CreatePngBytes(BitmapSource bitmap)
        {
            using (var stream = new MemoryStream())
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                encoder.Save(stream);
                return stream.ToArray();
            }
        }

        private static byte[] CreateDibBytes(BitmapSource bitmap)
        {
            var width = bitmap.PixelWidth;
            var height = bitmap.PixelHeight;
            var stride = checked(width * 4);
            var pixels = new byte[checked(stride * height)];
            bitmap.CopyPixels(pixels, stride, 0);

            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write(40); // BITMAPINFOHEADER
                writer.Write(width);
                writer.Write(height); // positive height = bottom-up DIB
                writer.Write((short)1);
                writer.Write((short)32);
                writer.Write(0); // BI_RGB
                writer.Write(checked(stride * height));
                writer.Write(0);
                writer.Write(0);
                writer.Write(0);
                writer.Write(0);

                for (var y = height - 1; y >= 0; y--)
                {
                    writer.Write(pixels, y * stride, stride);
                }

                writer.Flush();
                return stream.ToArray();
            }
        }

        private static double NormalizeDpi(double dpi)
        {
            return double.IsNaN(dpi) || double.IsInfinity(dpi) || dpi < 10 || dpi > 2400 ? 96.0 : dpi;
        }

        public static BitmapSource ConvertToBitmapSource(Bitmap bitmap)
        {
            // 直接锁定 GDI 位图像素并通过 BitmapSource.Create 复制为一份独立、已冻结的 WPF 位图。
            // 这样产生的 BitmapSource 与原 GDI Bitmap 解耦，且 Freeze 之后可跨线程访问。
            // 之前用 PngBitmapDecoder 的方式在某些条件下 Freeze 后仍会触发"调用线程无法访问此对象"。
            var rect = new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
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
