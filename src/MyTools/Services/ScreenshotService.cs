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
                return ConvertToBitmapSource(bitmap);
            }
        }

        public static BitmapSource ConvertToBitmapSource(Bitmap bitmap)
        {
            using (var ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                var decoder = new PngBitmapDecoder(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                var frame = decoder.Frames[0];
                frame.Freeze();
                return frame;
            }
        }
    }
}
