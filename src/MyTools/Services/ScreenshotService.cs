using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace MyTools.Services
{
    public static class ScreenshotService
    {
        public static BitmapSource CaptureFullScreen()
        {
            var left   = (int)SystemParameters.VirtualScreenLeft;
            var top    = (int)SystemParameters.VirtualScreenTop;
            var width  = (int)SystemParameters.VirtualScreenWidth;
            var height = (int)SystemParameters.VirtualScreenHeight;

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
                return decoder.Frames[0];
            }
        }
    }
}
