using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyTools.Services
{
    public static class ScreenshotService
    {
        private static readonly object Gate = new object();

        /// <summary>
        /// Opens a region picker on the UI thread, then captures and saves a PNG next to the executable.
        /// </summary>
        public static async Task<string> CaptureRegionInteractiveAsync()
        {
            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher == null)
            {
                throw new InvalidOperationException("应用程序尚未初始化。");
            }

            Rectangle? region = null;
            await dispatcher.InvokeAsync(() =>
            {
                using (var picker = new RegionPickerForm())
                {
                    var result = picker.ShowDialog();
                    if (result == DialogResult.OK && !picker.Selection.IsEmpty)
                    {
                        region = picker.Selection;
                    }
                }
            });

            if (region == null || region.Value.Width <= 0 || region.Value.Height <= 0)
            {
                return null;
            }

            return await Task.Run(() => CaptureAndSave(region.Value)).ConfigureAwait(false);
        }

        private static string CaptureAndSave(Rectangle region)
        {
            Directory.CreateDirectory(GetScreenshotsDirectory());

            string path = BuildFilePath();
            lock (Gate)
            {
                // Avoid rare collisions when saving within the same second.
                while (File.Exists(path))
                {
                    path = BuildFilePath();
                }

                using (var bitmap = new Bitmap(region.Width, region.Height, PixelFormat.Format32bppArgb))
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(region.Left, region.Top, 0, 0, region.Size, CopyPixelOperation.SourceCopy);
                    bitmap.Save(path, ImageFormat.Png);
                }
            }

            return path;
        }

        public static string GetScreenshotsDirectory()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Screenshots");
        }

        private static string BuildFilePath()
        {
            string fileName = $"MyTools_{DateTime.Now:yyyyMMdd_HHmmss}.png";
            return Path.Combine(GetScreenshotsDirectory(), fileName);
        }
    }
}
