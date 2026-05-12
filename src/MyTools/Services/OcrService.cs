using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Windows.Globalization;
using Windows.Media.Ocr;
using Windows.Storage.Streams;
using WinRtImaging = Windows.Graphics.Imaging;

namespace MyTools.Services
{
    /// <summary>
    /// 基于 Windows.Media.Ocr 的本地 OCR 识别，仅支持 Windows 10 1903 及以上。
    /// 中文识别需要系统已安装中文语言包（控制面板 → 区域 → 语言 → 添加"中文（简体）"）。
    /// </summary>
    public static class OcrService
    {
        public static bool IsSupported => OsVersionService.IsWindows10OrGreater;

        public static async Task<string> RecognizeAsync(BitmapSource bitmap)
        {
            if (!IsSupported)
            {
                throw new PlatformNotSupportedException("WindowsOCR 仅支持 Windows 10 1903 及以上系统。");
            }
            if (bitmap == null)
            {
                throw new ArgumentNullException(nameof(bitmap));
            }

            // 将 WPF BitmapSource 编码为 PNG，再交给 WinRT BitmapDecoder 解码为 SoftwareBitmap。
            byte[] pngBytes;
            using (var ms = new MemoryStream())
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                encoder.Save(ms);
                pngBytes = ms.ToArray();
            }

            using (var stream = new InMemoryRandomAccessStream())
            {
                using (var writer = new DataWriter(stream))
                {
                    writer.WriteBytes(pngBytes);
                    await writer.StoreAsync();
                    await writer.FlushAsync();
                    writer.DetachStream();
                }
                stream.Seek(0);

                var decoder = await WinRtImaging.BitmapDecoder.CreateAsync(stream);
                var softwareBitmap = await decoder.GetSoftwareBitmapAsync();

                var engine = ResolveOcrEngine();
                if (engine == null)
                {
                    throw new InvalidOperationException(
                        "未找到可用的 OCR 语言包。请在'设置 → 时间和语言 → 语言'中添加'中文（简体）'或'英语'，并勾选'可选功能 → 添加功能 → 中文（简体）光学字符识别'。");
                }

                var result = await engine.RecognizeAsync(softwareBitmap);
                return result?.Text ?? string.Empty;
            }
        }

        private static OcrEngine ResolveOcrEngine()
        {
            // 优先使用当前用户语言；fallback 到中文简体、英文。
            var engine = OcrEngine.TryCreateFromUserProfileLanguages();
            if (engine != null)
            {
                return engine;
            }

            var preferred = new[] { "zh-Hans-CN", "zh-Hans", "zh-CN", "en-US" };
            foreach (var tag in preferred)
            {
                try
                {
                    var lang = new Language(tag);
                    if (!OcrEngine.IsLanguageSupported(lang))
                    {
                        continue;
                    }
                    var e = OcrEngine.TryCreateFromLanguage(lang);
                    if (e != null)
                    {
                        return e;
                    }
                }
                catch
                {
                    // 忽略不支持的语言标签
                }
            }

            // 最后尝试系统内置任一可用语言
            var available = OcrEngine.AvailableRecognizerLanguages;
            var first = available?.FirstOrDefault();
            return first != null ? OcrEngine.TryCreateFromLanguage(first) : null;
        }
    }
}
