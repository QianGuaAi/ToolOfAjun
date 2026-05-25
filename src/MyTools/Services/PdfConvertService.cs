using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace MyTools.Services
{
    public static class PdfConvertService
    {
        public static bool IsPdfRenderSupported => OsVersionService.IsWindows8OrLater;

        public static Task<ConvertResult> ConvertImageToPdfAsync(string inputPath, CancellationToken ct)
        {
            return Task.Run(() =>
            {
                try
                {
                    ct.ThrowIfCancellationRequested();
                    if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
                    {
                        return new ConvertResult { Success = false, Message = "源图片不存在。" };
                    }

                    var inputSize = new FileInfo(inputPath).Length;
                    var outputPath = BuildUniqueOutputPath(
                        Path.GetDirectoryName(inputPath) ?? AppDomain.CurrentDomain.BaseDirectory,
                        Path.GetFileNameWithoutExtension(inputPath),
                        "pdf",
                        string.Empty);

                    byte[] pdfBytes;
                    using (var original = Image.FromFile(inputPath))
                    {
                        ct.ThrowIfCancellationRequested();
                        var dpiX = NormalizeDpi(original.HorizontalResolution);
                        var dpiY = NormalizeDpi(original.VerticalResolution);
                        var pageWidth = Math.Max(1.0, original.Width * 72.0 / dpiX);
                        var pageHeight = Math.Max(1.0, original.Height * 72.0 / dpiY);
                        var jpegBytes = CreateJpegBytes(original, dpiX, dpiY);
                        pdfBytes = CreateSinglePagePdf(jpegBytes, original.Width, original.Height, pageWidth, pageHeight);
                    }

                    ct.ThrowIfCancellationRequested();
                    File.WriteAllBytes(outputPath, pdfBytes);
                    var outputSize = new FileInfo(outputPath).Length;
                    return new ConvertResult
                    {
                        Success = true,
                        OutputPath = outputPath,
                        InputSize = inputSize,
                        OutputSize = outputSize,
                        Message = $"图片转 PDF 完成：{Path.GetFileName(outputPath)}"
                    };
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    return new ConvertResult { Success = false, Message = "图片转 PDF 失败：" + ex.Message };
                }
            }, ct);
        }

        public static Task<ConvertResult> ConvertPdfToImagesAsync(string inputPath, CancellationToken ct)
        {
            return Task.Run(async () =>
            {
                try
                {
                    ct.ThrowIfCancellationRequested();
                    if (!IsPdfRenderSupported)
                    {
                        return new ConvertResult { Success = false, Message = "PDF 转图片需要 Windows 8 或更高版本。" };
                    }

                    if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
                    {
                        return new ConvertResult { Success = false, Message = "源 PDF 不存在。" };
                    }

                    var inputSize = new FileInfo(inputPath).Length;
                    var outputDirectory = Path.GetDirectoryName(inputPath) ?? AppDomain.CurrentDomain.BaseDirectory;
                    var baseName = Path.GetFileNameWithoutExtension(inputPath);
                    var storageFile = await StorageFile.GetFileFromPathAsync(inputPath);
                    var document = await PdfDocument.LoadFromFileAsync(storageFile);
                    if (document.PageCount == 0)
                    {
                        return new ConvertResult { Success = false, Message = "PDF 没有可转换的页面。" };
                    }

                    var outputPaths = new List<string>();
                    for (uint pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
                    {
                        ct.ThrowIfCancellationRequested();
                        using (var page = document.GetPage(pageIndex))
                        using (var stream = new InMemoryRandomAccessStream())
                        {
                            await page.RenderToStreamAsync(stream);
                            ct.ThrowIfCancellationRequested();
                            if (stream.Size == 0 || stream.Size > int.MaxValue)
                            {
                                return new ConvertResult { Success = false, Message = "PDF 页面渲染结果异常。" };
                            }

                            var outputPath = BuildUniqueOutputPath(
                                outputDirectory,
                                baseName + "_page_" + (pageIndex + 1).ToString("000", CultureInfo.InvariantCulture),
                                "png",
                                string.Empty);
                            var bytes = new byte[(int)stream.Size];
                            using (var reader = new DataReader(stream.GetInputStreamAt(0)))
                            {
                                await reader.LoadAsync((uint)bytes.Length);
                                reader.ReadBytes(bytes);
                            }

                            File.WriteAllBytes(outputPath, bytes);
                            outputPaths.Add(outputPath);
                        }
                    }

                    var outputSize = 0L;
                    foreach (var path in outputPaths)
                    {
                        outputSize += new FileInfo(path).Length;
                    }

                    return new ConvertResult
                    {
                        Success = true,
                        OutputPath = outputPaths.Count == 1 ? outputPaths[0] : outputDirectory,
                        InputSize = inputSize,
                        OutputSize = outputSize,
                        Message = $"PDF 转图片完成：{outputPaths.Count} 页"
                    };
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    return new ConvertResult { Success = false, Message = "PDF 转图片失败：" + ex.Message };
                }
            }, ct);
        }

        private static byte[] CreateJpegBytes(Image original, float dpiX, float dpiY)
        {
            using (var bitmap = new Bitmap(original.Width, original.Height, PixelFormat.Format24bppRgb))
            {
                bitmap.SetResolution(dpiX, dpiY);
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.DrawImage(original, 0, 0, original.Width, original.Height);
                }

                using (var stream = new MemoryStream())
                using (var encoderParams = new EncoderParameters(1))
                {
                    encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);
                    var codec = GetEncoder(ImageFormat.Jpeg);
                    if (codec != null)
                    {
                        bitmap.Save(stream, codec, encoderParams);
                    }
                    else
                    {
                        bitmap.Save(stream, ImageFormat.Jpeg);
                    }

                    return stream.ToArray();
                }
            }
        }

        private static byte[] CreateSinglePagePdf(byte[] jpegBytes, int imageWidth, int imageHeight, double pageWidth, double pageHeight)
        {
            using (var stream = new MemoryStream())
            {
                var offsets = new List<long>();
                WriteAscii(stream, "%PDF-1.4\n");
                WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>\n");
                WriteObject(stream, offsets, 2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>\n");
                WriteObject(stream, offsets, 3,
                    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + FormatPdfNumber(pageWidth) + " " + FormatPdfNumber(pageHeight) + "] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>\n");

                offsets.Add(stream.Position);
                WriteAscii(stream, "4 0 obj\n<< /Type /XObject /Subtype /Image /Width " + imageWidth.ToString(CultureInfo.InvariantCulture) + " /Height " + imageHeight.ToString(CultureInfo.InvariantCulture) + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpegBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
                stream.Write(jpegBytes, 0, jpegBytes.Length);
                WriteAscii(stream, "\nendstream\nendobj\n");

                var content = "q\n" + FormatPdfNumber(pageWidth) + " 0 0 " + FormatPdfNumber(pageHeight) + " 0 0 cm\n/Im0 Do\nQ\n";
                var contentBytes = Encoding.ASCII.GetBytes(content);
                offsets.Add(stream.Position);
                WriteAscii(stream, "5 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
                stream.Write(contentBytes, 0, contentBytes.Length);
                WriteAscii(stream, "endstream\nendobj\n");

                var xrefOffset = stream.Position;
                WriteAscii(stream, "xref\n0 6\n0000000000 65535 f \n");
                foreach (var offset in offsets)
                {
                    WriteAscii(stream, offset.ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n \n");
                }
                WriteAscii(stream, "trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n" + xrefOffset.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
                return stream.ToArray();
            }
        }

        private static void WriteObject(Stream stream, ICollection<long> offsets, int number, string body)
        {
            offsets.Add(stream.Position);
            WriteAscii(stream, number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + body + "endobj\n");
        }

        private static void WriteAscii(Stream stream, string value)
        {
            var bytes = Encoding.ASCII.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string BuildUniqueOutputPath(string directory, string baseName, string extension, string suffix)
        {
            Directory.CreateDirectory(directory);
            var safeBaseName = SanitizeFileName(baseName);
            var safeSuffix = suffix ?? string.Empty;
            var safeExtension = extension.TrimStart('.');
            var first = Path.Combine(directory, safeBaseName + safeSuffix + "." + safeExtension);
            if (!File.Exists(first))
            {
                return first;
            }

            for (var i = 2; i < 1000; i++)
            {
                var candidate = Path.Combine(directory, safeBaseName + safeSuffix + "_" + i.ToString(CultureInfo.InvariantCulture) + "." + safeExtension);
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return Path.Combine(directory, safeBaseName + safeSuffix + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture) + "." + safeExtension);
        }

        private static string SanitizeFileName(string value)
        {
            var safeName = string.IsNullOrWhiteSpace(value) ? "output" : value.Trim();
            foreach (var ch in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(ch, '_');
            }

            return string.IsNullOrWhiteSpace(safeName) ? "output" : safeName;
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            foreach (var codec in ImageCodecInfo.GetImageEncoders())
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }

            return null;
        }

        private static float NormalizeDpi(float dpi)
        {
            return float.IsNaN(dpi) || float.IsInfinity(dpi) || dpi < 10 || dpi > 2400 ? 96f : dpi;
        }

        private static string FormatPdfNumber(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
