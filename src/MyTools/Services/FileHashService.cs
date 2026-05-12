using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public class FileHashResult
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public long FileSize { get; set; }
        public string Md5 { get; set; }
        public string Sha1 { get; set; }
        public string Sha256 { get; set; }
        public string Crc32 { get; set; }
    }

    public static class FileHashService
    {
        public static Task<FileHashResult> ComputeAsync(string filePath, IProgress<string> progress, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException("文件不存在。", filePath);
            }

            return Task.Run(() =>
            {
                var info = new FileInfo(filePath);
                var result = new FileHashResult
                {
                    FilePath = filePath,
                    FileName = info.Name,
                    FileSize = info.Length
                };

                progress?.Report("计算中（单遍扫描 MD5 / SHA-1 / SHA-256 / CRC32）…");

                // Single-pass: read file once, feed into 3 hash algorithms + CRC32 accumulator.
                using (var md5 = new MD5CryptoServiceProvider())
                using (var sha1 = new SHA1CryptoServiceProvider())
                using (var sha256 = new SHA256CryptoServiceProvider())
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, useAsync: false))
                {
                    var buffer = new byte[1024 * 1024]; // 1 MB buffer — better sequential throughput
                    uint crc = 0xFFFFFFFFu;
                    long totalRead = 0;
                    long lastReport = 0;
                    int read;
                    while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ct.ThrowIfCancellationRequested();

                        // Feed hashes (TransformBlock: incremental update)
                        md5.TransformBlock(buffer, 0, read, null, 0);
                        sha1.TransformBlock(buffer, 0, read, null, 0);
                        sha256.TransformBlock(buffer, 0, read, null, 0);

                        // CRC32
                        for (int i = 0; i < read; i++)
                        {
                            crc = (crc >> 8) ^ Crc32Table[(crc ^ buffer[i]) & 0xFF];
                        }

                        totalRead += read;
                        if (info.Length > 0 && totalRead - lastReport >= 16 * 1024 * 1024)
                        {
                            lastReport = totalRead;
                            progress?.Report($"已扫描 {totalRead * 100.0 / info.Length:0.#}% ({totalRead / 1024 / 1024} / {info.Length / 1024 / 1024} MB)");
                        }
                    }

                    md5.TransformFinalBlock(buffer, 0, 0);
                    sha1.TransformFinalBlock(buffer, 0, 0);
                    sha256.TransformFinalBlock(buffer, 0, 0);

                    result.Md5 = BitConverter.ToString(md5.Hash).Replace("-", string.Empty).ToUpperInvariant();
                    result.Sha1 = BitConverter.ToString(sha1.Hash).Replace("-", string.Empty).ToUpperInvariant();
                    result.Sha256 = BitConverter.ToString(sha256.Hash).Replace("-", string.Empty).ToUpperInvariant();
                    result.Crc32 = (crc ^ 0xFFFFFFFFu).ToString("X8");
                }

                return result;
            }, ct);
        }

        private static readonly uint[] Crc32Table = BuildCrc32Table();

        private static uint[] BuildCrc32Table()
        {
            const uint poly = 0xEDB88320u;
            var table = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int j = 0; j < 8; j++)
                {
                    c = (c & 1) != 0 ? (poly ^ (c >> 1)) : (c >> 1);
                }
                table[i] = c;
            }
            return table;
        }

    }
}
