using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public static class CodexConfigProfileService
    {
        public const string ConfigFileName = "config.toml";
        public const string AuthFileName = "auth.json";

        public static async Task<CodexConfigProfileSourceFiles> ReadProfileFromFolderAsync(string sourceFolderPath, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(sourceFolderPath))
            {
                throw new InvalidOperationException("配置文件夹路径不能为空。");
            }

            var sourceFolder = Path.GetFullPath(sourceFolderPath);
            if (!Directory.Exists(sourceFolder))
            {
                throw new DirectoryNotFoundException("配置文件夹不存在：" + sourceFolder);
            }

            var configFilePath = Path.Combine(sourceFolder, ConfigFileName);
            var authFilePath = Path.Combine(sourceFolder, AuthFileName);

            var missingFiles = new List<string>();
            if (!File.Exists(configFilePath))
            {
                missingFiles.Add(ConfigFileName);
            }

            if (!File.Exists(authFilePath))
            {
                missingFiles.Add(AuthFileName);
            }

            if (missingFiles.Count > 0)
            {
                throw new FileNotFoundException("配置文件夹缺少：" + string.Join(", ", missingFiles));
            }

            return new CodexConfigProfileSourceFiles
            {
                SourceFolderPath = sourceFolder,
                ConfigTomlBytes = await ReadAllBytesAsync(configFilePath, cancellationToken).ConfigureAwait(false),
                AuthJsonBytes = await ReadAllBytesAsync(authFilePath, cancellationToken).ConfigureAwait(false)
            };
        }

        public static async Task<CodexConfigApplyResult> ApplyAsync(byte[] configTomlBytes, byte[] authJsonBytes, CancellationToken cancellationToken)
        {
            if (configTomlBytes == null)
            {
                throw new InvalidOperationException("未找到 config.toml 的内容。");
            }

            if (authJsonBytes == null)
            {
                throw new InvalidOperationException("未找到 auth.json 的内容。");
            }

            var targetFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".codex");
            Directory.CreateDirectory(targetFolder);

            var targetConfigPath = Path.GetFullPath(Path.Combine(targetFolder, ConfigFileName));
            var targetAuthPath = Path.GetFullPath(Path.Combine(targetFolder, AuthFileName));

            await WriteAllBytesAsync(targetConfigPath, configTomlBytes, cancellationToken).ConfigureAwait(false);
            await WriteAllBytesAsync(targetAuthPath, authJsonBytes, cancellationToken).ConfigureAwait(false);

            AppLogService.Information(
                "Applied embedded Codex config profile to {TargetFolder}",
                targetFolder);

            return new CodexConfigApplyResult
            {
                TargetFolderPath = targetFolder,
                AppliedFilePaths = new List<string>
                {
                    targetConfigPath,
                    targetAuthPath
                }
            };
        }

        public static async Task<CodexConfigApplyResult> ApplyAsync(string sourceFolderPath, CancellationToken cancellationToken)
        {
            var profileFiles = await ReadProfileFromFolderAsync(sourceFolderPath, cancellationToken).ConfigureAwait(false);
            return await ApplyAsync(profileFiles.ConfigTomlBytes, profileFiles.AuthJsonBytes, cancellationToken).ConfigureAwait(false);
        }

        public static string ProtectBytesToBase64(byte[] plainBytes)
        {
            if (plainBytes == null)
            {
                return string.Empty;
            }

            var protectedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        public static byte[] UnprotectBytesFromBase64(string protectedBase64)
        {
            if (string.IsNullOrWhiteSpace(protectedBase64))
            {
                return null;
            }

            try
            {
                var protectedBytes = Convert.FromBase64String(protectedBase64);
                return ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("配置内容解密失败，可能来自其他 Windows 用户或数据已损坏。", ex);
            }
        }

        private static async Task<byte[]> ReadAllBytesAsync(string filePath, CancellationToken cancellationToken)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
            using (var memoryStream = new MemoryStream())
            {
                await stream.CopyToAsync(memoryStream, 81920, cancellationToken).ConfigureAwait(false);
                return memoryStream.ToArray();
            }
        }

        private static async Task WriteAllBytesAsync(string filePath, byte[] bytes, CancellationToken cancellationToken)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
            {
                await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    public class CodexConfigProfileSourceFiles
    {
        public string SourceFolderPath { get; set; }
        public byte[] ConfigTomlBytes { get; set; }
        public byte[] AuthJsonBytes { get; set; }
    }

    public class CodexConfigApplyResult
    {
        public string TargetFolderPath { get; set; }
        public List<string> AppliedFilePaths { get; set; } = new List<string>();
    }
}
