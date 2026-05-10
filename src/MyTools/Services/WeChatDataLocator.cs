using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace MyTools.Services
{
    public sealed class WeChatDataLocator
    {
        private static readonly Regex WxIdPattern = new Regex(@"^(wxid_[A-Za-z0-9]+|\d+_.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public IReadOnlyList<WeChatRoot> LocateRoots()
        {
            var result = new Dictionary<string, WeChatRoot>(StringComparer.OrdinalIgnoreCase);

            var docsWeChatRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WeChat Files");
            AddRootsFromBasePath(result, docsWeChatRoot, WeChatVariant.LegacyWeChat);

            var xwechatRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Tencent", "xwechat_files");
            AddRootsFromBasePath(result, xwechatRoot, WeChatVariant.XWechat);

            var registryPath = ReadRegistryFileSavePath();
            if (!string.IsNullOrWhiteSpace(registryPath))
            {
                var variant = registryPath.IndexOf("xwechat_files", StringComparison.OrdinalIgnoreCase) >= 0
                    ? WeChatVariant.XWechat
                    : WeChatVariant.LegacyWeChat;
                AddRootsFromBasePath(result, registryPath, variant);
            }

            return result.Values
                .OrderBy(x => x.Variant)
                .ThenBy(x => x.WxIdOrUserName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void AddRootsFromBasePath(
            IDictionary<string, WeChatRoot> target,
            string basePath,
            WeChatVariant variant)
        {
            if (string.IsNullOrWhiteSpace(basePath))
            {
                return;
            }

            string normalizedBasePath;
            try
            {
                normalizedBasePath = Path.GetFullPath(basePath);
            }
            catch
            {
                return;
            }

            if (!Directory.Exists(normalizedBasePath))
            {
                return;
            }

            string[] children;
            try
            {
                children = Directory.GetDirectories(normalizedBasePath, "*", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                children = Array.Empty<string>();
            }

            foreach (var child in children)
            {
                var name = Path.GetFileName(child.TrimEnd('\\'));
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!WxIdPattern.IsMatch(name))
                {
                    continue;
                }

                var normalizedChild = Normalize(child);
                if (string.IsNullOrWhiteSpace(normalizedChild))
                {
                    continue;
                }

                if (target.ContainsKey(normalizedChild))
                {
                    continue;
                }

                target[normalizedChild] = new WeChatRoot
                {
                    WxIdOrUserName = name,
                    RootPath = normalizedChild,
                    Variant = variant
                };
            }
        }

        private static string ReadRegistryFileSavePath()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Tencent\WeChat"))
                {
                    var value = key?.GetValue("FileSavePath") as string;
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        return string.Empty;
                    }

                    if (value.StartsWith("MyDocument:", StringComparison.OrdinalIgnoreCase))
                    {
                        return string.Empty;
                    }

                    return value;
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Reading WeChat FileSavePath from registry failed.");
                return string.Empty;
            }
        }

        private static string Normalize(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(path).TrimEnd('\\');
            }
            catch
            {
                return path.Trim().TrimEnd('\\');
            }
        }
    }

    public sealed class WeChatRoot
    {
        public string WxIdOrUserName { get; set; }
        public string RootPath { get; set; }
        public WeChatVariant Variant { get; set; }
        public string VariantDisplay => Variant == WeChatVariant.XWechat ? "XWechat" : "LegacyWeChat";
    }

    public enum WeChatVariant
    {
        LegacyWeChat,
        XWechat
    }
}
