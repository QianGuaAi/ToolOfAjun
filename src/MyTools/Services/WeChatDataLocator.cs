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

            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Documents\WeChat Files  (旧版桌面微信)
            var legacyRoot = Path.Combine(docs, "WeChat Files");
            AddRootsFromBasePath(result, legacyRoot, WeChatVariant.LegacyWeChat);

            // Documents\Tencent Files  (部分桌面微信版本)
            var tencentFilesRoot = Path.Combine(docs, "Tencent Files");
            AddRootsFromBasePath(result, tencentFilesRoot, WeChatVariant.LegacyWeChat);

            // 如果 Documents 下没有子目录，说明 Documents 是空目录（Junction 指向 D:\），直接从 D:\Documents 扫描
            try
            {
                if (!Directory.EnumerateFileSystemEntries(docs).Any())
                {
                    var realDocs = TryResolveRealDocumentsPath();
                    if (!string.IsNullOrEmpty(realDocs) && Directory.Exists(realDocs))
                    {
                        AddRootsFromBasePath(result, Path.Combine(realDocs, "WeChat Files"), WeChatVariant.LegacyWeChat);
                        AddRootsFromBasePath(result, Path.Combine(realDocs, "Tencent Files"), WeChatVariant.LegacyWeChat);
                        AddRootsFromBasePath(result, Path.Combine(realDocs, "xwechat_files"), WeChatVariant.XWechat);
                        AddRootsFromBasePath(result, Path.Combine(realDocs, "xwechat_files", "all_users"), WeChatVariant.XWechat);
                    }
                }
            }
            catch
            {
                // ignore
            }

            // Documents\xwechat_files\{wxid}  (新版微信)
            var xwechatRoot = Path.Combine(docs, "xwechat_files");
            AddRootsFromBasePath(result, xwechatRoot, WeChatVariant.XWechat);

            // Documents\xwechat_files\all_users\{wxid}
            var xwechatAllUsersRoot = Path.Combine(xwechatRoot, "all_users");
            AddRootsFromBasePath(result, xwechatAllUsersRoot, WeChatVariant.XWechat);

            // AppData\Roaming\Tencent\xwechat_files  (旧版 XWechat)
            var xwechatAppDataRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Tencent", "xwechat_files");
            AddRootsFromBasePath(result, xwechatAppDataRoot, WeChatVariant.XWechat);

            // 注册表自定义目录：HKCU\Software\Tencent\WeChat  和  HKCU\Software\Tencent\Weixin
            var registryPaths = ReadAllRegistryFileSavePaths();
            foreach (var kvp in registryPaths)
            {
                var variant = kvp.Value.IndexOf("xwechat_files", StringComparison.OrdinalIgnoreCase) >= 0
                    ? WeChatVariant.XWechat
                    : WeChatVariant.LegacyWeChat;
                AddRootsFromBasePath(result, kvp.Value, variant);
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

        private static IReadOnlyDictionary<string, string> ReadAllRegistryFileSavePaths()
        {
            var results = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var keys = new[]
            {
                @"Software\Tencent\WeChat",
                @"Software\Tencent\Weixin"
            };
            foreach (var keyPath in keys)
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
                    {
                        var value = key?.GetValue("FileSavePath") as string;
                        if (!string.IsNullOrWhiteSpace(value) && !value.StartsWith("MyDocument:", StringComparison.OrdinalIgnoreCase))
                        {
                            results[keyPath] = value;
                        }
                    }
                }
                catch
                {
                    // ignore
                }
            }
            return results;
        }

        private static string TryResolveRealDocumentsPath()
        {
            try
            {
                // 读取 User Shell Folders 中的可扩展字符串，并展开环境变量
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"))
                {
                    var value = key?.GetValue("Personal") as string;
                    if (string.IsNullOrWhiteSpace(value)) return null;

                    var expanded = Environment.ExpandEnvironmentVariables(value);
                    var resolved = Path.GetFullPath(expanded);

                    // 只有当展开后的路径与标准 Documents 不同且存在时，才是真正的路径
                    var standardDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    if (!string.Equals(resolved, standardDocs, StringComparison.OrdinalIgnoreCase)
                        && Directory.Exists(resolved))
                    {
                        return resolved;
                    }
                }
            }
            catch
            {
                // ignore
            }
            return null;
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
