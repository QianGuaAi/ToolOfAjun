using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace MyTools.Shared
{
    public static class MediaFileAssociationCore
    {
        public const string ProgId = "MyTools.MediaFile";
        public const string ApplicationName = "MyTools";
        public const string ProductDisplayName = "阿君的工具";

        public static readonly string[] VideoExtensions =
        {
            ".mp4", ".m4v", ".mov", ".wmv", ".avi", ".mkv", ".webm", ".mpg", ".mpeg", ".mpe",
            ".flv", ".3gp", ".3g2", ".ts", ".mts", ".m2ts", ".vob", ".asf", ".divx", ".ogv"
        };

        public static readonly string[] AudioExtensions =
        {
            ".mp3", ".wav", ".wma", ".m4a", ".aac", ".flac", ".ogg", ".oga", ".opus", ".alac",
            ".aiff", ".aif", ".ape", ".amr", ".mid", ".midi", ".mka"
        };

        public static string[] MediaExtensions => VideoExtensions.Concat(AudioExtensions)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(ext => ext, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        public static string[] GetExtensions(MediaAssociationKind kind)
        {
            switch (kind)
            {
                case MediaAssociationKind.Video:
                    return VideoExtensions.ToArray();
                case MediaAssociationKind.Audio:
                    return AudioExtensions.ToArray();
                default:
                    return MediaExtensions;
            }
        }

        public static bool IsSupportedMediaExtension(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
            {
                return false;
            }

            var normalized = extension.StartsWith(".", StringComparison.Ordinal)
                ? extension
                : "." + extension;
            return MediaExtensions.Contains(normalized, StringComparer.OrdinalIgnoreCase);
        }

        public static int RegisterForCurrentUser(string appPath)
        {
            return RegisterForCurrentUser(appPath, MediaAssociationKind.All);
        }

        public static int RegisterForCurrentUser(string appPath, MediaAssociationKind kind)
        {
            ValidateAppPath(appPath);
            var extensions = GetExtensions(kind);
            RegisterUnderRoot(Registry.CurrentUser, @"Software\Classes", @"Software\RegisteredApplications", @"Software\Ajun\MyTools\Capabilities", appPath, extensions);
            ResetCurrentUserChoices(extensions);
            NotifyShellAssociationChanged();
            return extensions.Length;
        }

        public static int RegisterForLocalMachine(string appPath)
        {
            return RegisterForLocalMachine(appPath, MediaAssociationKind.All);
        }

        public static int RegisterForLocalMachine(string appPath, MediaAssociationKind kind)
        {
            ValidateAppPath(appPath);
            var extensions = GetExtensions(kind);
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, GetWritableRegistryView()))
            {
                RegisterUnderRoot(root, @"SOFTWARE\Classes", @"SOFTWARE\RegisteredApplications", @"SOFTWARE\Ajun\MyTools\Capabilities", appPath, extensions);
            }

            NotifyShellAssociationChanged();
            return extensions.Length;
        }

        public static void UnregisterForCurrentUser()
        {
            UnregisterUnderRoot(Registry.CurrentUser, @"Software\Classes", @"Software\RegisteredApplications", @"Software\Ajun\MyTools\Capabilities");
            NotifyShellAssociationChanged();
        }

        public static void UnregisterForLocalMachine()
        {
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, GetWritableRegistryView()))
            {
                UnregisterUnderRoot(root, @"SOFTWARE\Classes", @"SOFTWARE\RegisteredApplications", @"SOFTWARE\Ajun\MyTools\Capabilities");
            }

            NotifyShellAssociationChanged();
        }

        public static int RestoreSystemDefaultForCurrentUser(MediaAssociationKind kind)
        {
            var extensions = GetExtensions(kind);
            var changed = 0;
            try
            {
                using (var classes = Registry.CurrentUser.OpenSubKey(@"Software\Classes", true))
                {
                    if (classes != null)
                    {
                        foreach (var extension in extensions)
                        {
                            using (var extensionKey = classes.OpenSubKey(extension, true))
                            {
                                if (extensionKey == null)
                                {
                                    continue;
                                }

                                var currentDefault = Convert.ToString(extensionKey.GetValue(string.Empty));
                                if (string.Equals(currentDefault, ProgId, StringComparison.OrdinalIgnoreCase))
                                {
                                    extensionKey.DeleteValue(string.Empty, false);
                                    changed++;
                                }

                                using (var openWithProgids = extensionKey.OpenSubKey("OpenWithProgids", true))
                                {
                                    if (openWithProgids != null && openWithProgids.GetValueNames().Contains(ProgId, StringComparer.OrdinalIgnoreCase))
                                    {
                                        openWithProgids.DeleteValue(ProgId, false);
                                    }
                                }
                            }
                        }
                    }
                }

                using (var fileExts = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts", true))
                {
                    if (fileExts != null)
                    {
                        foreach (var extension in extensions)
                        {
                            using (var extensionKey = fileExts.OpenSubKey(extension, true))
                            {
                                if (extensionKey == null)
                                {
                                    continue;
                                }

                                var progId = string.Empty;
                                using (var userChoice = extensionKey.OpenSubKey("UserChoice"))
                                {
                                    progId = Convert.ToString(userChoice?.GetValue("ProgId"));
                                }

                                if (string.Equals(progId, ProgId, StringComparison.OrdinalIgnoreCase))
                                {
                                    try
                                    {
                                        extensionKey.DeleteSubKeyTree("UserChoice", false);
                                        changed++;
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            NotifyShellAssociationChanged();
            return changed;
        }

        public static MediaAssociationStatus GetCurrentUserStatus(MediaAssociationKind kind)
        {
            var extensions = GetExtensions(kind);
            var status = new MediaAssociationStatus
            {
                Kind = kind,
                TotalCount = extensions.Length
            };

            using (var classes = Registry.CurrentUser.OpenSubKey(@"Software\Classes"))
            using (var fileExts = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"))
            {
                foreach (var extension in extensions)
                {
                    var classDefault = string.Empty;
                    var hasOpenWith = false;
                    using (var extensionKey = classes?.OpenSubKey(extension))
                    {
                        classDefault = Convert.ToString(extensionKey?.GetValue(string.Empty));
                        using (var openWithProgids = extensionKey?.OpenSubKey("OpenWithProgids"))
                        {
                            hasOpenWith = openWithProgids != null
                                && openWithProgids.GetValueNames().Contains(ProgId, StringComparer.OrdinalIgnoreCase);
                        }
                    }

                    var userChoice = string.Empty;
                    using (var extensionKey = fileExts?.OpenSubKey(extension))
                    using (var userChoiceKey = extensionKey?.OpenSubKey("UserChoice"))
                    {
                        userChoice = Convert.ToString(userChoiceKey?.GetValue("ProgId"));
                    }

                    var isDefault = string.Equals(userChoice, ProgId, StringComparison.OrdinalIgnoreCase)
                        || (string.IsNullOrWhiteSpace(userChoice) && string.Equals(classDefault, ProgId, StringComparison.OrdinalIgnoreCase));
                    if (isDefault)
                    {
                        status.AssociatedCount++;
                    }

                    if (string.Equals(classDefault, ProgId, StringComparison.OrdinalIgnoreCase))
                    {
                        status.ProgIdDefaultCount++;
                    }

                    if (hasOpenWith)
                    {
                        status.OpenWithCount++;
                    }
                }
            }

            return status;
        }

        private static void RegisterUnderRoot(RegistryKey root, string classesPath, string registeredApplicationsPath, string capabilitiesPath, string appPath, string[] extensions)
        {
            using (var classes = root.CreateSubKey(classesPath))
            {
                if (classes == null)
                {
                    throw new InvalidOperationException("无法写入文件关联注册表项。");
                }

                RegisterProgId(classes, appPath);
                RegisterApplication(classes, appPath, extensions);

                foreach (var extension in extensions)
                {
                    RegisterExtension(classes, extension);
                }
            }

            using (var capabilities = root.CreateSubKey(capabilitiesPath))
            {
                if (capabilities != null)
                {
                    capabilities.SetValue("ApplicationName", ProductDisplayName, RegistryValueKind.String);
                    capabilities.SetValue("ApplicationDescription", "使用阿君的工具打开音频和视频文件。", RegistryValueKind.String);
                    using (var fileAssociations = capabilities.CreateSubKey("FileAssociations"))
                    {
                        if (fileAssociations != null)
                        {
                            foreach (var extension in extensions)
                            {
                                fileAssociations.SetValue(extension, ProgId, RegistryValueKind.String);
                            }
                        }
                    }
                }
            }

            using (var registeredApplications = root.CreateSubKey(registeredApplicationsPath))
            {
                registeredApplications?.SetValue(ApplicationName, capabilitiesPath, RegistryValueKind.String);
            }
        }

        private static void RegisterProgId(RegistryKey classes, string appPath)
        {
            using (var progId = classes.CreateSubKey(ProgId))
            {
                if (progId == null)
                {
                    throw new InvalidOperationException("无法写入媒体文件类型。");
                }

                progId.SetValue(string.Empty, "阿君的工具媒体文件", RegistryValueKind.String);
                progId.SetValue("FriendlyTypeName", "阿君的工具媒体文件", RegistryValueKind.String);

                using (var icon = progId.CreateSubKey("DefaultIcon"))
                {
                    icon?.SetValue(string.Empty, Quote(appPath) + ",0", RegistryValueKind.String);
                }

                using (var command = progId.CreateSubKey(@"shell\open\command"))
                {
                    command?.SetValue(string.Empty, Quote(appPath) + " " + Quote("%1"), RegistryValueKind.String);
                }
            }
        }

        private static void RegisterApplication(RegistryKey classes, string appPath, string[] extensions)
        {
            using (var app = classes.CreateSubKey(@"Applications\MyTools.exe"))
            {
                app?.SetValue("FriendlyAppName", ProductDisplayName, RegistryValueKind.String);
            }

            using (var command = classes.CreateSubKey(@"Applications\MyTools.exe\shell\open\command"))
            {
                command?.SetValue(string.Empty, Quote(appPath) + " " + Quote("%1"), RegistryValueKind.String);
            }

            using (var supportedTypes = classes.CreateSubKey(@"Applications\MyTools.exe\SupportedTypes"))
            {
                if (supportedTypes != null)
                {
                    foreach (var extension in extensions)
                    {
                        supportedTypes.SetValue(extension, string.Empty, RegistryValueKind.String);
                    }
                }
            }
        }

        private static void RegisterExtension(RegistryKey classes, string extension)
        {
            using (var extensionKey = classes.CreateSubKey(extension))
            {
                if (extensionKey == null)
                {
                    return;
                }

                extensionKey.SetValue(string.Empty, ProgId, RegistryValueKind.String);
                extensionKey.SetValue("PerceivedType", IsAudioExtension(extension) ? "audio" : "video", RegistryValueKind.String);
                using (var openWithProgids = extensionKey.CreateSubKey("OpenWithProgids"))
                {
                    openWithProgids?.SetValue(ProgId, new byte[0], RegistryValueKind.Binary);
                }
            }
        }

        private static void ResetCurrentUserChoices(string[] extensions)
        {
            try
            {
                using (var fileExts = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts", true))
                {
                    if (fileExts == null)
                    {
                        return;
                    }

                    foreach (var extension in extensions)
                    {
                        try
                        {
                            using (var extensionKey = fileExts.OpenSubKey(extension, true))
                            {
                                extensionKey?.DeleteSubKeyTree("UserChoice", false);
                            }
                        }
                        catch
                        {
                            // Windows may protect UserChoice for some extensions. Registration still makes MyTools available.
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private static void UnregisterUnderRoot(RegistryKey root, string classesPath, string registeredApplicationsPath, string capabilitiesPath)
        {
            try
            {
                using (var classes = root.OpenSubKey(classesPath, true))
                {
                    if (classes != null)
                    {
                        foreach (var extension in MediaExtensions)
                        {
                            using (var extensionKey = classes.OpenSubKey(extension, true))
                            {
                                if (extensionKey == null)
                                {
                                    continue;
                                }

                                var currentDefault = Convert.ToString(extensionKey.GetValue(string.Empty));
                                if (string.Equals(currentDefault, ProgId, StringComparison.OrdinalIgnoreCase))
                                {
                                    extensionKey.DeleteValue(string.Empty, false);
                                }

                                using (var openWithProgids = extensionKey.OpenSubKey("OpenWithProgids", true))
                                {
                                    openWithProgids?.DeleteValue(ProgId, false);
                                }
                            }
                        }

                        classes.DeleteSubKeyTree(ProgId, false);
                        classes.DeleteSubKeyTree(@"Applications\MyTools.exe", false);
                    }
                }

                using (var registeredApplications = root.OpenSubKey(registeredApplicationsPath, true))
                {
                    registeredApplications?.DeleteValue(ApplicationName, false);
                }

                root.DeleteSubKeyTree(capabilitiesPath, false);
            }
            catch
            {
                // Best effort cleanup. File associations should not block uninstall.
            }
        }

        private static bool IsAudioExtension(string extension)
        {
            return AudioExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        private static void ValidateAppPath(string appPath)
        {
            if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            {
                throw new FileNotFoundException("找不到 MyTools.exe，无法写入文件关联。", appPath ?? string.Empty);
            }
        }

        private static RegistryView GetWritableRegistryView()
        {
            return Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
        }

        private static string Quote(string value)
        {
            return "\"" + value + "\"";
        }

        private static void NotifyShellAssociationChanged()
        {
            try
            {
                SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
            }
            catch
            {
            }
        }

        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
    }

    public enum MediaAssociationKind
    {
        All,
        Video,
        Audio
    }

    public sealed class MediaAssociationStatus
    {
        public MediaAssociationKind Kind { get; set; }
        public int TotalCount { get; set; }
        public int AssociatedCount { get; set; }
        public int ProgIdDefaultCount { get; set; }
        public int OpenWithCount { get; set; }

        public bool IsFullyAssociated => TotalCount > 0 && AssociatedCount == TotalCount;

        public string Summary
        {
            get
            {
                var name = Kind == MediaAssociationKind.Video
                    ? "视频"
                    : Kind == MediaAssociationKind.Audio ? "音频" : "音视频";
                if (IsFullyAssociated)
                {
                    return $"{name}：{AssociatedCount}/{TotalCount} 已默认由 MyTools 打开";
                }

                if (OpenWithCount == TotalCount)
                {
                    return $"{name}：{AssociatedCount}/{TotalCount} 默认打开，MyTools 已在打开方式列表";
                }

                return $"{name}：{AssociatedCount}/{TotalCount} 默认打开，{OpenWithCount}/{TotalCount} 已注册到打开方式";
            }
        }
    }
}
