using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public static class AppSettingsService
    {
        private static readonly string SettingsPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "MyTools.settings.json");

        private static readonly SemaphoreSlim SettingsLock = new SemaphoreSlim(1, 1);

        public static async Task<AppSettings> LoadAsync()
        {
            await SettingsLock.WaitAsync().ConfigureAwait(false);
            try
            {
                return await LoadCoreAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading app settings failed.");
                return new AppSettings();
            }
            finally
            {
                SettingsLock.Release();
            }
        }

        public static async Task SaveAsync(AppSettings settings)
        {
            if (settings == null) return;

            await SettingsLock.WaitAsync().ConfigureAwait(false);
            try
            {
                await SaveCoreAsync(settings).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Saving app settings failed.");
            }
            finally
            {
                SettingsLock.Release();
            }
        }

        public static async Task UpdateAsync(Action<AppSettings> update)
        {
            if (update == null) return;

            await SettingsLock.WaitAsync().ConfigureAwait(false);
            try
            {
                var settings = await LoadCoreAsync().ConfigureAwait(false);
                update(settings);
                await SaveCoreAsync(settings).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Updating app settings failed.");
            }
            finally
            {
                SettingsLock.Release();
            }
        }

        private static async Task<AppSettings> LoadCoreAsync()
        {
            if (!File.Exists(SettingsPath))
                return new AppSettings();

            using (var stream = new FileStream(SettingsPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                var json = await reader.ReadToEndAsync().ConfigureAwait(false);
                return JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
            }
        }

        private static async Task SaveCoreAsync(AppSettings settings)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? AppDomain.CurrentDomain.BaseDirectory);
            using (var stream = new FileStream(SettingsPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                await writer.WriteAsync(json).ConfigureAwait(false);
            }
        }
    }

    public class AppSettings
    {
        public HotkeySettings ScreenshotHotkey { get; set; } = new HotkeySettings
        {
            Modifiers = 0x0006,
            Key = 0x5A,
            DisplayText = "Ctrl+Shift+Z"
        };
        public HotkeySettings VideoRecordHotkey { get; set; } = new HotkeySettings
        {
            Modifiers = 0,
            Key = 0,
            DisplayText = "未设置"
        };
        public HotkeySettings AudioRecordHotkey { get; set; } = new HotkeySettings
        {
            Modifiers = 0,
            Key = 0,
            DisplayText = "未设置"
        };
        public bool ShowEditorAfterCapture { get; set; } = true;
        /// <summary>截图模式：FullScreen / Region / Window。</summary>
        public string ScreenshotMode { get; set; } = "FullScreen";
        public List<CodexProfileSettings> CodexProfiles { get; set; } = new List<CodexProfileSettings>();
        public string RecordingOutputFolder { get; set; } = string.Empty;
        public string AudioOutputFolder { get; set; } = string.Empty;
        public List<RecentPlaylistSettings> RecentPlaylists { get; set; } = new List<RecentPlaylistSettings>();
        public List<FavoritePlaylistSettings> FavoritePlaylists { get; set; } = new List<FavoritePlaylistSettings>();
        public List<RecentWeChatBackupSettings> RecentWeChatBackups { get; set; } = new List<RecentWeChatBackupSettings>();
    }

    public class HotkeySettings
    {
        public uint Modifiers { get; set; }
        public uint Key { get; set; }
        public string DisplayText { get; set; } = "Ctrl+Shift+Z";
    }

    public class CodexProfileSettings
    {
        public string Name { get; set; }
        public string Remark { get; set; }
        public string Tags { get; set; }
        public string FolderPath { get; set; }
        public string ConfigTomlContentProtected { get; set; }
        public string AuthJsonContentProtected { get; set; }
        public DateTime? LastAppliedAt { get; set; }
    }

    public class RecentPlaylistSettings
    {
        public string FilePath { get; set; }
        public DateTime LastUsedAt { get; set; }
    }

    public class RecentWeChatBackupSettings
    {
        public string FilePath { get; set; }
        public DateTime LastUsedAt { get; set; }
        public int FileCount { get; set; }
        public long TotalBytes { get; set; }
    }

    public class FavoritePlaylistSettings
    {
        public string Name { get; set; }
        public List<string> FilePaths { get; set; } = new List<string>();
        public DateTime CreatedAt { get; set; }
        public DateTime LastUsedAt { get; set; }
    }
}
