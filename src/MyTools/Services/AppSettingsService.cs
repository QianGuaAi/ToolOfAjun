using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public static class AppSettingsService
    {
        private static readonly string SettingsPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "MyTools.settings.json");

        public static async Task<AppSettings> LoadAsync()
        {
            try
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
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading app settings failed.");
                return new AppSettings();
            }
        }

        public static async Task SaveAsync(AppSettings settings)
        {
            if (settings == null) return;
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? AppDomain.CurrentDomain.BaseDirectory);
                using (var stream = new FileStream(SettingsPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                    await writer.WriteAsync(json).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Saving app settings failed.");
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
        public bool ShowEditorAfterCapture { get; set; } = true;
    }

    public class HotkeySettings
    {
        public uint Modifiers { get; set; }
        public uint Key { get; set; }
        public string DisplayText { get; set; } = "Ctrl+Shift+Z";
    }
}
