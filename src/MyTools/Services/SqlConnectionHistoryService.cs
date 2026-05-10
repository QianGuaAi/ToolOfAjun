using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public static class SqlConnectionHistoryService
    {
        private const int MaxHistoryItems = 12;
        private static readonly string SettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyTools.sqlhistory.json");

        public static Task<SqlConnectionHistory> LoadAsync()
        {
            return LoadAsync(SqlProviderKind.SqlServer);
        }

        public static async Task<SqlConnectionHistory> LoadAsync(SqlProviderKind providerKind)
        {
            try
            {
                var storage = await LoadStorageAsync().ConfigureAwait(false);
                MigrateLegacySqlServerData(storage);
                return BuildHistory(GetOrCreateProviderData(storage, providerKind));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading SQL connection history failed.");
                return new SqlConnectionHistory();
            }
        }

        public static async Task SaveAsync(SqlServerConnectionOptions options)
        {
            if (options == null)
            {
                return;
            }

            var providerKind = options.ProviderKind;
            var storage = await LoadStorageAsync().ConfigureAwait(false);
            MigrateLegacySqlServerData(storage);

            var providerData = GetOrCreateProviderData(storage, providerKind);
            var current = BuildHistory(providerData);
            storage.Providers[providerKind.ToString()] = new SqlConnectionHistoryData
            {
                LastServerAddress = options.ServerAddress?.Trim(),
                LastPort = options.Port?.Trim(),
                LastUsername = options.Username?.Trim(),
                LastPassword = Protect(options.Password),
                ServerAddresses = MoveToTop(current.ServerAddresses, options.ServerAddress),
                Usernames = MoveToTop(current.Usernames, options.Username),
                Passwords = MoveToTop(current.Passwords, options.Password).Select(Protect).ToList()
            };

            await SaveStorageAsync(storage).ConfigureAwait(false);
        }

        private static async Task<SqlConnectionHistoryStorage> LoadStorageAsync()
        {
            if (!File.Exists(SettingsPath))
            {
                return new SqlConnectionHistoryStorage();
            }

            using (var stream = new FileStream(SettingsPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                var json = await reader.ReadToEndAsync().ConfigureAwait(false);
                return JsonConvert.DeserializeObject<SqlConnectionHistoryStorage>(json) ?? new SqlConnectionHistoryStorage();
            }
        }

        private static async Task SaveStorageAsync(SqlConnectionHistoryStorage storage)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? AppDomain.CurrentDomain.BaseDirectory);
            using (var stream = new FileStream(SettingsPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                var json = JsonConvert.SerializeObject(storage, Formatting.Indented);
                await writer.WriteAsync(json).ConfigureAwait(false);
            }
        }

        private static void MigrateLegacySqlServerData(SqlConnectionHistoryStorage storage)
        {
            if (storage == null)
            {
                return;
            }

            if (storage.Providers == null)
            {
                storage.Providers = new Dictionary<string, SqlConnectionHistoryData>(StringComparer.OrdinalIgnoreCase);
            }

            var legacyHasData =
                !string.IsNullOrWhiteSpace(storage.LastServerAddress)
                || !string.IsNullOrWhiteSpace(storage.LastUsername)
                || !string.IsNullOrWhiteSpace(storage.LastPassword)
                || (storage.ServerAddresses?.Count ?? 0) > 0
                || (storage.Usernames?.Count ?? 0) > 0
                || (storage.Passwords?.Count ?? 0) > 0;
            if (!legacyHasData)
            {
                return;
            }

            var key = SqlProviderKind.SqlServer.ToString();
            if (!storage.Providers.ContainsKey(key))
            {
                storage.Providers[key] = new SqlConnectionHistoryData
                {
                    LastServerAddress = storage.LastServerAddress,
                    LastPort = storage.LastPort,
                    LastUsername = storage.LastUsername,
                    LastPassword = storage.LastPassword,
                    ServerAddresses = storage.ServerAddresses ?? new List<string>(),
                    Usernames = storage.Usernames ?? new List<string>(),
                    Passwords = storage.Passwords ?? new List<string>()
                };
            }
        }

        private static SqlConnectionHistoryData GetOrCreateProviderData(SqlConnectionHistoryStorage storage, SqlProviderKind providerKind)
        {
            if (storage.Providers == null)
            {
                storage.Providers = new Dictionary<string, SqlConnectionHistoryData>(StringComparer.OrdinalIgnoreCase);
            }

            var key = providerKind.ToString();
            if (!storage.Providers.TryGetValue(key, out var data) || data == null)
            {
                data = new SqlConnectionHistoryData();
                storage.Providers[key] = data;
            }

            data.ServerAddresses = data.ServerAddresses ?? new List<string>();
            data.Usernames = data.Usernames ?? new List<string>();
            data.Passwords = data.Passwords ?? new List<string>();
            return data;
        }

        private static SqlConnectionHistory BuildHistory(SqlConnectionHistoryData data)
        {
            return new SqlConnectionHistory
            {
                LastServerAddress = data?.LastServerAddress,
                LastPort = data?.LastPort,
                LastUsername = data?.LastUsername,
                LastPassword = Unprotect(data?.LastPassword),
                ServerAddresses = Clean(data?.ServerAddresses),
                Usernames = Clean(data?.Usernames),
                Passwords = Clean((data?.Passwords ?? new List<string>()).Select(Unprotect))
            };
        }

        private static List<string> MoveToTop(IEnumerable<string> values, string value)
        {
            var result = new List<string>();
            AddIfValid(result, value);

            foreach (var item in values ?? Enumerable.Empty<string>())
            {
                AddIfValid(result, item);
                if (result.Count >= MaxHistoryItems)
                {
                    break;
                }
            }

            return result;
        }

        private static List<string> Clean(IEnumerable<string> values)
        {
            var result = new List<string>();
            foreach (var value in values ?? Enumerable.Empty<string>())
            {
                AddIfValid(result, value);
                if (result.Count >= MaxHistoryItems)
                {
                    break;
                }
            }

            return result;
        }

        private static void AddIfValid(ICollection<string> values, string value)
        {
            var normalized = value?.Trim();
            if (string.IsNullOrWhiteSpace(normalized) || values.Any(item => string.Equals(item, normalized, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            values.Add(normalized);
        }

        private static string Protect(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var bytes = Encoding.UTF8.GetBytes(value);
            return Convert.ToBase64String(ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser));
        }

        private static string Unprotect(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            try
            {
                var bytes = Convert.FromBase64String(value);
                return Encoding.UTF8.GetString(ProtectedData.Unprotect(bytes, null, DataProtectionScope.CurrentUser));
            }
            catch
            {
                return string.Empty;
            }
        }

        private sealed class SqlConnectionHistoryStorage
        {
            public Dictionary<string, SqlConnectionHistoryData> Providers { get; set; } =
                new Dictionary<string, SqlConnectionHistoryData>(StringComparer.OrdinalIgnoreCase);

            // Legacy single-provider fields for backward compatibility.
            public string LastServerAddress { get; set; }
            public string LastPort { get; set; }
            public string LastUsername { get; set; }
            public string LastPassword { get; set; }
            public List<string> ServerAddresses { get; set; }
            public List<string> Usernames { get; set; }
            public List<string> Passwords { get; set; }
        }

        private sealed class SqlConnectionHistoryData
        {
            public string LastServerAddress { get; set; }
            public string LastPort { get; set; }
            public string LastUsername { get; set; }
            public string LastPassword { get; set; }
            public List<string> ServerAddresses { get; set; } = new List<string>();
            public List<string> Usernames { get; set; } = new List<string>();
            public List<string> Passwords { get; set; } = new List<string>();
        }
    }

    public class SqlConnectionHistory
    {
        public string LastServerAddress { get; set; }
        public string LastPort { get; set; }
        public string LastUsername { get; set; }
        public string LastPassword { get; set; }
        public List<string> ServerAddresses { get; set; } = new List<string>();
        public List<string> Usernames { get; set; } = new List<string>();
        public List<string> Passwords { get; set; } = new List<string>();
    }
}
