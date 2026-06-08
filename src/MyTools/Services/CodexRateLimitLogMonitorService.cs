using System;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace MyTools.Services
{
    public static class CodexRateLimitLogMonitorService
    {
        private const int MaxRowsPerProbe = 50;

        public static string LogsDatabasePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".codex",
            "logs_2.sqlite");

        public static Task<CodexRateLimitProbeResult> InitializeBaselineAsync(CancellationToken ct)
        {
            return Task.Run(() =>
            {
                try
                {
                    if (!File.Exists(LogsDatabasePath))
                    {
                        return CodexRateLimitProbeResult.NotAvailable("未找到 Codex 日志库。");
                    }

                    using (var connection = OpenReadOnlyConnection())
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT IFNULL(MAX(id), 0) FROM logs";
                        var value = command.ExecuteScalar();
                        var lastId = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                        return new CodexRateLimitProbeResult
                        {
                            Success = true,
                            LastSeenLogId = lastId,
                            Message = "已开始监控 Codex 429。"
                        };
                    }
                }
                catch (Exception ex)
                {
                    return CodexRateLimitProbeResult.NotAvailable("Codex 429 监控未就绪：" + ex.GetType().Name);
                }
            }, ct);
        }

        public static Task<CodexRateLimitProbeResult> ProbeAsync(long lastSeenLogId, CancellationToken ct)
        {
            return Task.Run(() =>
            {
                try
                {
                    if (!File.Exists(LogsDatabasePath))
                    {
                        return CodexRateLimitProbeResult.NotAvailable("未找到 Codex 日志库。");
                    }

                    using (var connection = OpenReadOnlyConnection())
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = @"
SELECT id, ts, ts_nanos, target, feedback_log_body
FROM logs
WHERE id > @lastSeenId
ORDER BY id ASC
LIMIT @maxRows";
                        command.Parameters.AddWithValue("@lastSeenId", lastSeenLogId);
                        command.Parameters.AddWithValue("@maxRows", MaxRowsPerProbe);

                        var result = new CodexRateLimitProbeResult
                        {
                            Success = true,
                            LastSeenLogId = lastSeenLogId,
                            Message = "未检测到新的 Codex 429。"
                        };

                        using (var reader = command.ExecuteReader(CommandBehavior.SequentialAccess))
                        {
                            while (reader.Read())
                            {
                                ct.ThrowIfCancellationRequested();
                                var id = reader.GetInt64(0);
                                result.LastSeenLogId = Math.Max(result.LastSeenLogId, id);

                                var target = reader.IsDBNull(3) ? string.Empty : reader.GetString(3);
                                if (!LooksLikeRateLimitTarget(target))
                                {
                                    continue;
                                }

                                var body = reader.IsDBNull(4) ? string.Empty : reader.GetString(4);
                                var detected = TryReadRateLimitHit(body, out var resetAfter);
                                if (!detected)
                                {
                                    continue;
                                }

                                result.Detected = true;
                                result.LogId = id;
                                result.ObservedAtUtc = TryConvertLogTimestamp(reader.GetInt64(1));
                                result.ResetAfter = resetAfter;
                                result.Message = resetAfter.HasValue
                                    ? "检测到 Codex 429，预计 " + FormatResetAfter(resetAfter.Value) + " 后恢复。"
                                    : "检测到 Codex 429。";
                            }
                        }

                        return result;
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    return CodexRateLimitProbeResult.NotAvailable("Codex 429 监控读取失败：" + ex.GetType().Name);
                }
            }, ct);
        }

        private static SQLiteConnection OpenReadOnlyConnection()
        {
            ConfigureSQLiteBaseDirectory();
            var builder = new SQLiteConnectionStringBuilder
            {
                DataSource = LogsDatabasePath,
                ReadOnly = true,
                FailIfMissing = true
            };
            var connection = new SQLiteConnection(builder.ConnectionString);
            connection.Open();
            return connection;
        }

        private static void ConfigureSQLiteBaseDirectory()
        {
            if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("SQLite_ConfigureDirectory")))
            {
                var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                if (!string.IsNullOrWhiteSpace(baseDirectory))
                {
                    Environment.SetEnvironmentVariable("SQLite_ConfigureDirectory", baseDirectory);
                }
            }

            if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("SQLite_NoPlugins")))
            {
                Environment.SetEnvironmentVariable("SQLite_NoPlugins", "1");
            }
        }

        private static bool LooksLikeRateLimitTarget(string target)
        {
            return string.Equals(target, "codex.rate_limits", StringComparison.OrdinalIgnoreCase)
                || (target ?? string.Empty).IndexOf("rate_limit", StringComparison.OrdinalIgnoreCase) >= 0
                || (target ?? string.Empty).IndexOf("rate-limit", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool TryReadRateLimitHit(string body, out TimeSpan? resetAfter)
        {
            resetAfter = null;
            if (string.IsNullOrWhiteSpace(body))
            {
                return false;
            }

            var jsonStart = body.IndexOf('{');
            if (jsonStart < 0)
            {
                return ContainsLimitReachedText(body);
            }

            try
            {
                var root = JObject.Parse(body.Substring(jsonStart));
                resetAfter = FindResetAfter(root);
                return ContainsBoolean(root, "allowed", false)
                    || ContainsBoolean(root, "limit_reached", true)
                    || ContainsBoolean(root, "limitReached", true);
            }
            catch
            {
                return ContainsLimitReachedText(body);
            }
        }

        private static bool ContainsLimitReachedText(string body)
        {
            var text = body ?? string.Empty;
            return text.IndexOf("\"allowed\":false", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("\"limit_reached\":true", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("\"limitReached\":true", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool ContainsBoolean(JToken token, string name, bool expected)
        {
            if (token == null)
            {
                return false;
            }

            if (token.Type == JTokenType.Object)
            {
                foreach (var property in ((JObject)token).Properties())
                {
                    if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)
                        && property.Value.Type == JTokenType.Boolean
                        && property.Value.Value<bool>() == expected)
                    {
                        return true;
                    }

                    if (ContainsBoolean(property.Value, name, expected))
                    {
                        return true;
                    }
                }
            }
            else if (token.Type == JTokenType.Array)
            {
                foreach (var child in token.Children())
                {
                    if (ContainsBoolean(child, name, expected))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static TimeSpan? FindResetAfter(JToken token)
        {
            var seconds = FindNumber(token, "reset_after_seconds") ?? FindNumber(token, "resetAfterSeconds");
            if (!seconds.HasValue || seconds.Value < 0)
            {
                return null;
            }

            return TimeSpan.FromSeconds(seconds.Value);
        }

        private static double? FindNumber(JToken token, string name)
        {
            if (token == null)
            {
                return null;
            }

            if (token.Type == JTokenType.Object)
            {
                foreach (var property in ((JObject)token).Properties())
                {
                    if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)
                        && (property.Value.Type == JTokenType.Integer || property.Value.Type == JTokenType.Float))
                    {
                        return property.Value.Value<double>();
                    }

                    var nested = FindNumber(property.Value, name);
                    if (nested.HasValue)
                    {
                        return nested;
                    }
                }
            }
            else if (token.Type == JTokenType.Array)
            {
                foreach (var child in token.Children())
                {
                    var nested = FindNumber(child, name);
                    if (nested.HasValue)
                    {
                        return nested;
                    }
                }
            }

            return null;
        }

        private static DateTime? TryConvertLogTimestamp(long value)
        {
            try
            {
                return DateTimeOffset.FromUnixTimeSeconds(value).UtcDateTime;
            }
            catch
            {
                return null;
            }
        }

        private static string FormatResetAfter(TimeSpan value)
        {
            if (value.TotalHours >= 1)
            {
                return Math.Ceiling(value.TotalHours).ToString("0", CultureInfo.InvariantCulture) + " 小时";
            }

            return Math.Max(1, Math.Ceiling(value.TotalMinutes)).ToString("0", CultureInfo.InvariantCulture) + " 分钟";
        }
    }

    public sealed class CodexRateLimitProbeResult
    {
        public bool Success { get; set; }
        public bool Detected { get; set; }
        public long LogId { get; set; }
        public long LastSeenLogId { get; set; }
        public DateTime? ObservedAtUtc { get; set; }
        public TimeSpan? ResetAfter { get; set; }
        public string Message { get; set; }

        public static CodexRateLimitProbeResult NotAvailable(string message)
        {
            return new CodexRateLimitProbeResult
            {
                Success = false,
                Message = message ?? string.Empty
            };
        }
    }
}
