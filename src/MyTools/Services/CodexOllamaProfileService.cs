using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace MyTools.Services
{
    public static class CodexOllamaProfileService
    {
        public const string ProviderId = "mytools_local_ollama";
        public const string BaseUrl = "http://127.0.0.1:11434/v1";

        private const string TagsEndpoint = "http://127.0.0.1:11434/api/tags";
        private const int TimeoutSeconds = 8;

        public static async Task<List<CodexOllamaProfileDefinition>> LoadInstalledProfilesAsync(CancellationToken ct)
        {
            var models = await LoadInstalledModelsAsync(ct).ConfigureAwait(false);
            return models
                .Select(model => new CodexOllamaProfileDefinition
                {
                    ModelName = model.Name,
                    DisplayName = BuildDisplayName(model.Name),
                    ConfigTomlBytes = Encoding.UTF8.GetBytes(BuildConfigToml(model.Name)),
                    AuthJsonBytes = Encoding.UTF8.GetBytes("{}")
                })
                .ToList();
        }

        private static async Task<List<OllamaModelInfo>> LoadInstalledModelsAsync(CancellationToken ct)
        {
            EnableTls12();
            using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(TimeoutSeconds) })
            using (var request = new HttpRequestMessage(HttpMethod.Get, TagsEndpoint))
            {
                request.Headers.UserAgent.ParseAdd("MyTools-CodexOllamaImport/1.0");
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                using (var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException("Ollama 模型列表读取失败：HTTP " + ((int)response.StatusCode).ToString() + " " + response.ReasonPhrase);
                    }

                    var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var root = JObject.Parse(json);
                    var rawModels = (root["models"] as JArray ?? new JArray())
                        .OfType<JObject>()
                        .Select(item => new OllamaModelInfo
                        {
                            Name = item.Value<string>("name") ?? string.Empty,
                            Digest = item.Value<string>("digest") ?? string.Empty
                        })
                        .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                        .ToList();

                    var importModels = PreferFriendlyNames(rawModels);
                    return importModels
                        .GroupBy(item => string.IsNullOrWhiteSpace(item.Digest) ? item.Name.Trim() : item.Digest.Trim(), StringComparer.OrdinalIgnoreCase)
                        .Select(group => group
                            .OrderBy(item => IsRawHuggingFaceName(item.Name))
                            .ThenBy(item => item.Name.Length)
                            .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                            .First())
                        .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }
        }

        private static string BuildConfigToml(string modelName)
        {
            var builder = new StringBuilder();
            builder.AppendLine("model = " + TomlQuote(modelName));
            builder.AppendLine("model_provider = " + TomlQuote(ProviderId));
            builder.AppendLine();
            builder.AppendLine("[model_providers." + ProviderId + "]");
            builder.AppendLine("name = \"Ollama (local)\"");
            builder.AppendLine("base_url = " + TomlQuote(BaseUrl));
            builder.AppendLine("wire_api = \"responses\"");
            return builder.ToString();
        }

        private static string BuildDisplayName(string modelName)
        {
            var name = (modelName ?? string.Empty).Trim();
            if (name.EndsWith(":latest", StringComparison.OrdinalIgnoreCase))
            {
                name = name.Substring(0, name.Length - ":latest".Length);
            }

            return "Ollama - " + name;
        }

        private static bool IsRawHuggingFaceName(string modelName)
        {
            return (modelName ?? string.Empty).Trim().StartsWith("hf.co/", StringComparison.OrdinalIgnoreCase);
        }

        private static List<OllamaModelInfo> PreferFriendlyNames(List<OllamaModelInfo> models)
        {
            var values = models ?? new List<OllamaModelInfo>();
            var friendly = values
                .Where(item => item != null && !IsRawHuggingFaceName(item.Name))
                .ToList();
            if (friendly.Count == 0)
            {
                return values;
            }

            var friendlyKeys = new HashSet<string>(
                friendly.SelectMany(item => BuildModelMatchKeys(item.Name)),
                StringComparer.OrdinalIgnoreCase);
            return values
                .Where(item => item != null
                               && (!IsRawHuggingFaceName(item.Name)
                                   || !BuildModelMatchKeys(item.Name).Any(friendlyKeys.Contains)))
                .ToList();
        }

        private static IEnumerable<string> BuildModelMatchKeys(string modelName)
        {
            var name = (modelName ?? string.Empty).Trim();
            if (name.Length == 0)
            {
                yield break;
            }

            var lower = name.ToLowerInvariant();
            if (lower.Contains("neuraldaredevil"))
            {
                yield return "neuraldaredevil";
            }

            if (lower.Contains("qwen3.6") || lower.Contains("qwen36"))
            {
                yield return "qwen3.6";
            }

            if (lower.Contains("dolphin3"))
            {
                yield return "dolphin3";
            }

            yield return NormalizeModelMatchKey(lower);
        }

        private static string NormalizeModelMatchKey(string value)
        {
            var builder = new StringBuilder();
            foreach (var ch in value ?? string.Empty)
            {
                if (char.IsLetterOrDigit(ch))
                {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private static string TomlQuote(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        private static void EnableTls12()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        }

        private sealed class OllamaModelInfo
        {
            public string Name { get; set; }
            public string Digest { get; set; }
        }
    }

    public sealed class CodexOllamaProfileDefinition
    {
        public string ModelName { get; set; }
        public string DisplayName { get; set; }
        public byte[] ConfigTomlBytes { get; set; }
        public byte[] AuthJsonBytes { get; set; }
    }
}
