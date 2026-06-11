using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyTools.Services
{
    public static class CodexRelayTestService
    {
        private const int TimeoutSeconds = 15;

        public static CodexRelayRuntimeInfo InspectRuntime(byte[] configTomlBytes, byte[] authJsonBytes)
        {
            var config = ParseConfig(configTomlBytes);
            return new CodexRelayRuntimeInfo
            {
                BaseUrl = config.BaseUrl ?? string.Empty,
                Model = config.Model ?? string.Empty,
                RequiresOpenAiAuth = config.RequiresOpenAiAuth,
                WireApi = config.WireApi ?? string.Empty,
                Token = ResolveToken(config, authJsonBytes)
            };
        }

        public static async Task<CodexRelayTestResult> TestAsync(byte[] configTomlBytes, byte[] authJsonBytes, CancellationToken ct)
        {
            var config = ParseConfig(configTomlBytes);
            if (string.IsNullOrWhiteSpace(config.BaseUrl))
            {
                return CodexRelayTestResult.Fail("config.toml 未找到 base_url。");
            }

            if (!Uri.TryCreate(config.BaseUrl, UriKind.Absolute, out var baseUri)
                || (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                return CodexRelayTestResult.Fail("base_url 不是有效的 http/https 地址。");
            }

            var token = ResolveToken(config, authJsonBytes);
            if (string.IsNullOrWhiteSpace(token))
            {
                if (config.RequiresOpenAiAuth)
                {
                    var gateway = await TestGatewayReachableAsync(baseUri, ct).ConfigureAwait(false);
                    return gateway.Success
                        ? CodexRelayTestResult.Pass($"可用：{baseUri.Host} 网关可达，当前档案使用 Codex 登录态。", baseUri.Host)
                        : CodexRelayTestResult.Fail(gateway.Message, baseUri.Host);
                }

                if (!AllowsMissingTokenForLocalProvider(baseUri))
                {
                    return CodexRelayTestResult.Fail("未找到可用于测试的 OPENAI_API_KEY、api_key 或 env_key 环境变量。", baseUri.Host);
                }
            }

            EnableTls12();
            using (var handler = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            })
            using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(TimeoutSeconds) })
            {
                if (string.IsNullOrWhiteSpace(config.Model))
                {
                    var models = await SendFirstSuccessfulAsync(client, HttpMethod.Get, BuildEndpoints(baseUri, "models"), token, null, ct).ConfigureAwait(false);
                    return models.Success
                        ? CodexRelayTestResult.Pass(
                            $"可用：{baseUri.Host} {models.EndpointPath} 返回 HTTP {(int)models.StatusCode}。",
                            baseUri.Host,
                            BuildEffectiveBaseUrl(models.RequestUri, "models"))
                        : CodexRelayTestResult.Fail(models.Message, baseUri.Host);
                }

                var responsesBody = new JObject
                {
                    ["model"] = config.Model,
                    ["input"] = "ping",
                    ["max_output_tokens"] = 1,
                    ["stream"] = true
                }.ToString(Formatting.None);
                RelayHttpProbeResult responses = null;
                if (ShouldTryWireApi(config, "responses"))
                {
                    responses = await SendFirstSuccessfulAsync(client, HttpMethod.Post, BuildEndpoints(baseUri, "responses"), token, responsesBody, ct).ConfigureAwait(false);
                    if (responses.Success)
                    {
                        return CodexRelayTestResult.Pass(
                            $"可用：{baseUri.Host} {responses.EndpointPath} 返回 HTTP {(int)responses.StatusCode}。",
                            baseUri.Host,
                            BuildEffectiveBaseUrl(responses.RequestUri, "responses"));
                    }

                    if (responses.IsAuthorizationFailure)
                    {
                        return CodexRelayTestResult.Fail(responses.Message, baseUri.Host);
                    }
                }

                var chatBody = new JObject
                {
                    ["model"] = config.Model,
                    ["messages"] = new JArray(new JObject
                    {
                        ["role"] = "user",
                        ["content"] = "ping"
                    }),
                    ["max_tokens"] = 1,
                    ["stream"] = true
                }.ToString(Formatting.None);
                var chat = await SendFirstSuccessfulAsync(client, HttpMethod.Post, BuildEndpoints(baseUri, "chat/completions"), token, chatBody, ct).ConfigureAwait(false);
                return chat.Success
                    ? CodexRelayTestResult.Pass(
                        $"可用：{baseUri.Host} {chat.EndpointPath} 返回 HTTP {(int)chat.StatusCode}。",
                        baseUri.Host,
                        BuildEffectiveBaseUrl(chat.RequestUri, "chat/completions"))
                    : CodexRelayTestResult.Fail(BuildCombinedFailureMessage(responses, chat), baseUri.Host);
            }
        }

        private static async Task<RelayHttpProbeResult> TestGatewayReachableAsync(Uri baseUri, CancellationToken ct)
        {
            EnableTls12();
            using (var handler = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            })
            using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(TimeoutSeconds) })
            {
                var root = await SendAsync(client, HttpMethod.Get, baseUri, string.Empty, null, ct).ConfigureAwait(false);
                if (root.Success || root.StatusCode == HttpStatusCode.NotFound || root.StatusCode == HttpStatusCode.MethodNotAllowed)
                {
                    root.Success = true;
                    root.EndpointPath = "/";
                    return root;
                }

                return root;
            }
        }

        public static bool AllowsMissingTokenForLocalProvider(string baseUrl)
        {
            return Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri)
                   && AllowsMissingTokenForLocalProvider(uri);
        }

        public static bool AllowsMissingTokenForLocalProvider(Uri uri)
        {
            if (uri == null || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                return false;
            }

            return string.Equals(uri.Host, "localhost", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(uri.Host, "127.0.0.1", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(uri.Host, "::1", StringComparison.OrdinalIgnoreCase)
                   || (IPAddress.TryParse(uri.Host, out var address) && IPAddress.IsLoopback(address));
        }

        private static async Task<RelayHttpProbeResult> SendFirstSuccessfulAsync(
            HttpClient client,
            HttpMethod method,
            IEnumerable<Uri> uris,
            string token,
            string jsonBody,
            CancellationToken ct)
        {
            RelayHttpProbeResult last = null;
            foreach (var uri in uris ?? Enumerable.Empty<Uri>())
            {
                var result = await SendAsync(client, method, uri, token, jsonBody, ct).ConfigureAwait(false);
                if (result.Success)
                {
                    return result;
                }

                last = result;
                if (result.IsAuthorizationFailure)
                {
                    return result;
                }
            }

            return last ?? RelayHttpProbeResult.Fail("请求失败。");
        }

        private static async Task<RelayHttpProbeResult> SendAsync(
            HttpClient client,
            HttpMethod method,
            Uri uri,
            string token,
            string jsonBody,
            CancellationToken ct)
        {
            try
            {
                using (var request = new HttpRequestMessage(method, uri))
                {
                    request.Headers.UserAgent.ParseAdd("MyTools-CodexRelayTest/1.0");
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    if (!string.IsNullOrWhiteSpace(token))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", NormalizeBearerToken(token));
                    }
                    if (!string.IsNullOrWhiteSpace(jsonBody))
                    {
                        request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                    }

                    using (var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        var status = response.StatusCode;
                        return new RelayHttpProbeResult
                        {
                            Success = response.IsSuccessStatusCode,
                            StatusCode = status,
                            EndpointPath = uri.AbsolutePath,
                            RequestUri = uri,
                            Message = $"HTTP {(int)status} {response.ReasonPhrase}"
                        };
                    }
                }
            }
            catch (TaskCanceledException) when (!ct.IsCancellationRequested)
            {
                return RelayHttpProbeResult.Fail("请求超时。");
            }
            catch (Exception ex)
            {
                return RelayHttpProbeResult.Fail(ex.Message);
            }
        }

        private static string BuildCombinedFailureMessage(params RelayHttpProbeResult[] results)
        {
            var messages = (results ?? new RelayHttpProbeResult[0])
                .Where(result => result != null && !string.IsNullOrWhiteSpace(result.Message))
                .Select(result => result.Message)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(3)
                .ToList();
            return messages.Count == 0
                ? "中转测试失败。"
                : "中转测试失败：" + string.Join("；", messages);
        }

        private static CodexRelayConfig ParseConfig(byte[] configTomlBytes)
        {
            var text = configTomlBytes == null ? string.Empty : Encoding.UTF8.GetString(configTomlBytes);
            var modelProvider = string.Empty;
            var model = string.Empty;
            var table = string.Empty;
            var baseUrls = new List<ConfigValue>();
            var envKeys = new List<ConfigValue>();
            var apiKeys = new List<ConfigValue>();
            var requiresOpenAiAuth = false;
            var wireApi = string.Empty;

            foreach (var raw in SplitLines(text))
            {
                var line = StripTomlComment(raw).Trim();
                if (line.Length == 0)
                {
                    continue;
                }

                if (line[0] == '\ufeff')
                {
                    line = line.Substring(1).TrimStart();
                }

                if (line.StartsWith("[", StringComparison.Ordinal) && line.EndsWith("]", StringComparison.Ordinal))
                {
                    table = NormalizeTableName(line.Substring(1, line.Length - 2));
                    continue;
                }

                if (!TryReadAssignment(line, out var key, out var value))
                {
                    continue;
                }

                if (string.Equals(key, "model_provider", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(modelProvider))
                {
                    modelProvider = value;
                }
                else if (string.Equals(key, "model", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(model))
                {
                    model = value;
                }
                else if (string.Equals(key, "base_url", StringComparison.OrdinalIgnoreCase))
                {
                    baseUrls.Add(new ConfigValue(table, value));
                }
                else if (string.Equals(key, "env_key", StringComparison.OrdinalIgnoreCase))
                {
                    envKeys.Add(new ConfigValue(table, value));
                }
                else if (string.Equals(key, "api_key", StringComparison.OrdinalIgnoreCase))
                {
                    apiKeys.Add(new ConfigValue(table, value));
                }
                else if (string.Equals(key, "requires_openai_auth", StringComparison.OrdinalIgnoreCase))
                {
                    requiresOpenAiAuth = IsTomlTrue(value);
                }
                else if (string.Equals(key, "wire_api", StringComparison.OrdinalIgnoreCase))
                {
                    wireApi = value;
                }
            }

            var selectedBase = SelectProviderValue(baseUrls, modelProvider) ?? baseUrls.FirstOrDefault();
            var selectedTable = selectedBase?.Table ?? string.Empty;
            var selectedEnvKey = SelectSameTableValue(envKeys, selectedTable)
                ?? SelectProviderValue(envKeys, modelProvider)
                ?? envKeys.FirstOrDefault();
            var selectedApiKey = SelectSameTableValue(apiKeys, selectedTable)
                ?? SelectProviderValue(apiKeys, modelProvider)
                ?? apiKeys.FirstOrDefault();

            return new CodexRelayConfig
            {
                BaseUrl = selectedBase?.Value ?? string.Empty,
                Model = model,
                EnvKey = selectedEnvKey?.Value ?? string.Empty,
                ApiKey = selectedApiKey?.Value ?? string.Empty,
                RequiresOpenAiAuth = requiresOpenAiAuth,
                WireApi = wireApi
            };
        }

        private static string ResolveToken(CodexRelayConfig config, byte[] authJsonBytes)
        {
            var configToken = ResolveConfigToken(config);
            if (!string.IsNullOrWhiteSpace(configToken))
            {
                return configToken;
            }

            if (authJsonBytes == null || authJsonBytes.Length == 0)
            {
                return string.Empty;
            }

            try
            {
                var root = JObject.Parse(Encoding.UTF8.GetString(authJsonBytes));
                return SelectJsonString(root,
                    "OPENAI_API_KEY",
                    "openai_api_key",
                    "tokens.access_token",
                    "tokens.accessToken",
                    "access_token",
                    "accessToken",
                    "api_key",
                    "apiKey");
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool ShouldTryWireApi(CodexRelayConfig config, string api)
        {
            return string.IsNullOrWhiteSpace(config?.WireApi)
                || string.Equals(config.WireApi, api, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTomlTrue(string value)
        {
            return string.Equals((value ?? string.Empty).Trim(), "true", StringComparison.OrdinalIgnoreCase);
        }

        private static string ResolveConfigToken(CodexRelayConfig config)
        {
            var apiKey = config?.ApiKey ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                var trimmed = apiKey.Trim();
                if (trimmed.StartsWith("env:", StringComparison.OrdinalIgnoreCase))
                {
                    return Environment.GetEnvironmentVariable(trimmed.Substring(4).Trim()) ?? string.Empty;
                }

                if (!trimmed.StartsWith("$", StringComparison.Ordinal) && !trimmed.Contains("{"))
                {
                    return trimmed;
                }
            }

            return string.IsNullOrWhiteSpace(config?.EnvKey)
                ? string.Empty
                : Environment.GetEnvironmentVariable(config.EnvKey.Trim()) ?? string.Empty;
        }

        private static string SelectJsonString(JToken root, params string[] paths)
        {
            foreach (var path in paths ?? new string[0])
            {
                var value = root.SelectToken(path)?.Value<string>();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return string.Empty;
        }

        private static ConfigValue SelectProviderValue(IEnumerable<ConfigValue> values, string provider)
        {
            if (string.IsNullOrWhiteSpace(provider))
            {
                return null;
            }

            var suffix = "model_providers." + NormalizeTableName(provider);
            return values.FirstOrDefault(value =>
                !string.IsNullOrWhiteSpace(value?.Table)
                && value.Table.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
        }

        private static ConfigValue SelectSameTableValue(IEnumerable<ConfigValue> values, string table)
        {
            return values.FirstOrDefault(value =>
                string.Equals(value?.Table ?? string.Empty, table ?? string.Empty, StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryReadAssignment(string line, out string key, out string value)
        {
            key = string.Empty;
            value = string.Empty;
            var index = line.IndexOf('=');
            if (index <= 0)
            {
                return false;
            }

            key = line.Substring(0, index).Trim();
            value = UnquoteTomlValue(line.Substring(index + 1).Trim());
            return key.Length > 0;
        }

        private static string StripTomlComment(string line)
        {
            var value = line ?? string.Empty;
            var inSingle = false;
            var inDouble = false;
            var escaped = false;
            for (var i = 0; i < value.Length; i++)
            {
                var ch = value[i];
                if (escaped)
                {
                    escaped = false;
                    continue;
                }

                if (inDouble && ch == '\\')
                {
                    escaped = true;
                    continue;
                }

                if (ch == '\'' && !inDouble)
                {
                    inSingle = !inSingle;
                    continue;
                }

                if (ch == '"' && !inSingle)
                {
                    inDouble = !inDouble;
                    continue;
                }

                if (ch == '#' && !inSingle && !inDouble)
                {
                    return value.Substring(0, i);
                }
            }

            return value;
        }

        private static string UnquoteTomlValue(string value)
        {
            var trimmed = (value ?? string.Empty).Trim();
            if (trimmed.Length >= 2
                && ((trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"')
                    || (trimmed[0] == '\'' && trimmed[trimmed.Length - 1] == '\'')))
            {
                trimmed = trimmed.Substring(1, trimmed.Length - 2);
            }

            return trimmed.Replace("\\\"", "\"").Trim();
        }

        private static string NormalizeTableName(string value)
        {
            return (value ?? string.Empty).Trim().Trim('"').Trim('\'').Replace("\"", string.Empty).Replace("'", string.Empty);
        }

        private static IEnumerable<string> SplitLines(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        }

        private static IEnumerable<Uri> BuildEndpoints(Uri baseUri, string relativePath)
        {
            yield return BuildEndpoint(baseUri, relativePath);

            var path = (baseUri.AbsolutePath ?? string.Empty).Trim('/');
            if (!path.Equals("v1", StringComparison.OrdinalIgnoreCase)
                && !path.StartsWith("v1/", StringComparison.OrdinalIgnoreCase))
            {
                yield return BuildEndpoint(baseUri, "v1/" + (relativePath ?? string.Empty).TrimStart('/'));
            }
        }

        private static Uri BuildEndpoint(Uri baseUri, string relativePath)
        {
            var builder = new UriBuilder(baseUri);
            var path = builder.Path ?? string.Empty;
            if (!path.EndsWith("/", StringComparison.Ordinal))
            {
                path += "/";
            }

            builder.Path = path + (relativePath ?? string.Empty).TrimStart('/');
            builder.Query = string.Empty;
            return builder.Uri;
        }

        private static string BuildEffectiveBaseUrl(Uri requestUri, string apiRelativePath)
        {
            if (requestUri == null)
            {
                return string.Empty;
            }

            var apiPath = "/" + (apiRelativePath ?? string.Empty).Trim('/');
            var path = requestUri.AbsolutePath ?? string.Empty;
            var basePath = path;
            if (!string.IsNullOrWhiteSpace(apiPath)
                && path.EndsWith(apiPath, StringComparison.OrdinalIgnoreCase))
            {
                basePath = path.Substring(0, path.Length - apiPath.Length);
            }

            if (string.IsNullOrWhiteSpace(basePath))
            {
                basePath = "/";
            }

            var builder = new UriBuilder(requestUri)
            {
                Path = basePath.TrimEnd('/'),
                Query = string.Empty,
                Fragment = string.Empty
            };

            return builder.Uri.ToString().TrimEnd('/');
        }

        private static string NormalizeBearerToken(string token)
        {
            var value = (token ?? string.Empty).Trim();
            return value.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase)
                ? value.Substring("Bearer ".Length).Trim()
                : value;
        }

        private static void EnableTls12()
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch
            {
                // Older Windows policy can reject changing the protocol flags; HttpClient will still use OS defaults.
            }
        }

        private sealed class CodexRelayConfig
        {
            public string BaseUrl { get; set; }
            public string Model { get; set; }
            public string EnvKey { get; set; }
            public string ApiKey { get; set; }
            public bool RequiresOpenAiAuth { get; set; }
            public string WireApi { get; set; }
        }

        private sealed class ConfigValue
        {
            public ConfigValue(string table, string value)
            {
                Table = table ?? string.Empty;
                Value = value ?? string.Empty;
            }

            public string Table { get; }
            public string Value { get; }
        }

        private sealed class RelayHttpProbeResult
        {
            public bool Success { get; set; }
            public HttpStatusCode StatusCode { get; set; }
            public string EndpointPath { get; set; }
            public Uri RequestUri { get; set; }
            public string Message { get; set; }

            public bool IsAuthorizationFailure =>
                StatusCode == HttpStatusCode.Unauthorized || StatusCode == HttpStatusCode.Forbidden;

            public static RelayHttpProbeResult Fail(string message)
            {
                return new RelayHttpProbeResult
                {
                    Success = false,
                    Message = string.IsNullOrWhiteSpace(message) ? "请求失败。" : message
                };
            }
        }
    }

    public sealed class CodexRelayTestResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string Host { get; set; }
        public string EffectiveBaseUrl { get; set; }

        public static CodexRelayTestResult Pass(string message, string host, string effectiveBaseUrl = "")
        {
            return new CodexRelayTestResult
            {
                Success = true,
                Message = message ?? string.Empty,
                Host = host ?? string.Empty,
                EffectiveBaseUrl = effectiveBaseUrl ?? string.Empty
            };
        }

        public static CodexRelayTestResult Fail(string message, string host = "")
        {
            return new CodexRelayTestResult
            {
                Success = false,
                Message = message ?? string.Empty,
                Host = host ?? string.Empty
            };
        }
    }

    public sealed class CodexRelayRuntimeInfo
    {
        public string BaseUrl { get; set; }
        public string Model { get; set; }
        public bool RequiresOpenAiAuth { get; set; }
        public string WireApi { get; set; }
        public string Token { get; set; }

        public bool HasToken => !string.IsNullOrWhiteSpace(Token);
    }
}
