using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyTools.Services
{
    internal enum CodexLocalRelayEndpoint
    {
        Unknown = 0,
        Models = 1,
        ChatCompletions = 2
    }

    internal static class CodexLocalRelayCompatibility
    {
        public static CodexLocalRelayEndpoint ResolveEndpoint(string localBasePath, string requestPath)
        {
            var relativePath = StripLocalBasePath(localBasePath, requestPath);
            if (string.Equals(relativePath, "models", StringComparison.OrdinalIgnoreCase))
            {
                return CodexLocalRelayEndpoint.Models;
            }

            if (string.Equals(relativePath, "chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                return CodexLocalRelayEndpoint.ChatCompletions;
            }

            return CodexLocalRelayEndpoint.Unknown;
        }

        public static bool CanUseModelsFallback(System.Net.HttpStatusCode statusCode)
        {
            return statusCode == System.Net.HttpStatusCode.BadRequest
                   || statusCode == System.Net.HttpStatusCode.NotFound
                   || statusCode == System.Net.HttpStatusCode.MethodNotAllowed
                   || statusCode == System.Net.HttpStatusCode.NotImplemented;
        }

        public static bool CanRetryChatAsResponses(System.Net.HttpStatusCode statusCode)
        {
            return statusCode == System.Net.HttpStatusCode.BadRequest
                   || statusCode == System.Net.HttpStatusCode.NotFound
                   || statusCode == System.Net.HttpStatusCode.MethodNotAllowed
                   || statusCode == System.Net.HttpStatusCode.NotImplemented
                   || statusCode == System.Net.HttpStatusCode.UnsupportedMediaType;
        }

        public static bool IsOpenAiCompatibleContentType(string contentType)
        {
            var value = (contentType ?? string.Empty).Trim();
            return value.IndexOf("application/json", StringComparison.OrdinalIgnoreCase) >= 0
                   || value.IndexOf("text/event-stream", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static string BuildModelsResponseJson(string model)
        {
            var data = new JArray();
            var modelId = string.IsNullOrWhiteSpace(model) ? string.Empty : model.Trim();
            if (!string.IsNullOrWhiteSpace(modelId))
            {
                data.Add(new JObject
                {
                    ["id"] = modelId,
                    ["object"] = "model",
                    ["created"] = 0,
                    ["owned_by"] = "mytools-local-relay"
                });
            }

            return new JObject
            {
                ["object"] = "list",
                ["data"] = data
            }.ToString(Formatting.None);
        }

        public static bool TryBuildResponsesRequestBody(
            byte[] chatBody,
            string fallbackModel,
            out byte[] responsesBody,
            out bool clientRequestedStream)
        {
            responsesBody = null;
            clientRequestedStream = false;

            JObject chat;
            try
            {
                chat = JObject.Parse(Encoding.UTF8.GetString(chatBody ?? new byte[0]));
            }
            catch
            {
                return false;
            }

            clientRequestedStream = chat.Value<bool?>("stream") == true;
            var responses = new JObject();
            CopyStringOrFallback(chat, responses, "model", fallbackModel);
            responses["input"] = BuildResponsesInput(chat["messages"] as JArray);
            responses["stream"] = false;

            CopyToken(chat, responses, "temperature");
            CopyToken(chat, responses, "top_p");
            CopyToken(chat, responses, "metadata");
            CopyToken(chat, responses, "user");
            CopyToken(chat, responses, "parallel_tool_calls");
            CopyToken(chat, responses, "response_format");
            CopyToken(chat, responses, "seed");

            var maxOutputTokens = chat["max_output_tokens"] ?? chat["max_completion_tokens"] ?? chat["max_tokens"];
            if (maxOutputTokens != null)
            {
                responses["max_output_tokens"] = maxOutputTokens.DeepClone();
            }

            var tools = BuildResponsesTools(chat["tools"] as JArray);
            if (tools != null && tools.Count > 0)
            {
                responses["tools"] = tools;
            }

            var toolChoice = BuildResponsesToolChoice(chat["tool_choice"]);
            if (toolChoice != null)
            {
                responses["tool_choice"] = toolChoice;
            }

            responsesBody = Encoding.UTF8.GetBytes(responses.ToString(Formatting.None));
            return true;
        }

        public static bool TryBuildChatCompletionJson(
            string responsesJson,
            byte[] originalChatBody,
            out string chatJson)
        {
            chatJson = null;

            JObject responses;
            JObject chat;
            try
            {
                responses = JObject.Parse(responsesJson ?? string.Empty);
                chat = JObject.Parse(Encoding.UTF8.GetString(originalChatBody ?? new byte[0]));
            }
            catch
            {
                return false;
            }

            var extracted = ExtractResponsesOutput(responses);
            var usage = BuildChatUsage(responses["usage"]);
            var message = new JObject
            {
                ["role"] = "assistant",
                ["content"] = extracted.ToolCalls.Count > 0 && string.IsNullOrEmpty(extracted.Content)
                    ? JValue.CreateNull()
                    : new JValue(extracted.Content)
            };

            if (extracted.ToolCalls.Count > 0)
            {
                message["tool_calls"] = extracted.ToolCalls;
            }

            var completion = new JObject
            {
                ["id"] = BuildChatId(responses.Value<string>("id")),
                ["object"] = "chat.completion",
                ["created"] = ResolveCreated(responses),
                ["model"] = ResolveModel(responses, chat),
                ["choices"] = new JArray(new JObject
                {
                    ["index"] = 0,
                    ["message"] = message,
                    ["finish_reason"] = extracted.FinishReason
                })
            };

            if (usage != null)
            {
                completion["usage"] = usage;
            }

            chatJson = completion.ToString(Formatting.None);
            return true;
        }

        public static bool TryBuildChatCompletionSsePayload(
            string responsesJson,
            byte[] originalChatBody,
            out string ssePayload)
        {
            ssePayload = null;

            JObject responses;
            JObject chat;
            try
            {
                responses = JObject.Parse(responsesJson ?? string.Empty);
                chat = JObject.Parse(Encoding.UTF8.GetString(originalChatBody ?? new byte[0]));
            }
            catch
            {
                return false;
            }

            var extracted = ExtractResponsesOutput(responses);
            var id = BuildChatId(responses.Value<string>("id"));
            var created = ResolveCreated(responses);
            var model = ResolveModel(responses, chat);
            var builder = new StringBuilder();

            AppendSseData(builder, BuildChatChunk(id, created, model, new JObject { ["role"] = "assistant" }, null));
            if (extracted.ToolCalls.Count > 0)
            {
                AppendSseData(builder, BuildChatChunk(id, created, model, new JObject { ["tool_calls"] = extracted.ToolCalls }, null));
            }
            else if (!string.IsNullOrEmpty(extracted.Content))
            {
                AppendSseData(builder, BuildChatChunk(id, created, model, new JObject { ["content"] = extracted.Content }, null));
            }

            AppendSseData(builder, BuildChatChunk(id, created, model, new JObject(), extracted.FinishReason));
            builder.Append("data: [DONE]\n\n");
            ssePayload = builder.ToString();
            return true;
        }

        private static string StripLocalBasePath(string localBasePath, string requestPath)
        {
            var local = NormalizePath(localBasePath);
            var request = NormalizePath(requestPath);
            if (request.Equals(local, StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            if (request.StartsWith(local + "/", StringComparison.OrdinalIgnoreCase))
            {
                request = request.Substring(local.Length);
            }

            return request.Trim('/');
        }

        private static string NormalizePath(string value)
        {
            var path = (value ?? string.Empty).Trim();
            if (path.Length == 0)
            {
                return "/";
            }

            if (!path.StartsWith("/", StringComparison.Ordinal))
            {
                path = "/" + path;
            }

            return path.TrimEnd('/');
        }

        private static JArray BuildResponsesInput(JArray messages)
        {
            var input = new JArray();
            foreach (var messageToken in messages ?? new JArray())
            {
                var message = messageToken as JObject;
                if (message == null)
                {
                    continue;
                }

                var role = (message.Value<string>("role") ?? "user").Trim();
                if (string.Equals(role, "tool", StringComparison.OrdinalIgnoreCase))
                {
                    input.Add(new JObject
                    {
                        ["type"] = "function_call_output",
                        ["call_id"] = message.Value<string>("tool_call_id") ?? string.Empty,
                        ["output"] = TokenToText(message["content"])
                    });
                    continue;
                }

                var toolCalls = message["tool_calls"] as JArray;
                if (toolCalls != null && toolCalls.Count > 0)
                {
                    var contentText = TokenToText(message["content"]);
                    if (!string.IsNullOrWhiteSpace(contentText))
                    {
                        input.Add(new JObject
                        {
                            ["role"] = NormalizeMessageRole(role),
                            ["content"] = contentText
                        });
                    }

                    foreach (var toolCall in toolCalls.OfType<JObject>())
                    {
                        var function = toolCall["function"] as JObject;
                        input.Add(new JObject
                        {
                            ["type"] = "function_call",
                            ["call_id"] = toolCall.Value<string>("id") ?? string.Empty,
                            ["name"] = function?.Value<string>("name") ?? string.Empty,
                            ["arguments"] = TokenToText(function?["arguments"])
                        });
                    }
                    continue;
                }

                input.Add(new JObject
                {
                    ["role"] = NormalizeMessageRole(role),
                    ["content"] = BuildResponsesContent(message["content"])
                });
            }

            return input;
        }

        private static JToken BuildResponsesContent(JToken content)
        {
            if (content == null || content.Type == JTokenType.Null)
            {
                return string.Empty;
            }

            if (content.Type == JTokenType.String)
            {
                return content.DeepClone();
            }

            var array = content as JArray;
            if (array == null)
            {
                return content.DeepClone();
            }

            var converted = new JArray();
            foreach (var partToken in array)
            {
                var part = partToken as JObject;
                if (part == null)
                {
                    converted.Add(partToken.DeepClone());
                    continue;
                }

                var type = part.Value<string>("type") ?? string.Empty;
                if (string.Equals(type, "text", StringComparison.OrdinalIgnoreCase))
                {
                    converted.Add(new JObject
                    {
                        ["type"] = "input_text",
                        ["text"] = part.Value<string>("text") ?? string.Empty
                    });
                    continue;
                }

                if (string.Equals(type, "image_url", StringComparison.OrdinalIgnoreCase))
                {
                    converted.Add(new JObject
                    {
                        ["type"] = "input_image",
                        ["image_url"] = part["image_url"]?["url"] ?? part["image_url"] ?? string.Empty
                    });
                    continue;
                }

                converted.Add(part.DeepClone());
            }

            return converted;
        }

        private static JArray BuildResponsesTools(JArray chatTools)
        {
            if (chatTools == null)
            {
                return null;
            }

            var tools = new JArray();
            foreach (var toolToken in chatTools)
            {
                var tool = toolToken as JObject;
                if (tool == null)
                {
                    continue;
                }

                var type = tool.Value<string>("type") ?? string.Empty;
                var function = tool["function"] as JObject;
                if (string.Equals(type, "function", StringComparison.OrdinalIgnoreCase) && function != null)
                {
                    var converted = new JObject
                    {
                        ["type"] = "function",
                        ["name"] = function.Value<string>("name") ?? string.Empty,
                        ["description"] = function.Value<string>("description") ?? string.Empty,
                        ["parameters"] = function["parameters"]?.DeepClone() ?? new JObject()
                    };
                    CopyToken(function, converted, "strict");
                    tools.Add(converted);
                    continue;
                }

                tools.Add(tool.DeepClone());
            }

            return tools;
        }

        private static JToken BuildResponsesToolChoice(JToken chatToolChoice)
        {
            if (chatToolChoice == null || chatToolChoice.Type == JTokenType.Null)
            {
                return null;
            }

            var choice = chatToolChoice as JObject;
            var function = choice?["function"] as JObject;
            if (function != null)
            {
                return new JObject
                {
                    ["type"] = "function",
                    ["name"] = function.Value<string>("name") ?? string.Empty
                };
            }

            return chatToolChoice.DeepClone();
        }

        private static ResponsesOutput ExtractResponsesOutput(JObject responses)
        {
            var content = new StringBuilder();
            var toolCalls = new JArray();
            var outputText = responses.Value<string>("output_text");
            if (!string.IsNullOrEmpty(outputText))
            {
                content.Append(outputText);
            }

            var output = responses["output"] as JArray;
            foreach (var item in output ?? new JArray())
            {
                var obj = item as JObject;
                if (obj == null)
                {
                    continue;
                }

                var type = obj.Value<string>("type") ?? string.Empty;
                if (string.Equals(type, "message", StringComparison.OrdinalIgnoreCase))
                {
                    AppendMessageContent(content, obj["content"] as JArray);
                    continue;
                }

                if (string.Equals(type, "function_call", StringComparison.OrdinalIgnoreCase))
                {
                    toolCalls.Add(BuildChatToolCall(obj, toolCalls.Count));
                }
            }

            return new ResponsesOutput
            {
                Content = content.ToString(),
                ToolCalls = toolCalls,
                FinishReason = toolCalls.Count > 0 ? "tool_calls" : "stop"
            };
        }

        private static void AppendMessageContent(StringBuilder content, JArray parts)
        {
            foreach (var part in parts ?? new JArray())
            {
                var obj = part as JObject;
                if (obj == null)
                {
                    continue;
                }

                var text = obj.Value<string>("text");
                if (!string.IsNullOrEmpty(text))
                {
                    content.Append(text);
                }
            }
        }

        private static JObject BuildChatToolCall(JObject functionCall, int index)
        {
            var callId = functionCall.Value<string>("call_id")
                         ?? functionCall.Value<string>("id")
                         ?? ("call_" + index.ToString(CultureInfo.InvariantCulture));
            return new JObject
            {
                ["id"] = callId,
                ["type"] = "function",
                ["function"] = new JObject
                {
                    ["name"] = functionCall.Value<string>("name") ?? string.Empty,
                    ["arguments"] = TokenToText(functionCall["arguments"])
                }
            };
        }

        private static JObject BuildChatUsage(JToken usageToken)
        {
            var usage = usageToken as JObject;
            if (usage == null)
            {
                return null;
            }

            return new JObject
            {
                ["prompt_tokens"] = usage["prompt_tokens"] ?? usage["input_tokens"] ?? 0,
                ["completion_tokens"] = usage["completion_tokens"] ?? usage["output_tokens"] ?? 0,
                ["total_tokens"] = usage["total_tokens"] ?? 0
            };
        }

        private static JObject BuildChatChunk(string id, long created, string model, JObject delta, string finishReason)
        {
            return new JObject
            {
                ["id"] = id,
                ["object"] = "chat.completion.chunk",
                ["created"] = created,
                ["model"] = model,
                ["choices"] = new JArray(new JObject
                {
                    ["index"] = 0,
                    ["delta"] = delta,
                    ["finish_reason"] = finishReason == null ? JValue.CreateNull() : new JValue(finishReason)
                })
            };
        }

        private static void AppendSseData(StringBuilder builder, JObject payload)
        {
            builder.Append("data: ")
                .Append(payload.ToString(Formatting.None))
                .Append("\n\n");
        }

        private static void CopyStringOrFallback(JObject source, JObject target, string name, string fallback)
        {
            var value = source.Value<string>(name);
            target[name] = string.IsNullOrWhiteSpace(value) ? (fallback ?? string.Empty) : value.Trim();
        }

        private static void CopyToken(JObject source, JObject target, string name)
        {
            var value = source?[name];
            if (value != null)
            {
                target[name] = value.DeepClone();
            }
        }

        private static string NormalizeMessageRole(string role)
        {
            var value = string.IsNullOrWhiteSpace(role) ? "user" : role.Trim().ToLowerInvariant();
            if (value == "system" || value == "developer" || value == "assistant" || value == "user")
            {
                return value;
            }

            return "user";
        }

        private static string TokenToText(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }

            return token.Type == JTokenType.String
                ? token.Value<string>() ?? string.Empty
                : token.ToString(Formatting.None);
        }

        private static string BuildChatId(string responseId)
        {
            var value = string.IsNullOrWhiteSpace(responseId) ? Guid.NewGuid().ToString("N") : responseId.Trim();
            return value.StartsWith("chatcmpl-", StringComparison.OrdinalIgnoreCase)
                ? value
                : "chatcmpl-" + value;
        }

        private static long ResolveCreated(JObject responses)
        {
            var created = responses.Value<long?>("created") ?? responses.Value<long?>("created_at");
            return created ?? DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        }

        private static string ResolveModel(JObject responses, JObject chat)
        {
            return responses.Value<string>("model")
                   ?? chat.Value<string>("model")
                   ?? string.Empty;
        }

        private sealed class ResponsesOutput
        {
            public string Content { get; set; }
            public JArray ToolCalls { get; set; }
            public string FinishReason { get; set; }
        }
    }
}
