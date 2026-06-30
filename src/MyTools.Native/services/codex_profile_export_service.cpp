#include "services/codex_profile_export_service.h"

#include <cctype>
#include <utility>
#include <vector>

#include <windows.h>

#include "services/file_system.h"
#include "services/secret_store_dpapi.h"

namespace mytools {
namespace {

std::wstring Utf8ToWide(const std::string& text) {
    if (text.empty()) {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8,
                                             MB_ERR_INVALID_CHARS,
                                             text.data(),
                                             static_cast<int>(text.size()),
                                             nullptr,
                                             0);
    if (required <= 0) {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    MultiByteToWideChar(CP_UTF8,
                        MB_ERR_INVALID_CHARS,
                        text.data(),
                        static_cast<int>(text.size()),
                        result.data(),
                        required);
    return result;
}

int HexValue(char value) {
    if (value >= '0' && value <= '9') {
        return value - '0';
    }
    if (value >= 'a' && value <= 'f') {
        return 10 + value - 'a';
    }
    if (value >= 'A' && value <= 'F') {
        return 10 + value - 'A';
    }
    return -1;
}

bool ReadHexCodeUnit(const std::string& json, size_t* index, uint32_t* code_unit) {
    if (index == nullptr || code_unit == nullptr || *index + 4 > json.size()) {
        return false;
    }

    uint32_t value = 0;
    for (size_t offset = 0; offset < 4; ++offset) {
        const int digit = HexValue(json[*index + offset]);
        if (digit < 0) {
            return false;
        }
        value = (value << 4) | static_cast<uint32_t>(digit);
    }

    *index += 4;
    *code_unit = value;
    return true;
}

void AppendUtf8CodePoint(uint32_t code_point, std::string* value) {
    if (value == nullptr) {
        return;
    }

    if (code_point <= 0x7F) {
        value->push_back(static_cast<char>(code_point));
    } else if (code_point <= 0x7FF) {
        value->push_back(static_cast<char>(0xC0 | (code_point >> 6)));
        value->push_back(static_cast<char>(0x80 | (code_point & 0x3F)));
    } else if (code_point <= 0xFFFF) {
        value->push_back(static_cast<char>(0xE0 | (code_point >> 12)));
        value->push_back(static_cast<char>(0x80 | ((code_point >> 6) & 0x3F)));
        value->push_back(static_cast<char>(0x80 | (code_point & 0x3F)));
    } else if (code_point <= 0x10FFFF) {
        value->push_back(static_cast<char>(0xF0 | (code_point >> 18)));
        value->push_back(static_cast<char>(0x80 | ((code_point >> 12) & 0x3F)));
        value->push_back(static_cast<char>(0x80 | ((code_point >> 6) & 0x3F)));
        value->push_back(static_cast<char>(0x80 | (code_point & 0x3F)));
    }
}

bool ReadJsonString(const std::string& json, size_t* index, std::string* value) {
    if (index == nullptr || value == nullptr || *index >= json.size() || json[*index] != '"') {
        return false;
    }

    ++(*index);
    value->clear();
    while (*index < json.size()) {
        const char ch = json[*index];
        ++(*index);
        if (ch == '"') {
            return true;
        }
        if (ch != '\\') {
            value->push_back(ch);
            continue;
        }
        if (*index >= json.size()) {
            return false;
        }

        const char escaped = json[*index];
        ++(*index);
        switch (escaped) {
            case '"':
            case '\\':
            case '/':
                value->push_back(escaped);
                break;
            case 'b':
                value->push_back('\b');
                break;
            case 'f':
                value->push_back('\f');
                break;
            case 'n':
                value->push_back('\n');
                break;
            case 'r':
                value->push_back('\r');
                break;
            case 't':
                value->push_back('\t');
                break;
            case 'u': {
                uint32_t code_unit = 0;
                if (!ReadHexCodeUnit(json, index, &code_unit)) {
                    return false;
                }
                if (code_unit >= 0xD800 && code_unit <= 0xDBFF) {
                    if (*index + 6 > json.size() || json[*index] != '\\' ||
                        json[*index + 1] != 'u') {
                        return false;
                    }
                    *index += 2;
                    uint32_t low_surrogate = 0;
                    if (!ReadHexCodeUnit(json, index, &low_surrogate) ||
                        low_surrogate < 0xDC00 || low_surrogate > 0xDFFF) {
                        return false;
                    }
                    const uint32_t code_point =
                        0x10000 + (((code_unit - 0xD800) << 10) | (low_surrogate - 0xDC00));
                    AppendUtf8CodePoint(code_point, value);
                } else if (code_unit >= 0xDC00 && code_unit <= 0xDFFF) {
                    return false;
                } else {
                    AppendUtf8CodePoint(code_unit, value);
                }
                break;
            }
            default:
                return false;
        }
    }
    return false;
}

void SkipWhitespace(const std::string& json, size_t* index) {
    while (index != nullptr && *index < json.size() &&
           std::isspace(static_cast<unsigned char>(json[*index])) != 0) {
        ++(*index);
    }
}

std::vector<std::string> ExtractJsonStringValues(const std::string& json, const char* property_name) {
    std::vector<std::string> values;
    const std::string property = std::string("\"") + property_name + "\"";

    size_t search = 0;
    while (search < json.size()) {
        const size_t found = json.find(property, search);
        if (found == std::string::npos) {
            break;
        }

        size_t index = found + property.size();
        SkipWhitespace(json, &index);
        if (index >= json.size() || json[index] != ':') {
            search = found + property.size();
            continue;
        }
        ++index;
        SkipWhitespace(json, &index);
        if (index >= json.size() || json[index] != '"') {
            search = index;
            continue;
        }

        std::string value;
        if (ReadJsonString(json, &index, &value)) {
            values.push_back(std::move(value));
        }
        search = index;
    }

    return values;
}

std::string ExtractFirstJsonStringValue(const std::string& json, const char* property_name) {
    const std::vector<std::string> values = ExtractJsonStringValues(json, property_name);
    return values.empty() ? std::string() : values.front();
}

bool FindArrayAfterProperty(const std::string& json, const char* property_name, size_t* array_start) {
    if (property_name == nullptr || array_start == nullptr) {
        return false;
    }

    const std::string property = std::string("\"") + property_name + "\"";
    size_t search = 0;
    while (search < json.size()) {
        const size_t found = json.find(property, search);
        if (found == std::string::npos) {
            return false;
        }

        size_t index = found + property.size();
        SkipWhitespace(json, &index);
        if (index < json.size() && json[index] == ':') {
            ++index;
            SkipWhitespace(json, &index);
            if (index < json.size() && json[index] == '[') {
                *array_start = index;
                return true;
            }
        }
        search = found + property.size();
    }
    return false;
}

std::vector<std::string> ExtractJsonObjectArray(const std::string& json, const char* property_name) {
    std::vector<std::string> objects;

    size_t index = 0;
    if (!FindArrayAfterProperty(json, property_name, &index)) {
        return objects;
    }

    ++index;
    while (index < json.size()) {
        SkipWhitespace(json, &index);
        if (index >= json.size() || json[index] == ']') {
            break;
        }
        if (json[index] != '{') {
            ++index;
            continue;
        }

        const size_t object_start = index;
        size_t depth = 0;
        bool in_string = false;
        bool escaped = false;
        while (index < json.size()) {
            const char ch = json[index];
            if (in_string) {
                if (escaped) {
                    escaped = false;
                } else if (ch == '\\') {
                    escaped = true;
                } else if (ch == '"') {
                    in_string = false;
                }
                ++index;
                continue;
            }

            if (ch == '"') {
                in_string = true;
            } else if (ch == '{') {
                ++depth;
            } else if (ch == '}') {
                if (depth == 0) {
                    break;
                }
                --depth;
                if (depth == 0) {
                    ++index;
                    objects.push_back(json.substr(object_start, index - object_start));
                    break;
                }
            }
            ++index;
        }
    }

    return objects;
}

bool EqualsOrdinalIgnoreCase(const std::wstring& left, const std::wstring& right) {
    if (left.empty() || right.empty()) {
        return false;
    }
    return CompareStringOrdinal(left.c_str(), -1, right.c_str(), -1, TRUE) == CSTR_EQUAL;
}

std::string ProtectedField(const std::string& profile_object,
                           const char* primary_name,
                           const char* legacy_name) {
    std::string value = ExtractFirstJsonStringValue(profile_object, primary_name);
    if (value.empty()) {
        value = ExtractFirstJsonStringValue(profile_object, legacy_name);
    }
    return value;
}

void ClearString(std::string* value) {
    if (value != nullptr && !value->empty()) {
        SecureZeroMemory(value->data(), value->size());
    }
}

bool LoadTargetProfileObject(const std::wstring& profiles_json_path,
                             const std::wstring& target_display_name,
                             std::string* profile_object,
                             std::wstring* error_message) {
    if (profile_object == nullptr) {
        if (error_message != nullptr) {
            *error_message = L"LoadTargetProfileObject requires an output buffer.";
        }
        return false;
    }

    std::string protected_text;
    if (!FileSystem::ReadUtf8File(profiles_json_path, &protected_text, error_message)) {
        return false;
    }
    if (protected_text.empty()) {
        if (error_message != nullptr) {
            *error_message = L"Codex profiles.json is missing or empty.";
        }
        return false;
    }

    SecretStoreDpapi secret_store;
    std::string json;
    if (!secret_store.UnprotectBase64ToUtf8(protected_text, &json, error_message)) {
        ClearString(&protected_text);
        return false;
    }
    ClearString(&protected_text);

    std::vector<std::string> objects = ExtractJsonObjectArray(json, "items");
    for (const std::string& object : objects) {
        std::string display_name = ExtractFirstJsonStringValue(object, "DisplayName");
        if (display_name.empty()) {
            display_name = ExtractFirstJsonStringValue(object, "Name");
        }

        if (EqualsOrdinalIgnoreCase(Utf8ToWide(display_name), target_display_name)) {
            *profile_object = object;
            break;
        }
    }

    for (std::string& object : objects) {
        ClearString(&object);
    }
    ClearString(&json);

    if (profile_object->empty()) {
        if (error_message != nullptr) {
            *error_message = L"Target Codex profile was not found in profiles.json.";
        }
        return false;
    }
    return true;
}

}  // namespace

bool CodexProfileExportService::ExportProfileByDisplayName(
    const CodexProfileExportRequest& request,
    CodexProfileExportResult* result) const {
    if (result == nullptr) {
        return false;
    }
    *result = CodexProfileExportResult{};

    if (request.target_display_name.empty()) {
        result->error_message = L"Target Codex profile display name is required.";
        return false;
    }
    if (request.output_directory.empty()) {
        result->error_message = L"Codex profile export output directory is required.";
        return false;
    }

    std::string profile_object;
    if (!LoadTargetProfileObject(request.profiles_json_path,
                                 request.target_display_name,
                                 &profile_object,
                                 &result->error_message)) {
        return false;
    }

    std::string protected_config =
        ProtectedField(profile_object, "ProtectedConfigTomlBase64", "ConfigTomlContentProtected");
    std::string protected_auth =
        ProtectedField(profile_object, "ProtectedAuthJsonBase64", "AuthJsonContentProtected");
    ClearString(&profile_object);

    if (protected_config.empty() || protected_auth.empty()) {
        ClearString(&protected_config);
        ClearString(&protected_auth);
        result->error_message = L"Target Codex profile does not contain exportable config/auth content.";
        return false;
    }

    SecretStoreDpapi secret_store;
    std::string config_toml;
    std::string auth_json;
    if (!secret_store.UnprotectBase64ToUtf8(protected_config, &config_toml, &result->error_message) ||
        !secret_store.UnprotectBase64ToUtf8(protected_auth, &auth_json, &result->error_message)) {
        ClearString(&protected_config);
        ClearString(&protected_auth);
        ClearString(&config_toml);
        ClearString(&auth_json);
        return false;
    }
    ClearString(&protected_config);
    ClearString(&protected_auth);

    result->config_path = FileSystem::JoinPath(request.output_directory, L"config.toml");
    result->auth_path = FileSystem::JoinPath(request.output_directory, L"auth.json");
    if (!FileSystem::WriteUtf8FileAtomic(result->config_path, config_toml, &result->error_message) ||
        !FileSystem::WriteUtf8FileAtomic(result->auth_path, auth_json, &result->error_message)) {
        ClearString(&config_toml);
        ClearString(&auth_json);
        return false;
    }

    ClearString(&config_toml);
    ClearString(&auth_json);
    return true;
}

}  // namespace mytools
