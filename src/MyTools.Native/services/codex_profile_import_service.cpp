#include "services/codex_profile_import_service.h"

#include <cctype>
#include <cstdio>
#include <sstream>
#include <vector>

#include <windows.h>

#include "services/file_system.h"
#include "services/secret_store_dpapi.h"

namespace mytools {
namespace {

std::string WideToUtf8(const std::wstring& text) {
    if (text.empty()) {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8,
                                             WC_ERR_INVALID_CHARS,
                                             text.data(),
                                             static_cast<int>(text.size()),
                                             nullptr,
                                             0,
                                             nullptr,
                                             nullptr);
    if (required <= 0) {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    WideCharToMultiByte(CP_UTF8,
                        WC_ERR_INVALID_CHARS,
                        text.data(),
                        static_cast<int>(text.size()),
                        result.data(),
                        required,
                        nullptr,
                        nullptr);
    return result;
}

std::string JsonEscape(const std::string& text) {
    std::ostringstream escaped;
    for (const unsigned char ch : text) {
        switch (ch) {
            case '"':
                escaped << "\\\"";
                break;
            case '\\':
                escaped << "\\\\";
                break;
            case '\b':
                escaped << "\\b";
                break;
            case '\f':
                escaped << "\\f";
                break;
            case '\n':
                escaped << "\\n";
                break;
            case '\r':
                escaped << "\\r";
                break;
            case '\t':
                escaped << "\\t";
                break;
            default:
                if (ch < 0x20) {
                    const char digits[] = "0123456789ABCDEF";
                    escaped << "\\u00" << digits[(ch >> 4) & 0x0F] << digits[ch & 0x0F];
                } else {
                    escaped << static_cast<char>(ch);
                }
                break;
        }
    }
    return escaped.str();
}

std::wstring TimestampUtcIsoWide() {
    SYSTEMTIME utc{};
    GetSystemTime(&utc);

    wchar_t buffer[40]{};
    swprintf_s(buffer,
               L"%04u-%02u-%02uT%02u:%02u:%02uZ",
               utc.wYear,
               utc.wMonth,
               utc.wDay,
               utc.wHour,
               utc.wMinute,
               utc.wSecond);
    return buffer;
}

std::wstring MachineName() {
    wchar_t buffer[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD size = MAX_COMPUTERNAME_LENGTH + 1;
    if (!GetComputerNameW(buffer, &size)) {
        return L"";
    }
    return std::wstring(buffer, size);
}

void ClearString(std::string* value) {
    if (value != nullptr && !value->empty()) {
        SecureZeroMemory(value->data(), value->size());
    }
}

void ClearBytes(std::vector<unsigned char>* bytes) {
    if (bytes != nullptr && !bytes->empty()) {
        SecureZeroMemory(bytes->data(), bytes->size());
    }
}

std::string BytesToString(const std::vector<unsigned char>& bytes) {
    if (bytes.empty()) {
        return {};
    }
    return std::string(reinterpret_cast<const char*>(bytes.data()),
                       reinterpret_cast<const char*>(bytes.data()) + bytes.size());
}

std::string BuildProfileItemJson(const std::wstring& display_name,
                                 const std::wstring& codex_home,
                                 const std::string& protected_config,
                                 const std::string& protected_auth) {
    const std::string name = JsonEscape(WideToUtf8(display_name));
    const std::string folder = JsonEscape(WideToUtf8(codex_home));
    const std::string imported_at = JsonEscape(WideToUtf8(TimestampUtcIsoWide()));

    std::ostringstream json;
    json << "    {\n"
         << "      \"DisplayName\": \"" << name << "\",\n"
         << "      \"Name\": \"" << name << "\",\n"
         << "      \"FolderPath\": \"" << folder << "\",\n"
         << "      \"Note\": \"\",\n"
         << "      \"Remark\": \"" << name << "\",\n"
         << "      \"Tags\": \"\",\n"
         << "      \"Status\": \"\\u672A\\u77E5\",\n"
         << "      \"LastImportedAt\": \"" << imported_at << "\",\n"
         << "      \"ProtectedConfigTomlBase64\": \"" << JsonEscape(protected_config) << "\",\n"
         << "      \"ProtectedAuthJsonBase64\": \"" << JsonEscape(protected_auth) << "\",\n"
         << "      \"ConfigTomlContentProtected\": \"" << JsonEscape(protected_config) << "\",\n"
         << "      \"AuthJsonContentProtected\": \"" << JsonEscape(protected_auth) << "\"\n"
         << "    }";
    return json.str();
}

std::string BuildNewProfilesJson(const std::string& profile_item_json) {
    std::ostringstream json;
    json << "{\n"
         << "  \"schemaVersion\": 2,\n"
         << "  \"machineName\": \"" << JsonEscape(WideToUtf8(MachineName())) << "\",\n"
         << "  \"createdAtUtc\": \"" << JsonEscape(WideToUtf8(TimestampUtcIsoWide())) << "\",\n"
         << "  \"items\": [\n"
         << profile_item_json << "\n"
         << "  ]\n"
         << "}\n";
    return json.str();
}

bool FindItemsArrayBounds(const std::string& json, size_t* array_start, size_t* array_end) {
    if (array_start == nullptr || array_end == nullptr) {
        return false;
    }

    const std::string property = "\"items\"";
    const size_t found = json.find(property);
    if (found == std::string::npos) {
        return false;
    }

    size_t index = found + property.size();
    while (index < json.size() && std::isspace(static_cast<unsigned char>(json[index])) != 0) {
        ++index;
    }
    if (index >= json.size() || json[index] != ':') {
        return false;
    }
    ++index;
    while (index < json.size() && std::isspace(static_cast<unsigned char>(json[index])) != 0) {
        ++index;
    }
    if (index >= json.size() || json[index] != '[') {
        return false;
    }

    *array_start = index;
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
        } else if (ch == '[') {
            ++depth;
        } else if (ch == ']') {
            if (depth == 0) {
                return false;
            }
            --depth;
            if (depth == 0) {
                *array_end = index;
                return true;
            }
        }
        ++index;
    }
    return false;
}

bool ItemsArrayHasContent(const std::string& json, size_t array_start, size_t array_end) {
    if (array_start >= array_end || array_end > json.size()) {
        return false;
    }
    for (size_t index = array_start + 1; index < array_end; ++index) {
        if (std::isspace(static_cast<unsigned char>(json[index])) == 0) {
            return true;
        }
    }
    return false;
}

std::string AppendProfileItemJson(const std::string& existing_json,
                                  const std::string& profile_item_json,
                                  bool* created_new_library) {
    if (created_new_library != nullptr) {
        *created_new_library = false;
    }

    size_t array_start = 0;
    size_t array_end = 0;
    if (existing_json.empty() || !FindItemsArrayBounds(existing_json, &array_start, &array_end)) {
        if (created_new_library != nullptr) {
            *created_new_library = true;
        }
        return BuildNewProfilesJson(profile_item_json);
    }

    std::string json = existing_json.substr(0, array_end);
    if (ItemsArrayHasContent(existing_json, array_start, array_end)) {
        json += ",\n";
    } else {
        json += "\n";
    }
    json += profile_item_json;
    json += existing_json.substr(array_end);
    return json;
}

bool LoadExistingProfilesJson(const std::wstring& profiles_json_path,
                              std::string* json,
                              std::wstring* error_message) {
    if (json == nullptr) {
        return false;
    }
    if (!FileSystem::Exists(profiles_json_path)) {
        json->clear();
        return true;
    }

    std::string protected_text;
    if (!FileSystem::ReadUtf8File(profiles_json_path, &protected_text, error_message)) {
        return false;
    }
    if (protected_text.empty()) {
        json->clear();
        return true;
    }

    SecretStoreDpapi secret_store;
    if (!secret_store.UnprotectBase64ToUtf8(protected_text, json, error_message)) {
        ClearString(&protected_text);
        return false;
    }
    ClearString(&protected_text);
    return true;
}

}  // namespace

bool CodexProfileImportService::ImportCurrentFolderProfile(
    const CodexProfileImportRequest& request,
    CodexProfileImportResult* result) const {
    if (result == nullptr) {
        return false;
    }
    *result = CodexProfileImportResult{};

    if (request.display_name.empty()) {
        result->error_message = L"Codex profile display name is required.";
        return false;
    }

    const std::wstring config_path = FileSystem::JoinPath(request.codex_home, L"config.toml");
    const std::wstring auth_path = FileSystem::JoinPath(request.codex_home, L"auth.json");
    if (!FileSystem::Exists(config_path) || !FileSystem::Exists(auth_path)) {
        result->error_message = L"Current Codex folder must contain config.toml and auth.json.";
        return false;
    }

    std::vector<unsigned char> config_bytes;
    std::vector<unsigned char> auth_bytes;
    if (!FileSystem::ReadFileBytes(config_path, &config_bytes, &result->error_message) ||
        !FileSystem::ReadFileBytes(auth_path, &auth_bytes, &result->error_message)) {
        ClearBytes(&config_bytes);
        ClearBytes(&auth_bytes);
        return false;
    }

    std::string config_text = BytesToString(config_bytes);
    std::string auth_text = BytesToString(auth_bytes);
    ClearBytes(&config_bytes);
    ClearBytes(&auth_bytes);

    SecretStoreDpapi secret_store;
    std::string protected_config;
    std::string protected_auth;
    if (!secret_store.ProtectUtf8ToBase64(config_text, &protected_config, &result->error_message) ||
        !secret_store.ProtectUtf8ToBase64(auth_text, &protected_auth, &result->error_message)) {
        ClearString(&config_text);
        ClearString(&auth_text);
        ClearString(&protected_config);
        ClearString(&protected_auth);
        return false;
    }
    ClearString(&config_text);
    ClearString(&auth_text);

    std::string existing_json;
    if (!LoadExistingProfilesJson(request.profiles_json_path, &existing_json, &result->error_message)) {
        ClearString(&protected_config);
        ClearString(&protected_auth);
        return false;
    }

    std::string item_json =
        BuildProfileItemJson(request.display_name, request.codex_home, protected_config, protected_auth);
    ClearString(&protected_config);
    ClearString(&protected_auth);

    bool created_new = false;
    std::string next_json = AppendProfileItemJson(existing_json, item_json, &created_new);
    ClearString(&item_json);
    ClearString(&existing_json);
    result->created_new_library = created_new;

    std::string protected_profiles;
    if (!secret_store.ProtectUtf8ToBase64(next_json, &protected_profiles, &result->error_message)) {
        ClearString(&next_json);
        return false;
    }
    ClearString(&next_json);

    if (!FileSystem::WriteUtf8FileAtomic(request.profiles_json_path,
                                         protected_profiles,
                                         &result->error_message)) {
        ClearString(&protected_profiles);
        return false;
    }

    ClearString(&protected_profiles);
    return true;
}

}  // namespace mytools
