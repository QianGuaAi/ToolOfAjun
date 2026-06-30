#include "services/codex_profile_backup_service.h"

#include <cstdio>
#include <cctype>
#include <sstream>
#include <vector>

#include <windows.h>
#include <wincrypt.h>

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

bool EncodeBytesBase64(const std::vector<unsigned char>& bytes, std::string* output) {
    if (output == nullptr) {
        return false;
    }
    if (bytes.empty()) {
        output->clear();
        return true;
    }

    DWORD required = 0;
    if (!CryptBinaryToStringA(bytes.data(),
                              static_cast<DWORD>(bytes.size()),
                              CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                              nullptr,
                              &required) ||
        required == 0) {
        return false;
    }

    output->assign(required, '\0');
    if (!CryptBinaryToStringA(bytes.data(),
                              static_cast<DWORD>(bytes.size()),
                              CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                              output->data(),
                              &required)) {
        output->clear();
        return false;
    }

    if (required > 0 && !output->empty() && output->back() == '\0') {
        output->resize(required - 1);
    } else {
        output->resize(required);
    }
    return true;
}

bool DecodeBase64(const std::string& text, std::vector<unsigned char>* bytes) {
    if (bytes == nullptr || text.empty()) {
        return false;
    }

    DWORD required = 0;
    if (!CryptStringToBinaryA(text.c_str(),
                              static_cast<DWORD>(text.size()),
                              CRYPT_STRING_BASE64 | CRYPT_STRING_STRICT,
                              nullptr,
                              &required,
                              nullptr,
                              nullptr) ||
        required == 0) {
        return false;
    }

    bytes->assign(required, 0);
    if (!CryptStringToBinaryA(text.c_str(),
                              static_cast<DWORD>(text.size()),
                              CRYPT_STRING_BASE64 | CRYPT_STRING_STRICT,
                              bytes->data(),
                              &required,
                              nullptr,
                              nullptr)) {
        bytes->clear();
        return false;
    }
    bytes->resize(required);
    return true;
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

std::string ExtractFirstJsonStringValue(const std::string& json, const char* property_name) {
    if (property_name == nullptr) {
        return {};
    }

    const std::string property = std::string("\"") + property_name + "\"";
    size_t search = 0;
    while (search < json.size()) {
        const size_t found = json.find(property, search);
        if (found == std::string::npos) {
            return {};
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
            return value;
        }
        search = index;
    }
    return {};
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

bool FindLatestBackupPath(const std::wstring& backup_directory,
                          std::wstring* backup_path,
                          std::wstring* error_message) {
    if (backup_path == nullptr) {
        if (error_message != nullptr) {
            *error_message = L"FindLatestBackupPath requires an output path.";
        }
        return false;
    }

    WIN32_FIND_DATAW data{};
    const std::wstring pattern = FileSystem::JoinPath(backup_directory, L"*.bak.dpapi");
    HANDLE find = FindFirstFileW(pattern.c_str(), &data);
    if (find == INVALID_HANDLE_VALUE) {
        if (error_message != nullptr) {
            const DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND) {
                *error_message = L"No native Codex backup file was found.";
            } else {
                *error_message = FormatLastErrorMessage(L"FindFirstFileW");
            }
        }
        return false;
    }

    bool found = false;
    FILETIME latest_write_time{};
    std::wstring latest_path;
    do {
        if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
            continue;
        }

        if (!found || CompareFileTime(&data.ftLastWriteTime, &latest_write_time) > 0) {
            latest_write_time = data.ftLastWriteTime;
            latest_path = FileSystem::JoinPath(backup_directory, data.cFileName);
            found = true;
        }
    } while (FindNextFileW(find, &data));

    const DWORD find_error = GetLastError();
    FindClose(find);
    if (find_error != ERROR_NO_MORE_FILES) {
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"FindNextFileW");
        }
        return false;
    }
    if (!found) {
        if (error_message != nullptr) {
            *error_message = L"No native Codex backup file was found.";
        }
        return false;
    }

    *backup_path = latest_path;
    return true;
}

std::wstring TimestampForFileName() {
    SYSTEMTIME local{};
    GetLocalTime(&local);

    wchar_t buffer[32]{};
    swprintf_s(buffer,
               L"%04u%02u%02u_%02u%02u%02u",
               local.wYear,
               local.wMonth,
               local.wDay,
               local.wHour,
               local.wMinute,
               local.wSecond);
    return buffer;
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

std::wstring SanitizeFileName(std::wstring name) {
    if (name.empty()) {
        name = L"codex";
    }

    for (wchar_t& ch : name) {
        if (ch < 32 || ch == L'<' || ch == L'>' || ch == L':' || ch == L'"' || ch == L'/' ||
            ch == L'\\' || ch == L'|' || ch == L'?' || ch == L'*') {
            ch = L'_';
        }
    }
    if (name.size() > 48) {
        name.resize(48);
    }
    return name.empty() ? std::wstring(L"codex") : name;
}

std::string BuildBackupJson(const std::wstring& active_display_name,
                            const std::string& config_base64,
                            const std::string& auth_base64) {
    std::ostringstream json;
    json << "{\n"
         << "  \"BackupKind\": \"native-codex-current-folder-before-switch\",\n"
         << "  \"ActiveDisplayName\": \"" << JsonEscape(WideToUtf8(active_display_name)) << "\",\n"
         << "  \"CreatedAtUtc\": \"" << JsonEscape(WideToUtf8(TimestampUtcIsoWide())) << "\",\n"
         << "  \"ConfigTomlBase64\": \"" << JsonEscape(config_base64) << "\",\n"
         << "  \"AuthJsonBase64\": \"" << JsonEscape(auth_base64) << "\"\n"
         << "}\n";
    return json.str();
}

bool ReadOptionalFileBase64(const std::wstring& path,
                            bool exists,
                            std::string* base64,
                            std::wstring* error_message) {
    if (base64 == nullptr) {
        return false;
    }
    if (!exists) {
        base64->clear();
        return true;
    }

    std::vector<unsigned char> bytes;
    if (!FileSystem::ReadFileBytes(path, &bytes, error_message)) {
        return false;
    }

    const bool encoded = EncodeBytesBase64(bytes, base64);
    if (!bytes.empty()) {
        SecureZeroMemory(bytes.data(), bytes.size());
    }
    if (!encoded && error_message != nullptr) {
        *error_message = L"Failed to base64 encode current Codex file.";
    }
    return encoded;
}

}  // namespace

bool CodexProfileBackupService::CreateCurrentFolderBackup(
    const std::wstring& codex_home,
    const std::wstring& backup_directory,
    const std::wstring& active_display_name,
    CodexCurrentFolderBackupResult* result) const {
    if (result == nullptr) {
        return false;
    }

    *result = CodexCurrentFolderBackupResult{};
    const std::wstring config_path = FileSystem::JoinPath(codex_home, L"config.toml");
    const std::wstring auth_path = FileSystem::JoinPath(codex_home, L"auth.json");
    const bool config_exists = FileSystem::Exists(config_path);
    const bool auth_exists = FileSystem::Exists(auth_path);

    if (!config_exists && !auth_exists) {
        result->skipped_no_files = true;
        return true;
    }

    std::string config_base64;
    std::string auth_base64;
    if (!ReadOptionalFileBase64(config_path, config_exists, &config_base64, &result->error_message) ||
        !ReadOptionalFileBase64(auth_path, auth_exists, &auth_base64, &result->error_message)) {
        return false;
    }

    std::string backup_json = BuildBackupJson(active_display_name, config_base64, auth_base64);
    SecureZeroMemory(config_base64.data(), config_base64.size());
    SecureZeroMemory(auth_base64.data(), auth_base64.size());

    SecretStoreDpapi secret_store;
    std::string protected_text;
    if (!secret_store.ProtectUtf8ToBase64(backup_json, &protected_text, &result->error_message)) {
        SecureZeroMemory(backup_json.data(), backup_json.size());
        return false;
    }
    SecureZeroMemory(backup_json.data(), backup_json.size());

    const std::wstring file_name =
        L"native_switch_" + SanitizeFileName(active_display_name) + L"_" + TimestampForFileName() +
        L".bak.dpapi";
    const std::wstring backup_path = FileSystem::JoinPath(backup_directory, file_name);
    if (!FileSystem::WriteUtf8FileAtomic(backup_path, protected_text, &result->error_message)) {
        SecureZeroMemory(protected_text.data(), protected_text.size());
        return false;
    }

    SecureZeroMemory(protected_text.data(), protected_text.size());
    result->backup_path = backup_path;
    return true;
}

bool CodexProfileBackupService::RestoreLatestBackup(const std::wstring& codex_home,
                                                    const std::wstring& backup_directory,
                                                    CodexBackupRestoreResult* result) const {
    if (result == nullptr) {
        return false;
    }

    *result = CodexBackupRestoreResult{};
    if (!FindLatestBackupPath(backup_directory, &result->backup_path, &result->error_message)) {
        return false;
    }

    std::string protected_text;
    if (!FileSystem::ReadUtf8File(result->backup_path, &protected_text, &result->error_message) ||
        protected_text.empty()) {
        if (result->error_message.empty()) {
            result->error_message = L"Selected native Codex backup file is empty.";
        }
        return false;
    }

    SecretStoreDpapi secret_store;
    std::string backup_json;
    if (!secret_store.UnprotectBase64ToUtf8(protected_text, &backup_json, &result->error_message)) {
        ClearString(&protected_text);
        return false;
    }
    ClearString(&protected_text);

    std::string config_base64 = ExtractFirstJsonStringValue(backup_json, "ConfigTomlBase64");
    std::string auth_base64 = ExtractFirstJsonStringValue(backup_json, "AuthJsonBase64");
    ClearString(&backup_json);

    if (config_base64.empty() && auth_base64.empty()) {
        result->error_message =
            L"Selected native Codex backup does not contain config.toml or auth.json bytes.";
        return false;
    }

    std::vector<unsigned char> config_bytes;
    if (!config_base64.empty()) {
        if (!DecodeBase64(config_base64, &config_bytes)) {
            ClearString(&config_base64);
            ClearString(&auth_base64);
            result->error_message = L"Selected native Codex backup has invalid config.toml payload.";
            return false;
        }
        ClearString(&config_base64);
    }

    std::vector<unsigned char> auth_bytes;
    if (!auth_base64.empty()) {
        if (!DecodeBase64(auth_base64, &auth_bytes)) {
            ClearBytes(&config_bytes);
            ClearString(&auth_base64);
            result->error_message = L"Selected native Codex backup has invalid auth.json payload.";
            return false;
        }
        ClearString(&auth_base64);
    }

    if (!config_bytes.empty()) {
        const std::wstring config_path = FileSystem::JoinPath(codex_home, L"config.toml");
        if (!FileSystem::WriteFileBytesAtomic(config_path, config_bytes, &result->error_message)) {
            ClearBytes(&config_bytes);
            ClearBytes(&auth_bytes);
            return false;
        }
        result->restored_config = true;
        ClearBytes(&config_bytes);
    }

    if (!auth_bytes.empty()) {
        const std::wstring auth_path = FileSystem::JoinPath(codex_home, L"auth.json");
        if (!FileSystem::WriteFileBytesAtomic(auth_path, auth_bytes, &result->error_message)) {
            ClearBytes(&auth_bytes);
            return false;
        }
        result->restored_auth = true;
        ClearBytes(&auth_bytes);
    }

    return result->restored_config || result->restored_auth;
}

}  // namespace mytools
