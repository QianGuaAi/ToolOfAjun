#include "services/codex_profile_box_service.h"

#include <algorithm>
#include <cctype>
#include <cstddef>
#include <cstdint>
#include <sstream>
#include <utility>
#include <vector>

#include <windows.h>
#include <bcrypt.h>
#include <wincrypt.h>

#include "services/file_system.h"
#include "services/secret_store_dpapi.h"

namespace mytools {
namespace {

constexpr uint32_t kBoxSchemaVersion = 1;
constexpr uint32_t kProfilesSchemaVersion = 2;
constexpr uint32_t kBoxIterations = 200000;
constexpr size_t kSaltSize = 16;
constexpr size_t kIvSize = 16;
constexpr size_t kAesKeySize = 32;
constexpr size_t kHmacKeySize = 32;
constexpr size_t kDerivedKeySize = kAesKeySize + kHmacKeySize;
constexpr char kBoxHeader[] = "CDXB";
constexpr char kPortableKind[] = "portable-codex-profiles-v2";

struct JsonObjectRange {
    size_t start = 0;
    size_t end = 0;
    std::string object;
};

struct BCryptAlgCloser {
    void operator()(BCRYPT_ALG_HANDLE handle) const {
        if (handle != nullptr) {
            BCryptCloseAlgorithmProvider(handle, 0);
        }
    }
};

struct BCryptKeyCloser {
    void operator()(BCRYPT_KEY_HANDLE handle) const {
        if (handle != nullptr) {
            BCryptDestroyKey(handle);
        }
    }
};

struct BCryptHashCloser {
    void operator()(BCRYPT_HASH_HANDLE handle) const {
        if (handle != nullptr) {
            BCryptDestroyHash(handle);
        }
    }
};

template <typename THandle, typename TCloser>
class ScopedBCryptHandle {
public:
    ScopedBCryptHandle() = default;
    ~ScopedBCryptHandle() { Reset(nullptr); }

    ScopedBCryptHandle(const ScopedBCryptHandle&) = delete;
    ScopedBCryptHandle& operator=(const ScopedBCryptHandle&) = delete;

    THandle* Put() {
        Reset(nullptr);
        return &handle_;
    }

    THandle Get() const { return handle_; }

    void Reset(THandle next) {
        if (handle_ != nullptr) {
            TCloser{}(handle_);
        }
        handle_ = next;
    }

private:
    THandle handle_ = nullptr;
};

using ScopedAlg = ScopedBCryptHandle<BCRYPT_ALG_HANDLE, BCryptAlgCloser>;
using ScopedKey = ScopedBCryptHandle<BCRYPT_KEY_HANDLE, BCryptKeyCloser>;
using ScopedHash = ScopedBCryptHandle<BCRYPT_HASH_HANDLE, BCryptHashCloser>;

bool StatusOk(NTSTATUS status) {
    return status >= 0;
}

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

void ClearStringVector(std::vector<std::string>* values) {
    if (values == nullptr) {
        return;
    }
    for (std::string& value : *values) {
        ClearString(&value);
    }
}

std::string BytesToString(const std::vector<unsigned char>& bytes) {
    if (bytes.empty()) {
        return {};
    }
    return std::string(reinterpret_cast<const char*>(bytes.data()),
                       reinterpret_cast<const char*>(bytes.data()) + bytes.size());
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
        return L"MyToolsNative";
    }
    return std::wstring(buffer, size);
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
    return CryptStringToBinaryA(text.c_str(),
                                static_cast<DWORD>(text.size()),
                                CRYPT_STRING_BASE64 | CRYPT_STRING_STRICT,
                                bytes->data(),
                                &required,
                                nullptr,
                                nullptr) != FALSE;
}

bool EncodeBase64(const unsigned char* bytes, size_t size, std::string* text) {
    if (text == nullptr || bytes == nullptr || size == 0 || size > MAXDWORD) {
        return false;
    }

    DWORD required = 0;
    if (!CryptBinaryToStringA(bytes,
                              static_cast<DWORD>(size),
                              CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                              nullptr,
                              &required) ||
        required == 0) {
        return false;
    }

    text->assign(required, '\0');
    if (!CryptBinaryToStringA(bytes,
                              static_cast<DWORD>(size),
                              CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                              text->data(),
                              &required)) {
        text->clear();
        return false;
    }

    if (required > 0 && !text->empty() && text->back() == '\0') {
        text->resize(required - 1);
    } else {
        text->resize(required);
    }
    return true;
}

bool EncodeBase64String(const std::string& value, std::string* text) {
    if (value.empty()) {
        return false;
    }
    return EncodeBase64(reinterpret_cast<const unsigned char*>(value.data()), value.size(), text);
}

bool EqualsOrdinalIgnoreCase(const std::wstring& left, const std::wstring& right) {
    if (left.empty() || right.empty()) {
        return false;
    }
    return CompareStringOrdinal(left.c_str(), -1, right.c_str(), -1, TRUE) == CSTR_EQUAL;
}

bool EqualsOrdinalIgnoreCaseUtf8(const std::string& left_utf8, const std::wstring& right) {
    return EqualsOrdinalIgnoreCase(Utf8ToWide(left_utf8), right);
}

std::string ProfileDisplayName(const std::string& object) {
    std::string display_name = ExtractFirstJsonStringValue(object, "DisplayName");
    if (display_name.empty()) {
        display_name = ExtractFirstJsonStringValue(object, "Name");
    }
    if (display_name.empty()) {
        display_name = "Codex Profile";
    }
    return display_name;
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

void AppendJsonStringField(std::ostringstream* json,
                           const char* property_name,
                           const std::string& value,
                           bool comma = true) {
    if (json == nullptr || property_name == nullptr) {
        return;
    }
    *json << "      \"" << property_name << "\": \"" << JsonEscape(value) << "\"";
    if (comma) {
        *json << ",";
    }
    *json << "\n";
}

void AppendJsonNullableStringField(std::ostringstream* json,
                                   const char* property_name,
                                   const std::string& value,
                                   bool comma = true) {
    if (json == nullptr || property_name == nullptr) {
        return;
    }
    *json << "      \"" << property_name << "\": ";
    if (value.empty()) {
        *json << "null";
    } else {
        *json << "\"" << JsonEscape(value) << "\"";
    }
    if (comma) {
        *json << ",";
    }
    *json << "\n";
}

std::string BuildPortableItemJson(const std::string& source_object,
                                  const std::string& config_base64,
                                  const std::string& auth_base64) {
    const std::string display_name = ProfileDisplayName(source_object);
    std::string imported_at = ExtractFirstJsonStringValue(source_object, "LastImportedAt");
    if (imported_at.empty()) {
        imported_at = WideToUtf8(TimestampUtcIsoWide());
    }

    std::ostringstream json;
    json << "    {\n";
    AppendJsonStringField(&json, "DisplayName", display_name);
    AppendJsonStringField(&json, "Name", display_name);
    AppendJsonStringField(&json, "AccountEmail", ExtractFirstJsonStringValue(source_object, "AccountEmail"));
    AppendJsonStringField(&json, "Note", ExtractFirstJsonStringValue(source_object, "Note"));
    AppendJsonStringField(&json, "Remark", ExtractFirstJsonStringValue(source_object, "Remark"));
    AppendJsonStringField(&json, "Tags", ExtractFirstJsonStringValue(source_object, "Tags"));
    AppendJsonStringField(&json, "FolderPath", ExtractFirstJsonStringValue(source_object, "FolderPath"));
    AppendJsonNullableStringField(&json, "LastAppliedAt", ExtractFirstJsonStringValue(source_object, "LastAppliedAt"));
    AppendJsonStringField(&json, "LastImportedAt", imported_at);
    AppendJsonNullableStringField(
        &json, "AccessTokenExpiresAt", ExtractFirstJsonStringValue(source_object, "AccessTokenExpiresAt"));
    json << "      \"RefreshTokenExpiresAt\": null,\n";
    AppendJsonStringField(&json, "Status", ExtractFirstJsonStringValue(source_object, "Status"));
    AppendJsonStringField(&json, "ConfigTomlBase64", config_base64);
    AppendJsonStringField(&json, "AuthJsonBase64", auth_base64);
    json << "      \"EnableRotation\": false,\n";
    json << "      \"RotationPriority\": 0,\n";
    AppendJsonStringField(&json, "RelayTestStatus", ExtractFirstJsonStringValue(source_object, "RelayTestStatus"));
    AppendJsonNullableStringField(&json, "RelayTestedAt", ExtractFirstJsonStringValue(source_object, "RelayTestedAt"));
    AppendJsonStringField(
        &json, "RelayTestMessage", ExtractFirstJsonStringValue(source_object, "RelayTestMessage"), false);
    json << "    }";
    return json.str();
}

std::string BuildPortableJson(const std::vector<std::string>& item_json) {
    std::ostringstream json;
    json << "{\n"
         << "  \"schemaVersion\": " << kProfilesSchemaVersion << ",\n"
         << "  \"packageKind\": \"" << kPortableKind << "\",\n"
         << "  \"machineName\": \"" << JsonEscape(WideToUtf8(MachineName())) << "\",\n"
         << "  \"createdAtUtc\": \"" << JsonEscape(WideToUtf8(TimestampUtcIsoWide())) << "\",\n"
         << "  \"items\": [\n";
    for (size_t index = 0; index < item_json.size(); ++index) {
        json << item_json[index];
        if (index + 1 < item_json.size()) {
            json << ",";
        }
        json << "\n";
    }
    json << "  ]\n"
         << "}\n";
    return json.str();
}

bool BuildPortableJsonFromProfiles(const std::string& profiles_json,
                                   std::string* portable_json,
                                   size_t* exported_count,
                                   std::wstring* error_message) {
    if (portable_json == nullptr || exported_count == nullptr) {
        return false;
    }
    *exported_count = 0;

    std::vector<std::string> objects = ExtractJsonObjectArray(profiles_json, "items");
    std::vector<std::string> portable_items;
    portable_items.reserve(objects.size());

    SecretStoreDpapi secret_store;
    for (std::string& object : objects) {
        std::string protected_config =
            ProtectedField(object, "ProtectedConfigTomlBase64", "ConfigTomlContentProtected");
        std::string protected_auth =
            ProtectedField(object, "ProtectedAuthJsonBase64", "AuthJsonContentProtected");
        if (protected_config.empty() || protected_auth.empty()) {
            if (error_message != nullptr) {
                *error_message = L"Codex profile library contains an item without config/auth content.";
            }
            ClearString(&protected_config);
            ClearString(&protected_auth);
            ClearStringVector(&portable_items);
            ClearStringVector(&objects);
            return false;
        }

        std::string config_text;
        std::string auth_text;
        if (!secret_store.UnprotectBase64ToUtf8(protected_config, &config_text, error_message) ||
            !secret_store.UnprotectBase64ToUtf8(protected_auth, &auth_text, error_message)) {
            ClearString(&protected_config);
            ClearString(&protected_auth);
            ClearString(&config_text);
            ClearString(&auth_text);
            ClearStringVector(&portable_items);
            ClearStringVector(&objects);
            return false;
        }
        ClearString(&protected_config);
        ClearString(&protected_auth);

        std::string config_base64;
        std::string auth_base64;
        const bool encoded =
            EncodeBase64String(config_text, &config_base64) && EncodeBase64String(auth_text, &auth_base64);
        ClearString(&config_text);
        ClearString(&auth_text);
        if (!encoded) {
            if (error_message != nullptr) {
                *error_message = L"Failed to encode Codex profile content for portable package.";
            }
            ClearString(&config_base64);
            ClearString(&auth_base64);
            ClearStringVector(&portable_items);
            ClearStringVector(&objects);
            return false;
        }

        portable_items.push_back(BuildPortableItemJson(object, config_base64, auth_base64));
        ClearString(&config_base64);
        ClearString(&auth_base64);
        ++(*exported_count);
    }

    *portable_json = BuildPortableJson(portable_items);
    ClearStringVector(&portable_items);
    ClearStringVector(&objects);
    return true;
}

bool GenerateRandomBytes(size_t size, std::vector<unsigned char>* bytes, std::wstring* error_message) {
    if (bytes == nullptr || size == 0 || size > MAXDWORD) {
        return false;
    }
    bytes->assign(size, 0);
    const NTSTATUS status =
        BCryptGenRandom(nullptr, bytes->data(), static_cast<ULONG>(bytes->size()), BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptGenRandom failed.";
        }
        bytes->clear();
        return false;
    }
    return true;
}

bool DeriveBoxKey(const std::string& password_utf8,
                  const std::vector<unsigned char>& salt,
                  std::vector<unsigned char>* key_material,
                  std::wstring* error_message) {
    if (key_material == nullptr || password_utf8.empty() || salt.empty() ||
        password_utf8.size() > MAXDWORD || salt.size() > MAXDWORD) {
        if (error_message != nullptr) {
            *error_message = L"Invalid .codexbox KDF input.";
        }
        return false;
    }

    ScopedAlg algorithm;
    NTSTATUS status = BCryptOpenAlgorithmProvider(
        algorithm.Put(), BCRYPT_SHA256_ALGORITHM, nullptr, BCRYPT_ALG_HANDLE_HMAC_FLAG);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptOpenAlgorithmProvider failed for PBKDF2.";
        }
        return false;
    }

    key_material->assign(kDerivedKeySize, 0);
    status = BCryptDeriveKeyPBKDF2(algorithm.Get(),
                                   reinterpret_cast<PUCHAR>(const_cast<char*>(password_utf8.data())),
                                   static_cast<ULONG>(password_utf8.size()),
                                   const_cast<PUCHAR>(salt.data()),
                                   static_cast<ULONG>(salt.size()),
                                   kBoxIterations,
                                   key_material->data(),
                                   static_cast<ULONG>(key_material->size()),
                                   0);
    if (!StatusOk(status)) {
        ClearBytes(key_material);
        if (error_message != nullptr) {
            *error_message = L"BCryptDeriveKeyPBKDF2 failed.";
        }
        return false;
    }
    return true;
}

bool HmacSha256(const std::vector<unsigned char>& key,
                const unsigned char* data,
                size_t size,
                std::vector<unsigned char>* mac,
                std::wstring* error_message) {
    if (mac == nullptr || key.empty() || data == nullptr || size == 0 || key.size() > MAXDWORD ||
        size > MAXDWORD) {
        return false;
    }

    ScopedAlg algorithm;
    NTSTATUS status = BCryptOpenAlgorithmProvider(
        algorithm.Put(), BCRYPT_SHA256_ALGORITHM, nullptr, BCRYPT_ALG_HANDLE_HMAC_FLAG);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptOpenAlgorithmProvider failed for HMAC.";
        }
        return false;
    }

    DWORD object_length = 0;
    DWORD returned = 0;
    status = BCryptGetProperty(algorithm.Get(),
                               BCRYPT_OBJECT_LENGTH,
                               reinterpret_cast<PUCHAR>(&object_length),
                               sizeof(object_length),
                               &returned,
                               0);
    if (!StatusOk(status) || object_length == 0) {
        if (error_message != nullptr) {
            *error_message = L"BCryptGetProperty failed for HMAC object length.";
        }
        return false;
    }

    std::vector<unsigned char> hash_object(object_length, 0);
    ScopedHash hash;
    status = BCryptCreateHash(algorithm.Get(),
                              hash.Put(),
                              hash_object.data(),
                              static_cast<ULONG>(hash_object.size()),
                              const_cast<PUCHAR>(key.data()),
                              static_cast<ULONG>(key.size()),
                              0);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptCreateHash failed for HMAC.";
        }
        ClearBytes(&hash_object);
        return false;
    }

    status = BCryptHashData(hash.Get(), const_cast<PUCHAR>(data), static_cast<ULONG>(size), 0);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptHashData failed for HMAC.";
        }
        ClearBytes(&hash_object);
        return false;
    }

    mac->assign(32, 0);
    status = BCryptFinishHash(hash.Get(), mac->data(), static_cast<ULONG>(mac->size()), 0);
    ClearBytes(&hash_object);
    if (!StatusOk(status)) {
        ClearBytes(mac);
        if (error_message != nullptr) {
            *error_message = L"BCryptFinishHash failed for HMAC.";
        }
        return false;
    }
    return true;
}

bool AesCbcCrypt(bool encrypt,
                 const std::vector<unsigned char>& input,
                 const unsigned char* key,
                 const std::vector<unsigned char>& iv,
                 std::vector<unsigned char>* output,
                 std::wstring* error_message) {
    if (output == nullptr || input.empty() || key == nullptr || iv.size() != kIvSize ||
        input.size() > MAXDWORD) {
        return false;
    }

    ScopedAlg algorithm;
    NTSTATUS status = BCryptOpenAlgorithmProvider(algorithm.Put(), BCRYPT_AES_ALGORITHM, nullptr, 0);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptOpenAlgorithmProvider failed for AES.";
        }
        return false;
    }

    status = BCryptSetProperty(algorithm.Get(),
                               BCRYPT_CHAINING_MODE,
                               reinterpret_cast<PUCHAR>(const_cast<wchar_t*>(BCRYPT_CHAIN_MODE_CBC)),
                               static_cast<ULONG>(sizeof(BCRYPT_CHAIN_MODE_CBC)),
                               0);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptSetProperty failed for AES-CBC.";
        }
        return false;
    }

    DWORD object_length = 0;
    DWORD returned = 0;
    status = BCryptGetProperty(algorithm.Get(),
                               BCRYPT_OBJECT_LENGTH,
                               reinterpret_cast<PUCHAR>(&object_length),
                               sizeof(object_length),
                               &returned,
                               0);
    if (!StatusOk(status) || object_length == 0) {
        if (error_message != nullptr) {
            *error_message = L"BCryptGetProperty failed for AES key object length.";
        }
        return false;
    }

    std::vector<unsigned char> key_object(object_length, 0);
    ScopedKey aes_key;
    status = BCryptGenerateSymmetricKey(algorithm.Get(),
                                        aes_key.Put(),
                                        key_object.data(),
                                        static_cast<ULONG>(key_object.size()),
                                        const_cast<PUCHAR>(key),
                                        static_cast<ULONG>(kAesKeySize),
                                        0);
    if (!StatusOk(status)) {
        if (error_message != nullptr) {
            *error_message = L"BCryptGenerateSymmetricKey failed.";
        }
        ClearBytes(&key_object);
        return false;
    }

    std::vector<unsigned char> iv_copy = iv;
    ULONG output_size = 0;
    if (encrypt) {
        status = BCryptEncrypt(aes_key.Get(),
                               const_cast<PUCHAR>(input.data()),
                               static_cast<ULONG>(input.size()),
                               nullptr,
                               iv_copy.data(),
                               static_cast<ULONG>(iv_copy.size()),
                               nullptr,
                               0,
                               &output_size,
                               BCRYPT_BLOCK_PADDING);
    } else {
        status = BCryptDecrypt(aes_key.Get(),
                               const_cast<PUCHAR>(input.data()),
                               static_cast<ULONG>(input.size()),
                               nullptr,
                               iv_copy.data(),
                               static_cast<ULONG>(iv_copy.size()),
                               nullptr,
                               0,
                               &output_size,
                               BCRYPT_BLOCK_PADDING);
    }
    if (!StatusOk(status) || output_size == 0) {
        if (error_message != nullptr) {
            *error_message = encrypt ? L"BCryptEncrypt size calculation failed."
                                     : L"BCryptDecrypt size calculation failed.";
        }
        ClearBytes(&key_object);
        ClearBytes(&iv_copy);
        return false;
    }

    output->assign(output_size, 0);
    iv_copy = iv;
    if (encrypt) {
        status = BCryptEncrypt(aes_key.Get(),
                               const_cast<PUCHAR>(input.data()),
                               static_cast<ULONG>(input.size()),
                               nullptr,
                               iv_copy.data(),
                               static_cast<ULONG>(iv_copy.size()),
                               output->data(),
                               static_cast<ULONG>(output->size()),
                               &output_size,
                               BCRYPT_BLOCK_PADDING);
    } else {
        status = BCryptDecrypt(aes_key.Get(),
                               const_cast<PUCHAR>(input.data()),
                               static_cast<ULONG>(input.size()),
                               nullptr,
                               iv_copy.data(),
                               static_cast<ULONG>(iv_copy.size()),
                               output->data(),
                               static_cast<ULONG>(output->size()),
                               &output_size,
                               BCRYPT_BLOCK_PADDING);
    }
    ClearBytes(&key_object);
    ClearBytes(&iv_copy);
    if (!StatusOk(status)) {
        ClearBytes(output);
        if (error_message != nullptr) {
            *error_message = encrypt ? L"BCryptEncrypt failed." : L"BCryptDecrypt failed.";
        }
        return false;
    }
    output->resize(output_size);
    return true;
}

bool ConstantTimeEquals(const std::vector<unsigned char>& left,
                        const std::vector<unsigned char>& right) {
    if (left.size() != right.size()) {
        return false;
    }

    unsigned char diff = 0;
    for (size_t index = 0; index < left.size(); ++index) {
        diff = static_cast<unsigned char>(diff | (left[index] ^ right[index]));
    }
    return diff == 0;
}

void AppendLe16(std::vector<unsigned char>* bytes, uint16_t value) {
    bytes->push_back(static_cast<unsigned char>(value & 0xFF));
    bytes->push_back(static_cast<unsigned char>((value >> 8) & 0xFF));
}

void AppendLe32(std::vector<unsigned char>* bytes, uint32_t value) {
    bytes->push_back(static_cast<unsigned char>(value & 0xFF));
    bytes->push_back(static_cast<unsigned char>((value >> 8) & 0xFF));
    bytes->push_back(static_cast<unsigned char>((value >> 16) & 0xFF));
    bytes->push_back(static_cast<unsigned char>((value >> 24) & 0xFF));
}

bool ReadLe16(const std::vector<unsigned char>& bytes, size_t* offset, uint16_t* value) {
    if (offset == nullptr || value == nullptr || *offset + 2 > bytes.size()) {
        return false;
    }
    *value = static_cast<uint16_t>(bytes[*offset] | (bytes[*offset + 1] << 8));
    *offset += 2;
    return true;
}

bool ReadLe32(const std::vector<unsigned char>& bytes, size_t* offset, uint32_t* value) {
    if (offset == nullptr || value == nullptr || *offset + 4 > bytes.size()) {
        return false;
    }
    *value = static_cast<uint32_t>(bytes[*offset]) |
             (static_cast<uint32_t>(bytes[*offset + 1]) << 8) |
             (static_cast<uint32_t>(bytes[*offset + 2]) << 16) |
             (static_cast<uint32_t>(bytes[*offset + 3]) << 24);
    *offset += 4;
    return true;
}

bool ReadBytes(const std::vector<unsigned char>& bytes,
               size_t* offset,
               size_t size,
               std::vector<unsigned char>* output) {
    if (offset == nullptr || output == nullptr || *offset + size > bytes.size()) {
        return false;
    }
    output->assign(bytes.begin() + static_cast<std::ptrdiff_t>(*offset),
                   bytes.begin() + static_cast<std::ptrdiff_t>(*offset + size));
    *offset += size;
    return true;
}

bool EncryptBoxPayload(const std::string& plain_json,
                       const std::wstring& password,
                       std::vector<unsigned char>* box_bytes,
                       std::wstring* error_message) {
    if (box_bytes == nullptr) {
        return false;
    }

    std::string password_utf8 = WideToUtf8(password);
    if (password_utf8.empty()) {
        if (error_message != nullptr) {
            *error_message = L".codexbox password is required.";
        }
        return false;
    }

    std::vector<unsigned char> salt;
    std::vector<unsigned char> iv;
    std::vector<unsigned char> key_material;
    std::vector<unsigned char> ciphertext;
    if (!GenerateRandomBytes(kSaltSize, &salt, error_message) ||
        !GenerateRandomBytes(kIvSize, &iv, error_message) ||
        !DeriveBoxKey(password_utf8, salt, &key_material, error_message)) {
        ClearString(&password_utf8);
        ClearBytes(&salt);
        ClearBytes(&iv);
        return false;
    }
    ClearString(&password_utf8);

    std::vector<unsigned char> plaintext(plain_json.begin(), plain_json.end());
    const bool encrypted =
        AesCbcCrypt(true, plaintext, key_material.data(), iv, &ciphertext, error_message);
    ClearBytes(&plaintext);
    if (!encrypted) {
        ClearBytes(&salt);
        ClearBytes(&iv);
        ClearBytes(&key_material);
        return false;
    }

    std::vector<unsigned char> pre_mac;
    pre_mac.insert(pre_mac.end(), kBoxHeader, kBoxHeader + 4);
    AppendLe32(&pre_mac, kBoxSchemaVersion);
    AppendLe16(&pre_mac, static_cast<uint16_t>(salt.size()));
    pre_mac.insert(pre_mac.end(), salt.begin(), salt.end());
    AppendLe32(&pre_mac, kBoxIterations);
    AppendLe16(&pre_mac, static_cast<uint16_t>(iv.size()));
    pre_mac.insert(pre_mac.end(), iv.begin(), iv.end());
    AppendLe32(&pre_mac, static_cast<uint32_t>(ciphertext.size()));
    pre_mac.insert(pre_mac.end(), ciphertext.begin(), ciphertext.end());

    std::vector<unsigned char> hmac_key(key_material.begin() + kAesKeySize, key_material.end());
    std::vector<unsigned char> mac;
    const bool hmac_ok = HmacSha256(hmac_key, pre_mac.data(), pre_mac.size(), &mac, error_message);
    ClearBytes(&hmac_key);
    ClearBytes(&key_material);
    ClearBytes(&salt);
    ClearBytes(&iv);
    ClearBytes(&ciphertext);
    if (!hmac_ok) {
        ClearBytes(&pre_mac);
        return false;
    }

    *box_bytes = pre_mac;
    AppendLe16(box_bytes, static_cast<uint16_t>(mac.size()));
    box_bytes->insert(box_bytes->end(), mac.begin(), mac.end());
    ClearBytes(&pre_mac);
    ClearBytes(&mac);
    return true;
}

bool DecryptBoxPayload(const std::vector<unsigned char>& box_bytes,
                       const std::wstring& password,
                       std::string* plain_json,
                       std::wstring* error_message) {
    if (plain_json == nullptr || box_bytes.size() < 4 + 4 + 2 + kSaltSize + 4 + 2 + kIvSize + 4 + 2 + 32) {
        if (error_message != nullptr) {
            *error_message = L".codexbox file is too small.";
        }
        return false;
    }

    std::string password_utf8 = WideToUtf8(password);
    if (password_utf8.empty()) {
        if (error_message != nullptr) {
            *error_message = L".codexbox password is required.";
        }
        return false;
    }

    size_t offset = 0;
    if (box_bytes[0] != 'C' || box_bytes[1] != 'D' || box_bytes[2] != 'X' || box_bytes[3] != 'B') {
        ClearString(&password_utf8);
        if (error_message != nullptr) {
            *error_message = L"Invalid .codexbox header.";
        }
        return false;
    }
    offset += 4;

    uint32_t version = 0;
    if (!ReadLe32(box_bytes, &offset, &version) || version != kBoxSchemaVersion) {
        ClearString(&password_utf8);
        if (error_message != nullptr) {
            *error_message = L"Unsupported .codexbox schema version.";
        }
        return false;
    }

    uint16_t salt_len = 0;
    std::vector<unsigned char> salt;
    if (!ReadLe16(box_bytes, &offset, &salt_len) || salt_len != kSaltSize ||
        !ReadBytes(box_bytes, &offset, salt_len, &salt)) {
        ClearString(&password_utf8);
        if (error_message != nullptr) {
            *error_message = L"Invalid .codexbox salt.";
        }
        return false;
    }

    uint32_t iterations = 0;
    if (!ReadLe32(box_bytes, &offset, &iterations) || iterations != kBoxIterations) {
        ClearString(&password_utf8);
        ClearBytes(&salt);
        if (error_message != nullptr) {
            *error_message = L"Unsupported .codexbox PBKDF2 iteration count.";
        }
        return false;
    }

    uint16_t iv_len = 0;
    std::vector<unsigned char> iv;
    if (!ReadLe16(box_bytes, &offset, &iv_len) || iv_len != kIvSize ||
        !ReadBytes(box_bytes, &offset, iv_len, &iv)) {
        ClearString(&password_utf8);
        ClearBytes(&salt);
        if (error_message != nullptr) {
            *error_message = L"Invalid .codexbox IV.";
        }
        return false;
    }

    uint32_t ciphertext_len = 0;
    std::vector<unsigned char> ciphertext;
    if (!ReadLe32(box_bytes, &offset, &ciphertext_len) || ciphertext_len == 0 ||
        !ReadBytes(box_bytes, &offset, ciphertext_len, &ciphertext)) {
        ClearString(&password_utf8);
        ClearBytes(&salt);
        ClearBytes(&iv);
        if (error_message != nullptr) {
            *error_message = L"Invalid .codexbox ciphertext.";
        }
        return false;
    }
    const size_t mac_data_length = offset;

    uint16_t mac_len = 0;
    std::vector<unsigned char> mac;
    if (!ReadLe16(box_bytes, &offset, &mac_len) || mac_len != 32 ||
        !ReadBytes(box_bytes, &offset, mac_len, &mac) || offset != box_bytes.size()) {
        ClearString(&password_utf8);
        ClearBytes(&salt);
        ClearBytes(&iv);
        ClearBytes(&ciphertext);
        if (error_message != nullptr) {
            *error_message = L"Invalid .codexbox MAC.";
        }
        return false;
    }

    std::vector<unsigned char> key_material;
    if (!DeriveBoxKey(password_utf8, salt, &key_material, error_message)) {
        ClearString(&password_utf8);
        ClearBytes(&salt);
        ClearBytes(&iv);
        ClearBytes(&ciphertext);
        ClearBytes(&mac);
        return false;
    }
    ClearString(&password_utf8);
    ClearBytes(&salt);

    std::vector<unsigned char> hmac_key(key_material.begin() + kAesKeySize, key_material.end());
    std::vector<unsigned char> computed_mac;
    const bool hmac_ok =
        HmacSha256(hmac_key, box_bytes.data(), mac_data_length, &computed_mac, error_message);
    ClearBytes(&hmac_key);
    if (!hmac_ok || !ConstantTimeEquals(mac, computed_mac)) {
        ClearBytes(&key_material);
        ClearBytes(&iv);
        ClearBytes(&ciphertext);
        ClearBytes(&mac);
        ClearBytes(&computed_mac);
        if (error_message != nullptr) {
            *error_message = L".codexbox password is wrong or the file is damaged.";
        }
        return false;
    }
    ClearBytes(&mac);
    ClearBytes(&computed_mac);

    std::vector<unsigned char> plaintext;
    const bool decrypted =
        AesCbcCrypt(false, ciphertext, key_material.data(), iv, &plaintext, error_message);
    ClearBytes(&key_material);
    ClearBytes(&iv);
    ClearBytes(&ciphertext);
    if (!decrypted) {
        return false;
    }

    plain_json->assign(reinterpret_cast<const char*>(plaintext.data()),
                       reinterpret_cast<const char*>(plaintext.data()) + plaintext.size());
    ClearBytes(&plaintext);
    return true;
}

bool FindJsonStringPropertyValueRange(const std::string& object,
                                      const char* property_name,
                                      size_t* value_start,
                                      size_t* value_end) {
    if (property_name == nullptr || value_start == nullptr || value_end == nullptr) {
        return false;
    }

    const std::string property = std::string("\"") + property_name + "\"";
    size_t search = 0;
    while (search < object.size()) {
        const size_t found = object.find(property, search);
        if (found == std::string::npos) {
            return false;
        }

        size_t index = found + property.size();
        SkipWhitespace(object, &index);
        if (index >= object.size() || object[index] != ':') {
            search = found + property.size();
            continue;
        }
        ++index;
        SkipWhitespace(object, &index);
        if (index >= object.size() || object[index] != '"') {
            search = index;
            continue;
        }

        std::string ignored_value;
        size_t value_cursor = index;
        if (!ReadJsonString(object, &value_cursor, &ignored_value)) {
            return false;
        }
        ClearString(&ignored_value);
        *value_start = index;
        *value_end = value_cursor;
        return true;
    }
    return false;
}

size_t ClosingObjectBraceIndex(const std::string& object) {
    if (object.empty() || object.front() != '{') {
        return std::string::npos;
    }

    size_t index = 0;
    size_t depth = 0;
    bool in_string = false;
    bool escaped = false;
    while (index < object.size()) {
        const char ch = object[index];
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
                return std::string::npos;
            }
            --depth;
            if (depth == 0) {
                return index;
            }
        }
        ++index;
    }
    return std::string::npos;
}

bool SetJsonStringProperty(std::string* object,
                           const char* property_name,
                           const std::string& value,
                           std::wstring* error_message) {
    if (object == nullptr || property_name == nullptr) {
        return false;
    }

    const std::string encoded = "\"" + JsonEscape(value) + "\"";
    size_t value_start = 0;
    size_t value_end = 0;
    if (FindJsonStringPropertyValueRange(*object, property_name, &value_start, &value_end)) {
        object->replace(value_start, value_end - value_start, encoded);
        return true;
    }

    const size_t close_index = ClosingObjectBraceIndex(*object);
    if (close_index == std::string::npos) {
        if (error_message != nullptr) {
            *error_message = L"Codex profile item JSON is malformed.";
        }
        return false;
    }

    size_t insert_index = close_index;
    while (insert_index > 1 &&
           std::isspace(static_cast<unsigned char>((*object)[insert_index - 1])) != 0) {
        --insert_index;
    }

    bool has_content = false;
    for (size_t scan = 1; scan < insert_index; ++scan) {
        if (std::isspace(static_cast<unsigned char>((*object)[scan])) == 0) {
            has_content = true;
            break;
        }
    }

    std::string insertion = has_content ? ",\n      \"" : "\n      \"";
    insertion += property_name;
    insertion += "\": ";
    insertion += encoded;
    object->insert(insert_index, insertion);
    return true;
}

std::string BuildLibraryItemFromPortable(const std::string& portable_object,
                                         const std::string& protected_config,
                                         const std::string& protected_auth) {
    const std::string display_name = ProfileDisplayName(portable_object);
    std::string imported_at = ExtractFirstJsonStringValue(portable_object, "LastImportedAt");
    if (imported_at.empty()) {
        imported_at = WideToUtf8(TimestampUtcIsoWide());
    }

    std::ostringstream json;
    json << "    {\n";
    AppendJsonStringField(&json, "DisplayName", display_name);
    AppendJsonStringField(&json, "Name", display_name);
    AppendJsonStringField(&json, "AccountEmail", ExtractFirstJsonStringValue(portable_object, "AccountEmail"));
    AppendJsonStringField(&json, "Note", ExtractFirstJsonStringValue(portable_object, "Note"));
    AppendJsonStringField(&json, "Remark", ExtractFirstJsonStringValue(portable_object, "Remark"));
    AppendJsonStringField(&json, "Tags", ExtractFirstJsonStringValue(portable_object, "Tags"));
    AppendJsonStringField(&json, "FolderPath", ExtractFirstJsonStringValue(portable_object, "FolderPath"));
    AppendJsonNullableStringField(&json, "LastAppliedAt", ExtractFirstJsonStringValue(portable_object, "LastAppliedAt"));
    AppendJsonStringField(&json, "LastImportedAt", imported_at);
    AppendJsonNullableStringField(
        &json, "AccessTokenExpiresAt", ExtractFirstJsonStringValue(portable_object, "AccessTokenExpiresAt"));
    json << "      \"RefreshTokenExpiresAt\": null,\n";
    AppendJsonStringField(&json, "Status", ExtractFirstJsonStringValue(portable_object, "Status"));
    AppendJsonStringField(&json, "ProtectedConfigTomlBase64", protected_config);
    AppendJsonStringField(&json, "ProtectedAuthJsonBase64", protected_auth);
    AppendJsonStringField(&json, "ConfigTomlContentProtected", protected_config);
    AppendJsonStringField(&json, "AuthJsonContentProtected", protected_auth);
    json << "      \"EnableRotation\": false,\n";
    json << "      \"RotationPriority\": 0,\n";
    AppendJsonStringField(&json, "RelayTestStatus", ExtractFirstJsonStringValue(portable_object, "RelayTestStatus"));
    AppendJsonNullableStringField(&json, "RelayTestedAt", ExtractFirstJsonStringValue(portable_object, "RelayTestedAt"));
    AppendJsonStringField(
        &json, "RelayTestMessage", ExtractFirstJsonStringValue(portable_object, "RelayTestMessage"), false);
    json << "    }";
    return json.str();
}

std::string BuildProfilesJsonFromItems(const std::vector<std::string>& item_json) {
    std::ostringstream json;
    json << "{\n"
         << "  \"schemaVersion\": " << kProfilesSchemaVersion << ",\n"
         << "  \"machineName\": \"" << JsonEscape(WideToUtf8(MachineName())) << "\",\n"
         << "  \"createdAtUtc\": \"" << JsonEscape(WideToUtf8(TimestampUtcIsoWide())) << "\",\n"
         << "  \"items\": [\n";
    for (size_t index = 0; index < item_json.size(); ++index) {
        json << item_json[index];
        if (index + 1 < item_json.size()) {
            json << ",";
        }
        json << "\n";
    }
    json << "  ]\n"
         << "}\n";
    return json.str();
}

bool LoadProfilesJson(const std::wstring& profiles_json_path,
                      std::string* json,
                      bool* created_new_library,
                      std::wstring* error_message) {
    if (json == nullptr || created_new_library == nullptr) {
        return false;
    }
    *created_new_library = false;

    if (!FileSystem::Exists(profiles_json_path)) {
        json->clear();
        *created_new_library = true;
        return true;
    }

    std::string protected_text;
    if (!FileSystem::ReadUtf8File(profiles_json_path, &protected_text, error_message)) {
        return false;
    }
    if (protected_text.empty()) {
        json->clear();
        *created_new_library = true;
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

bool SaveProfilesJson(const std::wstring& profiles_json_path,
                      const std::string& json,
                      std::wstring* error_message) {
    SecretStoreDpapi secret_store;
    std::string protected_text;
    if (!secret_store.ProtectUtf8ToBase64(json, &protected_text, error_message)) {
        return false;
    }

    const bool saved = FileSystem::WriteUtf8FileAtomic(profiles_json_path, protected_text, error_message);
    ClearString(&protected_text);
    return saved;
}

int FindProfileIndexByName(const std::vector<std::string>& objects, const std::string& display_name_utf8) {
    const std::wstring display_name = Utf8ToWide(display_name_utf8);
    for (size_t index = 0; index < objects.size(); ++index) {
        if (EqualsOrdinalIgnoreCaseUtf8(ProfileDisplayName(objects[index]), display_name)) {
            return static_cast<int>(index);
        }
    }
    return -1;
}

std::string UniqueDisplayName(const std::vector<std::string>& objects, const std::string& base_name) {
    std::string candidate = base_name.empty() ? "Codex Profile" : base_name;
    int suffix = 2;
    while (FindProfileIndexByName(objects, candidate) >= 0) {
        std::ostringstream renamed;
        renamed << (base_name.empty() ? "Codex Profile" : base_name) << " (" << suffix << ")";
        candidate = renamed.str();
        ++suffix;
    }
    return candidate;
}

bool BuildLibraryItemFromPortableObject(const std::string& portable_object,
                                        std::string* item_json,
                                        std::wstring* error_message) {
    if (item_json == nullptr) {
        return false;
    }

    std::string config_text;
    std::string auth_text;
    std::string portable_config_base64 = ExtractFirstJsonStringValue(portable_object, "ConfigTomlBase64");
    std::string portable_auth_base64 = ExtractFirstJsonStringValue(portable_object, "AuthJsonBase64");

    if (!portable_config_base64.empty() || !portable_auth_base64.empty()) {
        std::vector<unsigned char> config_bytes;
        std::vector<unsigned char> auth_bytes;
        if (!DecodeBase64(portable_config_base64, &config_bytes) ||
            !DecodeBase64(portable_auth_base64, &auth_bytes) ||
            config_bytes.empty() || auth_bytes.empty()) {
            ClearBytes(&config_bytes);
            ClearBytes(&auth_bytes);
            ClearString(&portable_config_base64);
            ClearString(&portable_auth_base64);
            if (error_message != nullptr) {
                *error_message = L".codexbox profile item contains invalid portable config/auth content.";
            }
            return false;
        }

        config_text = BytesToString(config_bytes);
        auth_text = BytesToString(auth_bytes);
        ClearBytes(&config_bytes);
        ClearBytes(&auth_bytes);
        ClearString(&portable_config_base64);
        ClearString(&portable_auth_base64);
    } else {
        std::string protected_config =
            ProtectedField(portable_object, "ProtectedConfigTomlBase64", "ConfigTomlContentProtected");
        std::string protected_auth =
            ProtectedField(portable_object, "ProtectedAuthJsonBase64", "AuthJsonContentProtected");
        if (protected_config.empty() || protected_auth.empty()) {
            ClearString(&protected_config);
            ClearString(&protected_auth);
            if (error_message != nullptr) {
                *error_message = L".codexbox profile item does not contain importable config/auth content.";
            }
            return false;
        }

        SecretStoreDpapi legacy_secret_store;
        if (!legacy_secret_store.UnprotectBase64ToUtf8(protected_config, &config_text, error_message) ||
            !legacy_secret_store.UnprotectBase64ToUtf8(protected_auth, &auth_text, error_message)) {
            ClearString(&protected_config);
            ClearString(&protected_auth);
            ClearString(&config_text);
            ClearString(&auth_text);
            if (error_message != nullptr && error_message->empty()) {
                *error_message = L"Legacy .codexbox content cannot be decrypted by the current Windows user.";
            }
            return false;
        }
        ClearString(&protected_config);
        ClearString(&protected_auth);
    }

    SecretStoreDpapi secret_store;
    std::string protected_config;
    std::string protected_auth;
    if (!secret_store.ProtectUtf8ToBase64(config_text, &protected_config, error_message) ||
        !secret_store.ProtectUtf8ToBase64(auth_text, &protected_auth, error_message)) {
        ClearString(&config_text);
        ClearString(&auth_text);
        ClearString(&protected_config);
        ClearString(&protected_auth);
        return false;
    }
    ClearString(&config_text);
    ClearString(&auth_text);

    *item_json = BuildLibraryItemFromPortable(portable_object, protected_config, protected_auth);
    ClearString(&protected_config);
    ClearString(&protected_auth);
    return true;
}

bool MergePortableJsonIntoLibrary(const std::string& portable_json,
                                  const CodexProfileBoxImportRequest& request,
                                  CodexProfileBoxImportResult* result) {
    if (result == nullptr) {
        return false;
    }

    std::vector<std::string> portable_objects = ExtractJsonObjectArray(portable_json, "items");
    if (portable_objects.empty()) {
        result->error_message = L".codexbox does not contain importable profile items.";
        return false;
    }

    std::string existing_json;
    if (!LoadProfilesJson(request.profiles_json_path,
                          &existing_json,
                          &result->created_new_library,
                          &result->error_message)) {
        ClearStringVector(&portable_objects);
        return false;
    }

    std::vector<std::string> library_objects;
    if (!existing_json.empty()) {
        library_objects = ExtractJsonObjectArray(existing_json, "items");
    }
    ClearString(&existing_json);

    for (std::string& portable_object : portable_objects) {
        std::string item_json;
        if (!BuildLibraryItemFromPortableObject(portable_object, &item_json, &result->error_message)) {
            ClearStringVector(&library_objects);
            ClearStringVector(&portable_objects);
            return false;
        }

        std::string display_name = ProfileDisplayName(item_json);
        const int existing_index = FindProfileIndexByName(library_objects, display_name);
        if (existing_index >= 0) {
            if (request.conflict_policy == CodexProfileBoxConflictPolicy::Skip) {
                ++result->skipped_count;
                ClearString(&item_json);
                continue;
            }
            if (request.conflict_policy == CodexProfileBoxConflictPolicy::Replace) {
                ClearString(&library_objects[static_cast<size_t>(existing_index)]);
                library_objects.erase(library_objects.begin() + existing_index);
                ++result->replaced_count;
            } else {
                const std::string renamed = UniqueDisplayName(library_objects, display_name);
                if (!SetJsonStringProperty(&item_json, "DisplayName", renamed, &result->error_message) ||
                    !SetJsonStringProperty(&item_json, "Name", renamed, &result->error_message)) {
                    ClearString(&item_json);
                    ClearStringVector(&library_objects);
                    ClearStringVector(&portable_objects);
                    return false;
                }
                ++result->renamed_count;
            }
        }

        library_objects.push_back(std::move(item_json));
        ++result->imported_count;
    }

    if (result->imported_count == 0) {
        ClearStringVector(&library_objects);
        ClearStringVector(&portable_objects);
        return true;
    }

    std::string next_json = BuildProfilesJsonFromItems(library_objects);
    ClearStringVector(&library_objects);
    ClearStringVector(&portable_objects);
    if (!SaveProfilesJson(request.profiles_json_path, next_json, &result->error_message)) {
        ClearString(&next_json);
        return false;
    }
    ClearString(&next_json);
    return true;
}

}  // namespace

bool CodexProfileBoxService::ExportBox(const CodexProfileBoxExportRequest& request,
                                       CodexProfileBoxExportResult* result) const {
    if (result == nullptr) {
        return false;
    }
    *result = CodexProfileBoxExportResult{};

    if (request.profiles_json_path.empty()) {
        result->error_message = L"Codex profiles.json path is required.";
        return false;
    }
    if (request.output_path.empty()) {
        result->error_message = L".codexbox output path is required.";
        return false;
    }
    if (request.password.empty()) {
        result->error_message = L".codexbox export password is required.";
        return false;
    }

    bool created_new_library = false;
    std::string profiles_json;
    if (!LoadProfilesJson(request.profiles_json_path,
                          &profiles_json,
                          &created_new_library,
                          &result->error_message) ||
        profiles_json.empty()) {
        ClearString(&profiles_json);
        if (result->error_message.empty()) {
            result->error_message = L"Codex profile library is missing or empty.";
        }
        return false;
    }
    (void)created_new_library;

    std::string portable_json;
    if (!BuildPortableJsonFromProfiles(
            profiles_json, &portable_json, &result->exported_count, &result->error_message)) {
        ClearString(&profiles_json);
        return false;
    }
    ClearString(&profiles_json);

    std::vector<unsigned char> box_bytes;
    if (!EncryptBoxPayload(portable_json, request.password, &box_bytes, &result->error_message)) {
        ClearString(&portable_json);
        return false;
    }
    ClearString(&portable_json);

    const bool saved = FileSystem::WriteFileBytesAtomic(request.output_path, box_bytes, &result->error_message);
    ClearBytes(&box_bytes);
    return saved;
}

bool CodexProfileBoxService::ImportBox(const CodexProfileBoxImportRequest& request,
                                       CodexProfileBoxImportResult* result) const {
    if (result == nullptr) {
        return false;
    }
    *result = CodexProfileBoxImportResult{};

    if (request.profiles_json_path.empty()) {
        result->error_message = L"Codex profiles.json path is required.";
        return false;
    }
    if (request.box_path.empty()) {
        result->error_message = L".codexbox input path is required.";
        return false;
    }
    if (request.password.empty()) {
        result->error_message = L".codexbox import password is required.";
        return false;
    }

    std::vector<unsigned char> box_bytes;
    if (!FileSystem::ReadFileBytes(request.box_path, &box_bytes, &result->error_message) ||
        box_bytes.empty()) {
        ClearBytes(&box_bytes);
        if (result->error_message.empty()) {
            result->error_message = L".codexbox file is missing or empty.";
        }
        return false;
    }

    std::string portable_json;
    if (!DecryptBoxPayload(box_bytes, request.password, &portable_json, &result->error_message)) {
        ClearBytes(&box_bytes);
        return false;
    }
    ClearBytes(&box_bytes);

    const bool merged = MergePortableJsonIntoLibrary(portable_json, request, result);
    ClearString(&portable_json);
    return merged;
}

}  // namespace mytools
