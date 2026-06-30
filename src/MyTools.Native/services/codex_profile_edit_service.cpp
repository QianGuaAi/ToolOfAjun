#include "services/codex_profile_edit_service.h"

#include <cctype>
#include <cwctype>
#include <sstream>
#include <utility>
#include <vector>

#include <windows.h>

#include "services/file_system.h"
#include "services/secret_store_dpapi.h"

namespace mytools {
namespace {

struct JsonObjectRange {
    size_t start = 0;
    size_t end = 0;
    std::string object;
};

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

std::wstring TrimWide(const std::wstring& value) {
    size_t first = 0;
    while (first < value.size() && std::iswspace(value[first]) != 0) {
        ++first;
    }

    size_t last = value.size();
    while (last > first && std::iswspace(value[last - 1]) != 0) {
        --last;
    }

    return value.substr(first, last - first);
}

bool ContainsControlCharacter(const std::wstring& value) {
    for (const wchar_t ch : value) {
        const unsigned int code_point = static_cast<unsigned int>(ch);
        if (code_point < 0x20 || (code_point >= 0x7F && code_point <= 0x9F)) {
            return true;
        }
    }
    return false;
}

bool ValidateMetadataText(const std::wstring& value,
                          size_t max_length,
                          const wchar_t* field_name,
                          std::wstring* error_message) {
    if (value.size() > max_length || ContainsControlCharacter(value)) {
        if (error_message != nullptr) {
            std::wstringstream message;
            message << L"Codex profile " << (field_name == nullptr ? L"metadata" : field_name)
                    << L" must be " << max_length
                    << L" characters or fewer and cannot contain control characters.";
            *error_message = message.str();
        }
        return false;
    }
    return true;
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
    SkipWhitespace(json, &index);
    if (index >= json.size() || json[index] != ':') {
        return false;
    }
    ++index;
    SkipWhitespace(json, &index);
    if (index >= json.size() || json[index] != '[') {
        return false;
    }

    const size_t open_array = index;
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
                *array_start = open_array + 1;
                *array_end = index;
                return true;
            }
        }
        ++index;
    }
    return false;
}

bool FindNextObjectRange(const std::string& json,
                         size_t* index,
                         size_t array_end,
                         JsonObjectRange* range) {
    if (index == nullptr || range == nullptr) {
        return false;
    }

    while (*index < array_end &&
           (std::isspace(static_cast<unsigned char>(json[*index])) != 0 || json[*index] == ',')) {
        ++(*index);
    }
    if (*index >= array_end || json[*index] != '{') {
        return false;
    }

    const size_t object_start = *index;
    size_t depth = 0;
    bool in_string = false;
    bool escaped = false;
    while (*index < array_end) {
        const char ch = json[*index];
        if (in_string) {
            if (escaped) {
                escaped = false;
            } else if (ch == '\\') {
                escaped = true;
            } else if (ch == '"') {
                in_string = false;
            }
            ++(*index);
            continue;
        }

        if (ch == '"') {
            in_string = true;
        } else if (ch == '{') {
            ++depth;
        } else if (ch == '}') {
            if (depth == 0) {
                return false;
            }
            --depth;
            if (depth == 0) {
                ++(*index);
                range->start = object_start;
                range->end = *index;
                range->object = json.substr(object_start, range->end - range->start);
                return true;
            }
        }
        ++(*index);
    }
    return false;
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
    return display_name;
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

bool LoadProfilesJson(const std::wstring& profiles_json_path,
                      std::string* json,
                      std::wstring* error_message) {
    if (json == nullptr) {
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

bool SaveActiveDisplayName(const std::wstring& active_json_path,
                           const std::wstring& display_name,
                           std::wstring* error_message) {
    if (active_json_path.empty()) {
        if (error_message != nullptr) {
            *error_message = L"Codex active.json path is required to rename the active profile.";
        }
        return false;
    }

    std::string active_json;
    if (!FileSystem::ReadUtf8File(active_json_path, &active_json, error_message)) {
        return false;
    }
    if (active_json.empty()) {
        if (error_message != nullptr) {
            *error_message = L"Codex active.json is missing or empty.";
        }
        return false;
    }

    if (!SetJsonStringProperty(
            &active_json, "ActiveDisplayName", WideToUtf8(display_name), error_message)) {
        ClearString(&active_json);
        return false;
    }

    const bool saved = FileSystem::WriteUtf8FileAtomic(active_json_path, active_json, error_message);
    ClearString(&active_json);
    return saved;
}

bool HasAnyUpdate(const CodexProfileEditRequest& request) {
    return request.update_display_name || request.update_note || request.update_remark ||
           request.update_tags;
}

}  // namespace

bool CodexProfileEditService::UpdateProfileMetadata(
    const CodexProfileEditRequest& request,
    CodexProfileEditResult* result) const {
    if (result == nullptr) {
        return false;
    }
    *result = CodexProfileEditResult{};

    if (request.target_display_name.empty()) {
        result->error_message = L"Target Codex profile display name is required.";
        return false;
    }
    if (!HasAnyUpdate(request)) {
        result->error_message = L"No Codex profile metadata update was requested.";
        return false;
    }
    std::wstring validated_new_display_name = request.new_display_name;
    std::wstring validated_note = request.note;
    std::wstring validated_remark = request.remark;
    std::wstring validated_tags = request.tags;
    if (request.update_display_name) {
        validated_new_display_name = TrimWide(request.new_display_name);
        if (validated_new_display_name.empty()) {
            result->error_message = L"New Codex profile display name is required.";
            return false;
        }
        if (validated_new_display_name.size() > 120 ||
            ContainsControlCharacter(validated_new_display_name)) {
            result->error_message =
                L"New Codex profile display name must be 120 characters or fewer and cannot contain control characters.";
            return false;
        }
    }
    if (request.update_note) {
        validated_note = TrimWide(request.note);
        if (!ValidateMetadataText(validated_note, 500, L"note", &result->error_message)) {
            return false;
        }
    }
    if (request.update_remark) {
        validated_remark = TrimWide(request.remark);
        if (!ValidateMetadataText(validated_remark, 500, L"remark", &result->error_message)) {
            return false;
        }
    }
    if (request.update_tags) {
        validated_tags = TrimWide(request.tags);
        if (!ValidateMetadataText(validated_tags, 200, L"tags", &result->error_message)) {
            return false;
        }
    }

    std::string json;
    if (!LoadProfilesJson(request.profiles_json_path, &json, &result->error_message)) {
        return false;
    }

    size_t array_start = 0;
    size_t array_end = 0;
    if (!FindItemsArrayBounds(json, &array_start, &array_end)) {
        ClearString(&json);
        result->error_message = L"Codex profiles.json does not contain an items array.";
        return false;
    }

    JsonObjectRange target;
    bool found_target = false;
    bool duplicate_display_name = false;
    const std::wstring next_display_name =
        request.update_display_name ? validated_new_display_name : request.target_display_name;

    size_t index = array_start;
    JsonObjectRange range;
    while (FindNextObjectRange(json, &index, array_end, &range)) {
        const std::string display_name = ProfileDisplayName(range.object);
        const bool is_target = EqualsOrdinalIgnoreCaseUtf8(display_name, request.target_display_name);
        if (is_target && !found_target) {
            target = range;
            found_target = true;
        } else if (request.update_display_name &&
                   EqualsOrdinalIgnoreCaseUtf8(display_name, next_display_name)) {
            duplicate_display_name = true;
        }
        ClearString(&range.object);
    }

    if (!found_target) {
        ClearString(&json);
        result->error_message = L"Target Codex profile was not found in profiles.json.";
        return false;
    }
    result->profile_found = true;

    if (duplicate_display_name) {
        ClearString(&target.object);
        ClearString(&json);
        result->error_message = L"Another Codex profile already uses the requested display name.";
        return false;
    }

    const std::wstring current_display_name = Utf8ToWide(ProfileDisplayName(target.object));
    result->saved_display_name = next_display_name;
    result->display_name_changed =
        request.update_display_name &&
        !EqualsOrdinalIgnoreCase(current_display_name, validated_new_display_name);
    const bool target_is_active =
        result->display_name_changed &&
        EqualsOrdinalIgnoreCase(current_display_name, request.active_display_name);
    if (target_is_active && request.active_json_path.empty()) {
        ClearString(&target.object);
        ClearString(&json);
        result->error_message =
            L"Renaming the active Codex profile requires active.json synchronization.";
        return false;
    }

    std::string next_object = target.object;
    if (request.update_display_name) {
        const std::string name_utf8 = WideToUtf8(validated_new_display_name);
        if (!SetJsonStringProperty(&next_object, "DisplayName", name_utf8, &result->error_message) ||
            !SetJsonStringProperty(&next_object, "Name", name_utf8, &result->error_message)) {
            ClearString(&next_object);
            ClearString(&target.object);
            ClearString(&json);
            return false;
        }
    }
    if (request.update_note &&
        !SetJsonStringProperty(&next_object, "Note", WideToUtf8(validated_note), &result->error_message)) {
        ClearString(&next_object);
        ClearString(&target.object);
        ClearString(&json);
        return false;
    }
    if (request.update_remark &&
        !SetJsonStringProperty(&next_object, "Remark", WideToUtf8(validated_remark), &result->error_message)) {
        ClearString(&next_object);
        ClearString(&target.object);
        ClearString(&json);
        return false;
    }
    if (request.update_tags &&
        !SetJsonStringProperty(&next_object, "Tags", WideToUtf8(validated_tags), &result->error_message)) {
        ClearString(&next_object);
        ClearString(&target.object);
        ClearString(&json);
        return false;
    }

    std::string next_json = json.substr(0, target.start);
    next_json += next_object;
    next_json += json.substr(target.end);
    ClearString(&next_object);
    ClearString(&target.object);
    ClearString(&json);

    if (!SaveProfilesJson(request.profiles_json_path, next_json, &result->error_message)) {
        ClearString(&next_json);
        return false;
    }
    result->profiles_json_updated = true;
    ClearString(&next_json);

    if (target_is_active &&
        !SaveActiveDisplayName(
            request.active_json_path, validated_new_display_name, &result->error_message)) {
        result->active_json_updated = false;
        const std::wstring detail = result->error_message;
        result->error_message =
            L"Codex profile was renamed in profiles.json, but active.json could not be synchronized.";
        if (!detail.empty()) {
            result->error_message += L"\n";
            result->error_message += detail;
        }
        return false;
    }
    result->active_json_updated = target_is_active;
    return true;
}

}  // namespace mytools
