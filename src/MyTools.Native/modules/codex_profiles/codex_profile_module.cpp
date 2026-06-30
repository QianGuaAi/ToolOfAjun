#include "modules/codex_profiles/codex_profile_module.h"

#include <algorithm>
#include <cctype>
#include <sstream>
#include <utility>

#include <windows.h>

#include "services/codex_profile_backup_service.h"
#include "services/codex_profile_box_service.h"
#include "services/codex_profile_diff_service.h"
#include "services/codex_profile_edit_service.h"
#include "services/codex_profile_export_service.h"
#include "services/codex_profile_import_service.h"
#include "services/codex_profile_switch_service.h"
#include "services/file_system.h"
#include "services/secret_store_dpapi.h"

namespace mytools {
namespace {

std::wstring EnvironmentVariable(const wchar_t* name) {
    DWORD required = GetEnvironmentVariableW(name, nullptr, 0);
    if (required == 0) {
        return {};
    }

    std::wstring value(required, L'\0');
    DWORD written = GetEnvironmentVariableW(name, value.data(), required);
    if (written == 0 || written >= required) {
        return {};
    }

    value.resize(written);
    return value;
}

std::wstring YesNo(bool value) {
    return value ? L"yes" : L"no";
}

std::wstring StateLine(const wchar_t* label, bool exists) {
    std::wstring line = label;
    line += L": ";
    line += YesNo(exists);
    return line;
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

std::string ExtractFirstJsonStringValue(const std::string& json, const char* property_name) {
    const std::vector<std::string> values = ExtractJsonStringValues(json, property_name);
    return values.empty() ? std::string() : values.front();
}

std::wstring Utf8ToWide(const std::string& text) {
    if (text.empty()) {
        return {};
    }

    int required = MultiByteToWideChar(CP_UTF8,
                                       MB_ERR_INVALID_CHARS,
                                       text.data(),
                                       static_cast<int>(text.size()),
                                       nullptr,
                                       0);
    if (required <= 0) {
        return L"[unreadable utf8]";
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

std::wstring SummaryLine(size_t index,
                         const std::string& name,
                         const std::string& status,
                         const std::string& imported_at,
                         const std::wstring& active_display_name) {
    std::wstringstream line;
    line << L"profile " << (index + 1) << L": ";
    const std::wstring wide_name = Utf8ToWide(name);
    if (!wide_name.empty()) {
        line << wide_name;
    } else {
        line << L"[unnamed]";
    }

    if (!active_display_name.empty() && !wide_name.empty() &&
        CompareStringOrdinal(wide_name.c_str(),
                             -1,
                             active_display_name.c_str(),
                             -1,
                             TRUE) == CSTR_EQUAL) {
        line << L" / active";
    }
    if (!status.empty()) {
        line << L" / status: " << Utf8ToWide(status);
    }
    if (!imported_at.empty()) {
        line << L" / imported: " << Utf8ToWide(imported_at);
    }
    return line.str();
}

std::wstring TimestampForGeneratedName() {
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

std::wstring ShortHash(const std::wstring& hash) {
    if (hash.empty()) {
        return L"[none]";
    }
    return hash.size() <= 12 ? hash : hash.substr(0, 12);
}

std::wstring FileDiffLine(const CodexProfileFileDiffSummary& file) {
    std::wstringstream line;
    line << file.file_name << L": profile=" << YesNo(file.profile_available)
         << L", current=" << YesNo(file.current_available)
         << L", same=" << YesNo(file.same)
         << L", bytes " << file.profile_size << L"/" << file.current_size
         << L", lines " << file.profile_line_count << L"/" << file.current_line_count
         << L", sha256 " << ShortHash(file.profile_sha256_hex) << L"/"
         << ShortHash(file.current_sha256_hex);
    return line.str();
}

void ClearWideString(std::wstring* value) {
    if (value != nullptr && !value->empty()) {
        SecureZeroMemory(value->data(), value->size() * sizeof(wchar_t));
        value->clear();
    }
}

std::wstring NormalizeDirectoryPath(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    const DWORD required = GetFullPathNameW(path.c_str(), 0, nullptr, nullptr);
    if (required == 0) {
        return path;
    }

    std::wstring normalized(required, L'\0');
    const DWORD written = GetFullPathNameW(path.c_str(), required, normalized.data(), nullptr);
    if (written == 0 || written >= required) {
        return path;
    }
    normalized.resize(written);
    while (normalized.size() > 3 &&
           (normalized.back() == L'\\' || normalized.back() == L'/')) {
        normalized.pop_back();
    }
    return normalized;
}

bool SameDirectoryPath(const std::wstring& left, const std::wstring& right) {
    const std::wstring normalized_left = NormalizeDirectoryPath(left);
    const std::wstring normalized_right = NormalizeDirectoryPath(right);
    return !normalized_left.empty() && !normalized_right.empty() &&
           CompareStringOrdinal(normalized_left.c_str(),
                                -1,
                                normalized_right.c_str(),
                                -1,
                                TRUE) == CSTR_EQUAL;
}

bool SameDisplayName(const std::wstring& left, const std::wstring& right) {
    return !left.empty() && !right.empty() &&
           CompareStringOrdinal(left.c_str(), -1, right.c_str(), -1, TRUE) == CSTR_EQUAL;
}

bool SelectedProfileTarget(const CodexProfileProbe& probe,
                           const std::wstring& requested_target,
                           std::wstring* target,
                           std::wstring* error_message) {
    if (target == nullptr) {
        return false;
    }
    if (!probe.summary_error.empty()) {
        if (error_message != nullptr) {
            *error_message = probe.summary_error;
        }
        return false;
    }
    if (!probe.profiles_summary_loaded || probe.profile_count == 0 ||
        probe.profile_display_names.empty()) {
        if (error_message != nullptr) {
            *error_message = L"No readable Codex profile is available for this explicit menu action.";
        }
        return false;
    }
    if (requested_target.empty()) {
        if (error_message != nullptr) {
            *error_message = L"No Codex profile was selected for this explicit menu action.";
        }
        return false;
    }
    const auto found = std::find_if(
        probe.profile_display_names.begin(),
        probe.profile_display_names.end(),
        [&requested_target](const std::wstring& candidate) {
            return SameDisplayName(candidate, requested_target);
        });
    if (found == probe.profile_display_names.end()) {
        if (error_message != nullptr) {
            *error_message =
                L"The selected Codex profile is no longer available. Refresh summaries and choose again.";
        }
        return false;
    }
    *target = *found;
    return true;
}

}  // namespace

CodexProfileModule::CodexProfileModule(Logger* logger) : logger_(logger) {}

ModuleInfo CodexProfileModule::BuildModuleInfo() const {
    const CodexProfileProbe probe = ProbeLocalState();

    ModuleInfo info;
    info.id = ModuleId::CodexProfiles;
    info.title = L"Codex Profiles";
    info.subtitle =
        L"Stage 3 scaffold: read DPAPI profile metadata without exposing embedded auth/config secrets.";
    info.status = L"Current location: Tools / Codex Profiles";
    info.bullets = {
        L"Profile library: " + probe.paths.profiles_json,
        StateLine(L"profiles.json exists", probe.profiles_json_exists),
        L"profile summaries loaded: " + YesNo(probe.profiles_summary_loaded),
        L"profile count: " + std::to_wstring(probe.profile_count),
        StateLine(L"active.json exists", probe.active_json_exists),
        L"active profile: " +
            (probe.active_display_name.empty() ? std::wstring(L"[not recorded]")
                                               : probe.active_display_name),
        L"first explicit action target: " +
            (probe.first_profile_display_name.empty() ? std::wstring(L"[not available]")
                                                      : probe.first_profile_display_name),
        L"selectable profile targets: " + std::to_wstring(probe.profile_display_names.size()),
        L"Switch backup directory: " + probe.paths.backups_dir,
        L"explicit menu actions: refresh, choose profile for diff/apply/export/metadata edits, backup current folder, import current folder, restore latest backup, .codexbox import/export",
        L"Current Codex config: " + probe.paths.config_toml,
        StateLine(L"config.toml exists", probe.config_toml_exists),
        StateLine(L"auth.json exists", probe.auth_json_exists),
        L".codexbox service: password/file dialogs are wired; import asks how to handle name conflicts",
        L"Next migrations: relay test and multi-profile metadata editor."};
    if (!probe.summary_error.empty()) {
        info.bullets.push_back(L"profile summary error: " + probe.summary_error);
    }
    for (const std::wstring& summary : probe.profile_summaries) {
        info.bullets.push_back(summary);
    }
    return info;
}

CodexProfileProbe CodexProfileModule::ProbeLocalState() const {
    CodexProfileProbe probe;
    probe.paths = ResolvePaths();
    probe.profiles_json_exists = FileSystem::Exists(probe.paths.profiles_json);
    probe.active_json_exists = FileSystem::Exists(probe.paths.active_json);
    probe.config_toml_exists = FileSystem::Exists(probe.paths.config_toml);
    probe.auth_json_exists = FileSystem::Exists(probe.paths.auth_json);
    LoadProfileSummaries(&probe);

    if (logger_ != nullptr) {
        logger_->Info(L"Codex profile module probed local state without reading secrets.");
    }
    return probe;
}

CodexProfileActionResult CodexProfileModule::RunUiAction(
    CodexProfileUiAction action,
    const CodexProfileActionOptions& options) const {
    CodexProfileActionResult action_result;
    const CodexProfileProbe probe = ProbeLocalState();

    switch (action) {
        case CodexProfileUiAction::Refresh: {
            action_result.ok = true;
            action_result.changed_state = true;
            action_result.title = L"Codex Profiles refreshed";
            std::wstringstream message;
            message << L"profiles.json: " << YesNo(probe.profiles_json_exists) << L"\n"
                    << L"profile summaries loaded: " << YesNo(probe.profiles_summary_loaded) << L"\n"
                    << L"profile count: " << probe.profile_count << L"\n"
                    << L"active profile: "
                    << (probe.active_display_name.empty() ? std::wstring(L"[not recorded]")
                                                           : probe.active_display_name);
            if (!probe.summary_error.empty()) {
                message << L"\nsummary error: " << probe.summary_error;
            }
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::DiffFirstProfile: {
            action_result.title = L"Codex profile diff";
            std::wstring target;
            if (!SelectedProfileTarget(
                    probe, options.target_display_name, &target, &action_result.message)) {
                break;
            }

            CodexProfileDiffRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.codex_home = probe.paths.codex_home;
            request.target_display_name = target;

            CodexProfileDiffService service;
            CodexProfileDiffResult result;
            if (!service.BuildProfileDiffSummary(request, &result)) {
                action_result.message = result.error_message;
                break;
            }

            std::wstringstream message;
            message << L"Target: " << target;
            for (const CodexProfileFileDiffSummary& file : result.files) {
                message << L"\n" << FileDiffLine(file);
            }
            action_result.ok = true;
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::BackupCurrentFolder: {
            action_result.title = L"Codex current folder backup";

            CodexProfileBackupService service;
            CodexCurrentFolderBackupResult result;
            if (!service.CreateCurrentFolderBackup(probe.paths.codex_home,
                                                   probe.paths.backups_dir,
                                                   probe.active_display_name,
                                                   &result)) {
                action_result.message = result.error_message;
                break;
            }

            action_result.ok = true;
            action_result.changed_state = !result.skipped_no_files;
            if (result.skipped_no_files) {
                action_result.message = L"No config.toml or auth.json exists in the current Codex folder.";
            } else {
                action_result.message = L"Backup created:\n" + result.backup_path;
            }
            break;
        }

        case CodexProfileUiAction::ApplyFirstProfile: {
            action_result.title = L"Codex profile applied";
            action_result.sensitive_write = true;

            std::wstring target;
            if (!SelectedProfileTarget(
                    probe, options.target_display_name, &target, &action_result.message)) {
                break;
            }

            CodexProfileSwitchRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.active_json_path = probe.paths.active_json;
            request.backup_directory = probe.paths.backups_dir;
            request.codex_home = probe.paths.codex_home;
            request.target_display_name = target;
            request.previous_active_display_name = probe.active_display_name;

            CodexProfileSwitchService service;
            CodexProfileSwitchResult result;
            if (!service.ApplyProfileByDisplayName(request, &result)) {
                action_result.message = result.error_message;
                break;
            }

            action_result.ok = true;
            action_result.changed_state = true;
            std::wstringstream message;
            message << L"Applied: " << target;
            if (result.backup_skipped_no_files) {
                message << L"\nNo previous config/auth files existed, so backup was skipped.";
            } else {
                message << L"\nBackup created before write:\n" << result.backup_path;
            }
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::ImportCurrentFolder: {
            action_result.title = L"Codex current folder imported";
            action_result.sensitive_write = true;

            CodexProfileImportRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.codex_home = probe.paths.codex_home;
            request.display_name = L"Native current " + TimestampForGeneratedName();

            CodexProfileImportService service;
            CodexProfileImportResult result;
            if (!service.ImportCurrentFolderProfile(request, &result)) {
                action_result.message = result.error_message;
                break;
            }

            action_result.ok = true;
            action_result.changed_state = true;
            action_result.message =
                L"Imported current config.toml/auth.json as:\n" + request.display_name +
                (result.created_new_library ? L"\nA new profiles library was created."
                                            : L"\nThe profile was appended to the existing library.");
            break;
        }

        case CodexProfileUiAction::RestoreLatestBackup: {
            action_result.title = L"Codex latest backup restored";
            action_result.sensitive_write = true;

            CodexProfileBackupService service;
            CodexBackupRestoreResult result;
            if (!service.RestoreLatestBackup(probe.paths.codex_home,
                                             probe.paths.backups_dir,
                                             &result)) {
                action_result.message = result.error_message;
                break;
            }

            action_result.ok = true;
            action_result.changed_state = true;
            std::wstringstream message;
            message << L"Restored from:\n" << result.backup_path << L"\nconfig.toml: "
                    << YesNo(result.restored_config) << L"\nauth.json: "
                    << YesNo(result.restored_auth);
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::ExportFirstProfileFiles: {
            action_result.title = L"Codex profile files exported";
            action_result.sensitive_write = true;

            std::wstring target;
            if (!SelectedProfileTarget(
                    probe, options.target_display_name, &target, &action_result.message)) {
                break;
            }
            if (SameDirectoryPath(options.output_directory, probe.paths.codex_home)) {
                action_result.message =
                    L"Exporting profile files directly into the current Codex folder is blocked. Choose another folder to avoid overwriting the active config/auth files without the apply-and-backup flow.";
                break;
            }

            CodexProfileExportRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.output_directory = options.output_directory;
            request.target_display_name = target;

            CodexProfileExportService service;
            CodexProfileExportResult result;
            if (!service.ExportProfileByDisplayName(request, &result)) {
                action_result.message = result.error_message;
                break;
            }

            action_result.ok = true;
            std::wstringstream message;
            message << L"Exported: " << target << L"\nconfig.toml:\n"
                    << result.config_path << L"\nauth.json:\n" << result.auth_path;
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::RenameFirstProfile: {
            action_result.title = L"Codex profile renamed";
            action_result.sensitive_write = true;

            std::wstring target;
            if (!SelectedProfileTarget(
                    probe, options.target_display_name, &target, &action_result.message)) {
                break;
            }
            if (SameDisplayName(target, options.new_display_name)) {
                action_result.message = L"The new display name matches the selected profile name.";
                break;
            }

            CodexProfileEditRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.active_json_path = probe.paths.active_json;
            request.active_display_name = probe.active_display_name;
            request.target_display_name = target;
            request.update_display_name = true;
            request.new_display_name = options.new_display_name;

            CodexProfileEditService service;
            CodexProfileEditResult result;
            if (!service.UpdateProfileMetadata(request, &result)) {
                action_result.message = result.error_message;
                action_result.changed_state = result.profiles_json_updated;
                break;
            }

            action_result.ok = true;
            action_result.changed_state = true;
            std::wstringstream message;
            message << L"Renamed profile:\n" << target << L"\nNew name:\n"
                    << result.saved_display_name;
            if (result.active_json_updated) {
                message << L"\nActive profile marker updated in active.json.";
            }
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::EditFirstProfileNote:
        case CodexProfileUiAction::EditFirstProfileRemark:
        case CodexProfileUiAction::EditFirstProfileTags: {
            action_result.title = L"Codex profile metadata updated";
            action_result.sensitive_write = true;

            std::wstring target;
            if (!SelectedProfileTarget(
                    probe, options.target_display_name, &target, &action_result.message)) {
                break;
            }

            CodexProfileEditRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.target_display_name = target;

            const wchar_t* field_name = L"metadata";
            switch (action) {
                case CodexProfileUiAction::EditFirstProfileNote:
                    request.update_note = true;
                    request.note = options.note;
                    field_name = L"note";
                    break;
                case CodexProfileUiAction::EditFirstProfileRemark:
                    request.update_remark = true;
                    request.remark = options.remark;
                    field_name = L"remark";
                    break;
                case CodexProfileUiAction::EditFirstProfileTags:
                    request.update_tags = true;
                    request.tags = options.tags;
                    field_name = L"tags";
                    break;
                default:
                    break;
            }

            CodexProfileEditService service;
            CodexProfileEditResult result;
            if (!service.UpdateProfileMetadata(request, &result)) {
                action_result.message = result.error_message;
                break;
            }

            action_result.ok = true;
            action_result.changed_state = true;
            std::wstringstream message;
            message << L"Updated " << field_name << L" for:\n" << target;
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::ExportBox: {
            action_result.title = L"Codex .codexbox exported";
            action_result.sensitive_write = true;

            CodexProfileBoxExportRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.output_path = options.box_path;
            request.password = options.password;

            CodexProfileBoxService service;
            CodexProfileBoxExportResult result;
            if (!service.ExportBox(request, &result)) {
                action_result.message = result.error_message;
                ClearWideString(&request.password);
                break;
            }

            ClearWideString(&request.password);
            action_result.ok = true;
            std::wstringstream message;
            message << L"Exported profiles: " << result.exported_count << L"\nPackage:\n"
                    << options.box_path;
            action_result.message = message.str();
            break;
        }

        case CodexProfileUiAction::ImportBox: {
            action_result.title = L"Codex .codexbox imported";
            action_result.sensitive_write = true;

            CodexProfileBoxImportRequest request;
            request.profiles_json_path = probe.paths.profiles_json;
            request.box_path = options.box_path;
            request.password = options.password;
            request.conflict_policy = options.box_conflict_policy;

            CodexProfileBoxService service;
            CodexProfileBoxImportResult result;
            if (!service.ImportBox(request, &result)) {
                action_result.message = result.error_message;
                ClearWideString(&request.password);
                break;
            }

            ClearWideString(&request.password);
            action_result.ok = true;
            action_result.changed_state =
                result.imported_count > 0 || result.replaced_count > 0 || result.renamed_count > 0;
            std::wstringstream message;
            message << L"Imported: " << result.imported_count << L"\nSkipped: "
                    << result.skipped_count << L"\nRenamed conflicts: "
                    << result.renamed_count << L"\nReplaced: " << result.replaced_count;
            if (result.created_new_library) {
                message << L"\nA new profiles library was created.";
            }
            action_result.message = message.str();
            break;
        }
    }

    if (logger_ != nullptr) {
        if (action_result.ok) {
            logger_->Info(L"Native Codex explicit menu action completed.");
        } else {
            logger_->Warn(L"Native Codex explicit menu action failed before completion.");
        }
    }
    return action_result;
}

void CodexProfileModule::LoadProfileSummaries(CodexProfileProbe* probe) const {
    if (probe == nullptr) {
        return;
    }

    if (probe->active_json_exists) {
        std::string active_json;
        std::wstring active_error;
        if (FileSystem::ReadUtf8File(probe->paths.active_json, &active_json, &active_error)) {
            probe->active_display_name =
                Utf8ToWide(ExtractFirstJsonStringValue(active_json, "ActiveDisplayName"));
        }
    }

    if (!probe->profiles_json_exists) {
        return;
    }

    std::string protected_text;
    std::wstring read_error;
    if (!FileSystem::ReadUtf8File(probe->paths.profiles_json, &protected_text, &read_error)) {
        probe->summary_error = read_error;
        return;
    }

    SecretStoreDpapi secret_store;
    std::string json;
    std::wstring dpapi_error;
    if (!secret_store.UnprotectBase64ToUtf8(protected_text, &json, &dpapi_error)) {
        probe->summary_error = dpapi_error;
        return;
    }

    std::vector<std::string> profile_objects = ExtractJsonObjectArray(json, "items");

    probe->profiles_summary_loaded = true;
    probe->profile_count = profile_objects.size();
    const size_t display_count = std::min<size_t>(profile_objects.size(), 5);
    for (size_t index = 0; index < profile_objects.size(); ++index) {
        std::string name = ExtractFirstJsonStringValue(profile_objects[index], "DisplayName");
        if (name.empty()) {
            name = ExtractFirstJsonStringValue(profile_objects[index], "Name");
        }
        const std::wstring wide_name = Utf8ToWide(name);
        if (probe->first_profile_display_name.empty()) {
            if (!wide_name.empty()) {
                probe->first_profile_display_name = wide_name;
            }
        }
        if (!wide_name.empty()) {
            probe->profile_display_names.push_back(wide_name);
        }
        if (index < display_count) {
            probe->profile_summaries.push_back(
                SummaryLine(index,
                            name,
                            ExtractFirstJsonStringValue(profile_objects[index], "Status"),
                            ExtractFirstJsonStringValue(profile_objects[index], "LastImportedAt"),
                            probe->active_display_name));
        }
    }
    if (profile_objects.size() > display_count) {
        probe->profile_summaries.push_back(L"additional profiles hidden in summary view.");
    }

    for (std::string& profile_object : profile_objects) {
        SecureZeroMemory(profile_object.data(), profile_object.size());
    }
    SecureZeroMemory(json.data(), json.size());
}

CodexProfilePaths CodexProfileModule::ResolvePaths() const {
    CodexProfilePaths paths;

    const std::wstring local_app_data = EnvironmentVariable(L"LOCALAPPDATA");
    const std::wstring user_profile = EnvironmentVariable(L"USERPROFILE");

    paths.library_dir = FileSystem::JoinPath(local_app_data, L"MyTools\\Codex");
    paths.profiles_json = FileSystem::JoinPath(paths.library_dir, L"profiles.json");
    paths.active_json = FileSystem::JoinPath(paths.library_dir, L"active.json");
    paths.backups_dir = FileSystem::JoinPath(paths.library_dir, L"Backups");

    paths.codex_home = FileSystem::JoinPath(user_profile, L".codex");
    paths.config_toml = FileSystem::JoinPath(paths.codex_home, L"config.toml");
    paths.auth_json = FileSystem::JoinPath(paths.codex_home, L"auth.json");
    return paths;
}

}  // namespace mytools
