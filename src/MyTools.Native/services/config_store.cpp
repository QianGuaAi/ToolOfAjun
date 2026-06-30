#include "services/config_store.h"

#include "ccore/json/mt_json_schema.h"
#include "services/file_system.h"

#include <utility>

namespace mytools {

ConfigStore::ConfigStore(std::wstring root_directory) : root_directory_(std::move(root_directory)) {}

ConfigLoadResult ConfigStore::LoadRawJson(const std::wstring& relative_path) const {
    ConfigLoadResult result;
    const std::wstring path = ResolvePath(relative_path);
    if (!FileSystem::Exists(path)) {
        return result;
    }

    result.found = true;
    if (!FileSystem::ReadUtf8File(path, &result.json, &result.error_message)) {
        return result;
    }

    if (!mt_json_has_schema_version(result.json.data(), result.json.size())) {
        result.error_message = L"Config file is missing schema_version: " + relative_path;
    }
    return result;
}

bool ConfigStore::SaveRawJson(const std::wstring& relative_path,
                              const std::string& json,
                              std::wstring* error_message) const {
    if (!mt_json_has_schema_version(json.data(), json.size())) {
        if (error_message != nullptr) {
            *error_message = L"Config JSON must contain schema_version before saving.";
        }
        return false;
    }

    return FileSystem::WriteUtf8FileAtomic(ResolvePath(relative_path), json, error_message);
}

std::wstring ConfigStore::ResolvePath(const std::wstring& relative_path) const {
    return FileSystem::JoinPath(root_directory_, relative_path);
}

}  // namespace mytools
