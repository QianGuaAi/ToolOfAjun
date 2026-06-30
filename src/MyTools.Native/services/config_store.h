#pragma once

#include <string>

namespace mytools {

struct ConfigLoadResult {
    bool found = false;
    std::string json;
    std::wstring error_message;
};

class ConfigStore {
public:
    explicit ConfigStore(std::wstring root_directory);

    ConfigLoadResult LoadRawJson(const std::wstring& relative_path) const;
    bool SaveRawJson(const std::wstring& relative_path,
                     const std::string& json,
                     std::wstring* error_message) const;
    std::wstring ResolvePath(const std::wstring& relative_path) const;

private:
    std::wstring root_directory_;
};

}  // namespace mytools
