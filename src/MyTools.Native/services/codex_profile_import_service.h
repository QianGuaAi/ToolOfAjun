#pragma once

#include <string>

namespace mytools {

struct CodexProfileImportRequest {
    std::wstring profiles_json_path;
    std::wstring codex_home;
    std::wstring display_name;
};

struct CodexProfileImportResult {
    bool created_new_library = false;
    std::wstring error_message;
};

class CodexProfileImportService {
public:
    bool ImportCurrentFolderProfile(const CodexProfileImportRequest& request,
                                    CodexProfileImportResult* result) const;
};

}  // namespace mytools
