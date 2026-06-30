#pragma once

#include <string>

namespace mytools {

struct CodexProfileSwitchRequest {
    std::wstring profiles_json_path;
    std::wstring active_json_path;
    std::wstring backup_directory;
    std::wstring codex_home;
    std::wstring target_display_name;
    std::wstring previous_active_display_name;
};

struct CodexProfileSwitchResult {
    bool backup_skipped_no_files = false;
    std::wstring backup_path;
    std::wstring error_message;
};

class CodexProfileSwitchService {
public:
    bool ApplyProfileByDisplayName(const CodexProfileSwitchRequest& request,
                                   CodexProfileSwitchResult* result) const;
};

}  // namespace mytools
