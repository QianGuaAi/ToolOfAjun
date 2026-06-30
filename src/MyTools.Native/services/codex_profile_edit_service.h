#pragma once

#include <string>

namespace mytools {

struct CodexProfileEditRequest {
    std::wstring profiles_json_path;
    std::wstring active_json_path;
    std::wstring active_display_name;
    std::wstring target_display_name;

    bool update_display_name = false;
    std::wstring new_display_name;

    bool update_note = false;
    std::wstring note;

    bool update_remark = false;
    std::wstring remark;

    bool update_tags = false;
    std::wstring tags;
};

struct CodexProfileEditResult {
    bool profile_found = false;
    bool display_name_changed = false;
    bool profiles_json_updated = false;
    bool active_json_updated = false;
    std::wstring saved_display_name;
    std::wstring error_message;
};

class CodexProfileEditService {
public:
    bool UpdateProfileMetadata(const CodexProfileEditRequest& request,
                               CodexProfileEditResult* result) const;
};

}  // namespace mytools
