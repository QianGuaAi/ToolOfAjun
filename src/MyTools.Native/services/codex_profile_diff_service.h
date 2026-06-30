#pragma once

#include <string>
#include <vector>

namespace mytools {

struct CodexProfileDiffRequest {
    std::wstring profiles_json_path;
    std::wstring codex_home;
    std::wstring target_display_name;
};

struct CodexProfileFileDiffSummary {
    std::wstring file_name;
    bool profile_available = false;
    bool current_available = false;
    bool same = false;
    unsigned long long profile_size = 0;
    unsigned long long current_size = 0;
    unsigned long long profile_line_count = 0;
    unsigned long long current_line_count = 0;
    std::wstring profile_sha256_hex;
    std::wstring current_sha256_hex;
};

struct CodexProfileDiffResult {
    std::vector<CodexProfileFileDiffSummary> files;
    std::wstring error_message;
};

class CodexProfileDiffService {
public:
    bool BuildProfileDiffSummary(const CodexProfileDiffRequest& request,
                                 CodexProfileDiffResult* result) const;
};

}  // namespace mytools
