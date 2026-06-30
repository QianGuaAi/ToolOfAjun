#pragma once

#include <cstddef>
#include <string>

namespace mytools {

enum class CodexProfileBoxConflictPolicy {
    Skip = 0,
    Replace = 1,
    Rename = 2
};

struct CodexProfileBoxExportRequest {
    std::wstring profiles_json_path;
    std::wstring output_path;
    std::wstring password;
};

struct CodexProfileBoxExportResult {
    size_t exported_count = 0;
    std::wstring error_message;
};

struct CodexProfileBoxImportRequest {
    std::wstring profiles_json_path;
    std::wstring box_path;
    std::wstring password;
    CodexProfileBoxConflictPolicy conflict_policy = CodexProfileBoxConflictPolicy::Rename;
};

struct CodexProfileBoxImportResult {
    size_t imported_count = 0;
    size_t skipped_count = 0;
    size_t replaced_count = 0;
    size_t renamed_count = 0;
    bool created_new_library = false;
    std::wstring error_message;
};

class CodexProfileBoxService {
public:
    bool ExportBox(const CodexProfileBoxExportRequest& request,
                   CodexProfileBoxExportResult* result) const;
    bool ImportBox(const CodexProfileBoxImportRequest& request,
                   CodexProfileBoxImportResult* result) const;
};

}  // namespace mytools
