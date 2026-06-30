#pragma once

#include <string>

namespace mytools {

struct CodexProfileExportRequest {
    std::wstring profiles_json_path;
    std::wstring output_directory;
    std::wstring target_display_name;
};

struct CodexProfileExportResult {
    std::wstring config_path;
    std::wstring auth_path;
    std::wstring error_message;
};

class CodexProfileExportService {
public:
    bool ExportProfileByDisplayName(const CodexProfileExportRequest& request,
                                    CodexProfileExportResult* result) const;
};

}  // namespace mytools
