#pragma once

#include <string>
#include <vector>

#include "modules/module_registry.h"
#include "services/codex_profile_box_service.h"
#include "services/logger.h"

namespace mytools {

struct CodexProfilePaths {
    std::wstring library_dir;
    std::wstring profiles_json;
    std::wstring active_json;
    std::wstring backups_dir;
    std::wstring codex_home;
    std::wstring config_toml;
    std::wstring auth_json;
};

struct CodexProfileProbe {
    CodexProfilePaths paths;
    bool profiles_json_exists = false;
    bool active_json_exists = false;
    bool config_toml_exists = false;
    bool auth_json_exists = false;
    bool profiles_summary_loaded = false;
    size_t profile_count = 0;
    std::wstring active_display_name;
    std::wstring first_profile_display_name;
    std::wstring summary_error;
    std::vector<std::wstring> profile_summaries;
};

enum class CodexProfileUiAction {
    Refresh,
    DiffFirstProfile,
    BackupCurrentFolder,
    ApplyFirstProfile,
    ImportCurrentFolder,
    RestoreLatestBackup,
    ExportFirstProfileFiles,
    RenameFirstProfile,
    EditFirstProfileNote,
    EditFirstProfileRemark,
    EditFirstProfileTags,
    ExportBox,
    ImportBox
};

struct CodexProfileActionOptions {
    std::wstring box_path;
    std::wstring output_directory;
    std::wstring new_display_name;
    std::wstring note;
    std::wstring remark;
    std::wstring tags;
    std::wstring password;
    CodexProfileBoxConflictPolicy box_conflict_policy = CodexProfileBoxConflictPolicy::Rename;
};

struct CodexProfileActionResult {
    bool ok = false;
    bool changed_state = false;
    bool sensitive_write = false;
    std::wstring title;
    std::wstring message;
};

class CodexProfileModule {
public:
    explicit CodexProfileModule(Logger* logger);

    ModuleInfo BuildModuleInfo() const;
    CodexProfileProbe ProbeLocalState() const;
    CodexProfileActionResult RunUiAction(
        CodexProfileUiAction action,
        const CodexProfileActionOptions& options = CodexProfileActionOptions{}) const;

private:
    CodexProfilePaths ResolvePaths() const;
    void LoadProfileSummaries(CodexProfileProbe* probe) const;

    Logger* logger_ = nullptr;
};

}  // namespace mytools
