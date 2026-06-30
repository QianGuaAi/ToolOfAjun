#pragma once

#include <string>

namespace mytools {

struct CodexCurrentFolderBackupResult {
    bool skipped_no_files = false;
    std::wstring backup_path;
    std::wstring error_message;
};

struct CodexBackupRestoreResult {
    bool restored_config = false;
    bool restored_auth = false;
    std::wstring backup_path;
    std::wstring error_message;
};

class CodexProfileBackupService {
public:
    bool CreateCurrentFolderBackup(const std::wstring& codex_home,
                                   const std::wstring& backup_directory,
                                   const std::wstring& active_display_name,
                                   CodexCurrentFolderBackupResult* result) const;
    bool RestoreLatestBackup(const std::wstring& codex_home,
                             const std::wstring& backup_directory,
                             CodexBackupRestoreResult* result) const;
};

}  // namespace mytools
