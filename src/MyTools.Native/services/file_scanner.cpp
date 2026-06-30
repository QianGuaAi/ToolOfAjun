#include "services/file_scanner.h"

#include <cwchar>
#include <utility>
#include <vector>

#include <windows.h>

#include "services/file_system.h"

namespace mytools {
namespace {

std::wstring SearchPattern(const std::wstring& directory) {
    return FileSystem::JoinPath(directory, L"*");
}

bool IsDotDirectory(const wchar_t* name) {
    return wcscmp(name, L".") == 0 || wcscmp(name, L"..") == 0;
}

}  // namespace

FileScanResult FileScanner::ScanFiles(const FileScanOptions& options,
                                      const ScanCancellation& cancellation) const {
    FileScanResult result;
    if (options.root_directory.empty()) {
        result.error_message = L"FileScanner requires a root directory.";
        return result;
    }

    std::vector<std::wstring> directories;
    directories.push_back(options.root_directory);

    while (!directories.empty()) {
        if (cancellation.IsCancellationRequested()) {
            result.cancelled = true;
            return result;
        }

        const std::wstring directory = std::move(directories.back());
        directories.pop_back();

        WIN32_FIND_DATAW find_data{};
        HANDLE find = FindFirstFileW(SearchPattern(directory).c_str(), &find_data);
        if (find == INVALID_HANDLE_VALUE) {
            const DWORD error = GetLastError();
            if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND) {
                result.error_message = FormatLastErrorMessage(L"FindFirstFileW");
                return result;
            }
            continue;
        }

        do {
            if (cancellation.IsCancellationRequested()) {
                result.cancelled = true;
                FindClose(find);
                return result;
            }

            if ((find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
                if (options.recursive && !IsDotDirectory(find_data.cFileName) &&
                    (find_data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0) {
                    directories.push_back(FileSystem::JoinPath(directory, find_data.cFileName));
                }
                continue;
            }

            result.files.push_back(FileSystem::JoinPath(directory, find_data.cFileName));
            if (result.files.size() >= options.max_results) {
                FindClose(find);
                return result;
            }
        } while (FindNextFileW(find, &find_data));

        const DWORD last_error = GetLastError();
        FindClose(find);
        if (last_error != ERROR_NO_MORE_FILES) {
            result.error_message = FormatLastErrorMessage(L"FindNextFileW");
            return result;
        }
    }

    return result;
}

}  // namespace mytools
