#include "services/file_system.h"

#include <string>
#include <vector>

#include <windows.h>

namespace mytools {
namespace {

bool IsSlash(wchar_t value) {
    return value == L'\\' || value == L'/';
}

bool IsDriveRoot(const std::wstring& path) {
    return path.size() == 3 && path[1] == L':' && IsSlash(path[2]);
}

bool CreateDirectoryRecursive(const std::wstring& path, std::wstring* error_message) {
    if (path.empty() || path == L"." || IsDriveRoot(path)) {
        return true;
    }

    const DWORD attributes = GetFileAttributesW(path.c_str());
    if (attributes != INVALID_FILE_ATTRIBUTES) {
        if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
            return true;
        }
        if (error_message != nullptr) {
            *error_message = L"Path exists but is not a directory: " + path;
        }
        return false;
    }

    const std::wstring parent = FileSystem::DirectoryName(path);
    if (!parent.empty() && parent != path && !CreateDirectoryRecursive(parent, error_message)) {
        return false;
    }

    if (CreateDirectoryW(path.c_str(), nullptr) || GetLastError() == ERROR_ALREADY_EXISTS) {
        return true;
    }

    if (error_message != nullptr) {
        *error_message = FormatLastErrorMessage(L"CreateDirectoryW");
    }
    return false;
}

bool WriteAll(HANDLE file, const char* data, size_t size, std::wstring* error_message) {
    size_t written_total = 0;
    while (written_total < size) {
        const size_t remaining = size - written_total;
        const DWORD chunk = remaining > MAXDWORD ? MAXDWORD : static_cast<DWORD>(remaining);
        DWORD written = 0;
        if (!WriteFile(file, data + written_total, chunk, &written, nullptr)) {
            if (error_message != nullptr) {
                *error_message = FormatLastErrorMessage(L"WriteFile");
            }
            return false;
        }
        if (written == 0) {
            if (error_message != nullptr) {
                *error_message = L"WriteFile returned zero bytes written.";
            }
            return false;
        }
        written_total += written;
    }
    return true;
}

}  // namespace

std::wstring FormatLastErrorMessage(const wchar_t* operation) {
    const DWORD error = GetLastError();
    wchar_t* message = nullptr;
    const DWORD length = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
                                            FORMAT_MESSAGE_IGNORE_INSERTS,
                                        nullptr,
                                        error,
                                        0,
                                        reinterpret_cast<LPWSTR>(&message),
                                        0,
                                        nullptr);

    std::wstring result = operation != nullptr ? operation : L"Windows API";
    result += L" failed";
    if (length > 0 && message != nullptr) {
        result += L": ";
        result += message;
        LocalFree(message);
    } else {
        result += L".";
    }
    return result;
}

std::wstring FileSystem::JoinPath(const std::wstring& directory, const std::wstring& file_name) {
    if (directory.empty()) {
        return file_name;
    }
    if (file_name.empty()) {
        return directory;
    }
    if (IsSlash(directory.back())) {
        return directory + file_name;
    }
    return directory + L"\\" + file_name;
}

std::wstring FileSystem::DirectoryName(const std::wstring& path) {
    const size_t slash = path.find_last_of(L"\\/");
    if (slash == std::wstring::npos) {
        return {};
    }
    if (slash == 0) {
        return path.substr(0, 1);
    }
    if (slash == 2 && path.size() > 2 && path[1] == L':') {
        return path.substr(0, 3);
    }
    return path.substr(0, slash);
}

bool FileSystem::Exists(const std::wstring& path) {
    return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;
}

bool FileSystem::EnsureDirectory(const std::wstring& path, std::wstring* error_message) {
    return CreateDirectoryRecursive(path, error_message);
}

bool FileSystem::ReadFileBytes(const std::wstring& path,
                               std::vector<unsigned char>* content,
                               std::wstring* error_message) {
    if (content == nullptr) {
        if (error_message != nullptr) {
            *error_message = L"ReadFileBytes requires a content output pointer.";
        }
        return false;
    }

    HANDLE file = CreateFileW(path.c_str(),
                              GENERIC_READ,
                              FILE_SHARE_READ,
                              nullptr,
                              OPEN_EXISTING,
                              FILE_ATTRIBUTE_NORMAL,
                              nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        if (GetLastError() == ERROR_FILE_NOT_FOUND || GetLastError() == ERROR_PATH_NOT_FOUND) {
            content->clear();
            return true;
        }
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"CreateFileW");
        }
        return false;
    }

    LARGE_INTEGER file_size{};
    if (!GetFileSizeEx(file, &file_size) || file_size.QuadPart < 0 ||
        file_size.QuadPart > static_cast<LONGLONG>(32 * 1024 * 1024)) {
        CloseHandle(file);
        if (error_message != nullptr) {
            *error_message = L"ReadFileBytes only accepts files up to 32 MB.";
        }
        return false;
    }

    content->assign(static_cast<size_t>(file_size.QuadPart), 0);
    if (content->empty()) {
        CloseHandle(file);
        return true;
    }

    DWORD read = 0;
    const BOOL ok = ReadFile(file,
                             content->data(),
                             static_cast<DWORD>(content->size()),
                             &read,
                             nullptr);
    CloseHandle(file);

    if (!ok || read != content->size()) {
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"ReadFile");
        }
        return false;
    }
    return true;
}

bool FileSystem::ReadUtf8File(const std::wstring& path,
                              std::string* content,
                              std::wstring* error_message) {
    if (content == nullptr) {
        if (error_message != nullptr) {
            *error_message = L"ReadUtf8File requires a content output pointer.";
        }
        return false;
    }

    HANDLE file = CreateFileW(path.c_str(),
                              GENERIC_READ,
                              FILE_SHARE_READ,
                              nullptr,
                              OPEN_EXISTING,
                              FILE_ATTRIBUTE_NORMAL,
                              nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        if (GetLastError() == ERROR_FILE_NOT_FOUND || GetLastError() == ERROR_PATH_NOT_FOUND) {
            content->clear();
            return true;
        }
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"CreateFileW");
        }
        return false;
    }

    LARGE_INTEGER file_size{};
    if (!GetFileSizeEx(file, &file_size) || file_size.QuadPart < 0 ||
        file_size.QuadPart > static_cast<LONGLONG>(32 * 1024 * 1024)) {
        CloseHandle(file);
        if (error_message != nullptr) {
            *error_message = L"ReadUtf8File only accepts files up to 32 MB.";
        }
        return false;
    }

    content->assign(static_cast<size_t>(file_size.QuadPart), '\0');
    if (content->empty()) {
        CloseHandle(file);
        return true;
    }

    DWORD read = 0;
    const BOOL ok = ReadFile(file,
                             content->data(),
                             static_cast<DWORD>(content->size()),
                             &read,
                             nullptr);
    CloseHandle(file);

    if (!ok || read != content->size()) {
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"ReadFile");
        }
        return false;
    }
    return true;
}

bool FileSystem::WriteFileBytesAtomic(const std::wstring& path,
                                      const std::vector<unsigned char>& content,
                                      std::wstring* error_message) {
    const std::wstring directory = DirectoryName(path);
    if (!directory.empty() && !EnsureDirectory(directory, error_message)) {
        return false;
    }

    const std::wstring temp_path = path + L".tmp";
    HANDLE file = CreateFileW(temp_path.c_str(),
                              GENERIC_WRITE,
                              0,
                              nullptr,
                              CREATE_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL,
                              nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"CreateFileW");
        }
        return false;
    }

    const char* data =
        content.empty() ? nullptr : reinterpret_cast<const char*>(content.data());
    const bool wrote = WriteAll(file, data, content.size(), error_message);
    const bool flushed = FlushFileBuffers(file) != FALSE;
    CloseHandle(file);

    if (!wrote || !flushed) {
        DeleteFileW(temp_path.c_str());
        if (!flushed && error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"FlushFileBuffers");
        }
        return false;
    }

    if (!MoveFileExW(temp_path.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        DeleteFileW(temp_path.c_str());
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"MoveFileExW");
        }
        return false;
    }

    return true;
}

bool FileSystem::WriteUtf8FileAtomic(const std::wstring& path,
                                     const std::string& content,
                                     std::wstring* error_message) {
    const std::wstring directory = DirectoryName(path);
    if (!directory.empty() && !EnsureDirectory(directory, error_message)) {
        return false;
    }

    const std::wstring temp_path = path + L".tmp";
    HANDLE file = CreateFileW(temp_path.c_str(),
                              GENERIC_WRITE,
                              0,
                              nullptr,
                              CREATE_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL,
                              nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"CreateFileW");
        }
        return false;
    }

    const bool wrote = WriteAll(file, content.data(), content.size(), error_message);
    const bool flushed = FlushFileBuffers(file) != FALSE;
    CloseHandle(file);

    if (!wrote || !flushed) {
        DeleteFileW(temp_path.c_str());
        if (!flushed && error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"FlushFileBuffers");
        }
        return false;
    }

    if (!MoveFileExW(temp_path.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        DeleteFileW(temp_path.c_str());
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"MoveFileExW");
        }
        return false;
    }

    return true;
}

}  // namespace mytools
