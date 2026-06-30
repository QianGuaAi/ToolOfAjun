#pragma once

#include <string>
#include <vector>

namespace mytools {

class FileSystem {
public:
    static std::wstring JoinPath(const std::wstring& directory, const std::wstring& file_name);
    static std::wstring DirectoryName(const std::wstring& path);
    static bool Exists(const std::wstring& path);
    static bool EnsureDirectory(const std::wstring& path, std::wstring* error_message);
    static bool ReadFileBytes(const std::wstring& path,
                              std::vector<unsigned char>* content,
                              std::wstring* error_message);
    static bool ReadUtf8File(const std::wstring& path, std::string* content, std::wstring* error_message);
    static bool WriteFileBytesAtomic(const std::wstring& path,
                                     const std::vector<unsigned char>& content,
                                     std::wstring* error_message);
    static bool WriteUtf8FileAtomic(const std::wstring& path,
                                    const std::string& content,
                                    std::wstring* error_message);
};

std::wstring FormatLastErrorMessage(const wchar_t* operation);

}  // namespace mytools
