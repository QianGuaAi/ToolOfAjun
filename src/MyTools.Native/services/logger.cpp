#include "services/logger.h"

#include <sstream>
#include <string>
#include <vector>

#include <windows.h>

namespace mytools {
namespace {

std::wstring JoinPath(const std::wstring& directory, const std::wstring& file_name) {
    if (directory.empty()) {
        return file_name;
    }
    const wchar_t last = directory.back();
    if (last == L'\\' || last == L'/') {
        return directory + file_name;
    }
    return directory + L"\\" + file_name;
}

std::wstring NowText() {
    SYSTEMTIME local_time{};
    GetLocalTime(&local_time);

    wchar_t buffer[64]{};
    swprintf_s(buffer,
               L"%04hu-%02hu-%02hu %02hu:%02hu:%02hu.%03hu",
               local_time.wYear,
               local_time.wMonth,
               local_time.wDay,
               local_time.wHour,
               local_time.wMinute,
               local_time.wSecond,
               local_time.wMilliseconds);
    return buffer;
}

std::string ToUtf8(const std::wstring& text) {
    if (text.empty()) {
        return {};
    }

    const int length = WideCharToMultiByte(
        CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (length <= 0) {
        return {};
    }

    std::string result(static_cast<size_t>(length), '\0');
    WideCharToMultiByte(
        CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), result.data(), length, nullptr, nullptr);
    return result;
}

void AppendUtf8Line(const std::wstring& path, const std::wstring& line) {
    HANDLE file = CreateFileW(path.c_str(),
                              FILE_APPEND_DATA,
                              FILE_SHARE_READ,
                              nullptr,
                              OPEN_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL,
                              nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return;
    }

    const std::string data = ToUtf8(line + L"\r\n");
    DWORD written = 0;
    WriteFile(file, data.data(), static_cast<DWORD>(data.size()), &written, nullptr);
    CloseHandle(file);
}

}  // namespace

Logger::Logger(std::wstring directory)
    : log_path_(JoinPath(directory, L"MyToolsNative.log")),
      startup_log_path_(JoinPath(directory, L"MyToolsNative.startup.log")),
      crash_log_path_(JoinPath(directory, L"MyToolsNative.crash.log")) {}

void Logger::Info(const std::wstring& message) {
    WriteLine(log_path_, L"INFO", message);
}

void Logger::Warn(const std::wstring& message) {
    WriteLine(log_path_, L"WARN", message);
}

void Logger::Error(const std::wstring& message) {
    WriteLine(log_path_, L"ERROR", message);
}

void Logger::Startup(const std::wstring& message) {
    WriteLine(startup_log_path_, L"STARTUP", message);
}

void Logger::Crash(const std::wstring& message) {
    WriteLine(crash_log_path_, L"CRASH", message);
}

void Logger::WriteLine(const std::wstring& path,
                       const std::wstring& level,
                       const std::wstring& message) {
    std::lock_guard<std::mutex> lock(mutex_);

    std::wstringstream line;
    line << NowText() << L" [" << level << L"] " << message;
    AppendUtf8Line(path, line.str());
}

}  // namespace mytools
