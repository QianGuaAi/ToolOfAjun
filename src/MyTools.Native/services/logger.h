#pragma once

#include <mutex>
#include <string>

namespace mytools {

class Logger {
public:
    explicit Logger(std::wstring directory);

    void Info(const std::wstring& message);
    void Warn(const std::wstring& message);
    void Error(const std::wstring& message);
    void Startup(const std::wstring& message);
    void Crash(const std::wstring& message);

    const std::wstring& log_path() const noexcept { return log_path_; }
    const std::wstring& startup_log_path() const noexcept { return startup_log_path_; }
    const std::wstring& crash_log_path() const noexcept { return crash_log_path_; }

private:
    void WriteLine(const std::wstring& path, const std::wstring& level, const std::wstring& message);

    std::mutex mutex_;
    std::wstring log_path_;
    std::wstring startup_log_path_;
    std::wstring crash_log_path_;
};

}  // namespace mytools
