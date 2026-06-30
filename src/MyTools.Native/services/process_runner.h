#pragma once

#include <string>

#include <windows.h>

namespace mytools {

struct ProcessStartOptions {
    std::wstring file_path;
    std::wstring arguments;
    std::wstring working_directory;
    bool create_no_window = true;
};

class ProcessHandle {
public:
    ProcessHandle() = default;
    ~ProcessHandle();

    ProcessHandle(const ProcessHandle&) = delete;
    ProcessHandle& operator=(const ProcessHandle&) = delete;
    ProcessHandle(ProcessHandle&& other) noexcept;
    ProcessHandle& operator=(ProcessHandle&& other) noexcept;

    bool IsValid() const noexcept { return process_ != nullptr; }
    DWORD process_id() const noexcept { return process_id_; }
    bool IsRunning() const;
    bool Wait(DWORD timeout_ms) const;
    bool Terminate(UINT exit_code);
    void Close();

private:
    friend class ProcessRunner;

    HANDLE process_ = nullptr;
    HANDLE thread_ = nullptr;
    DWORD process_id_ = 0;
};

class ProcessRunner {
public:
    static bool Start(const ProcessStartOptions& options,
                      ProcessHandle* handle,
                      std::wstring* error_message);
};

}  // namespace mytools
