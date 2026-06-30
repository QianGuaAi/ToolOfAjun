#include "services/process_runner.h"

#include "services/file_system.h"

#include <utility>

namespace mytools {
namespace {

std::wstring QuoteForCommandLine(const std::wstring& value) {
    std::wstring quoted = L"\"";
    for (wchar_t ch : value) {
        if (ch == L'"') {
            quoted += L'\\';
        }
        quoted += ch;
    }
    quoted += L"\"";
    return quoted;
}

}  // namespace

ProcessHandle::~ProcessHandle() {
    Close();
}

ProcessHandle::ProcessHandle(ProcessHandle&& other) noexcept {
    *this = std::move(other);
}

ProcessHandle& ProcessHandle::operator=(ProcessHandle&& other) noexcept {
    if (this != &other) {
        Close();
        process_ = other.process_;
        thread_ = other.thread_;
        process_id_ = other.process_id_;
        other.process_ = nullptr;
        other.thread_ = nullptr;
        other.process_id_ = 0;
    }
    return *this;
}

bool ProcessHandle::IsRunning() const {
    if (process_ == nullptr) {
        return false;
    }

    DWORD exit_code = 0;
    if (!GetExitCodeProcess(process_, &exit_code)) {
        return false;
    }
    return exit_code == STILL_ACTIVE;
}

bool ProcessHandle::Wait(DWORD timeout_ms) const {
    return process_ != nullptr && WaitForSingleObject(process_, timeout_ms) == WAIT_OBJECT_0;
}

bool ProcessHandle::Terminate(UINT exit_code) {
    return process_ != nullptr && TerminateProcess(process_, exit_code) != FALSE;
}

void ProcessHandle::Close() {
    if (thread_ != nullptr) {
        CloseHandle(thread_);
        thread_ = nullptr;
    }
    if (process_ != nullptr) {
        CloseHandle(process_);
        process_ = nullptr;
    }
    process_id_ = 0;
}

bool ProcessRunner::Start(const ProcessStartOptions& options,
                          ProcessHandle* handle,
                          std::wstring* error_message) {
    if (handle == nullptr || options.file_path.empty()) {
        if (error_message != nullptr) {
            *error_message = L"ProcessRunner::Start requires a file path and output handle.";
        }
        return false;
    }

    std::wstring command_line = QuoteForCommandLine(options.file_path);
    if (!options.arguments.empty()) {
        command_line += L" ";
        command_line += options.arguments;
    }

    STARTUPINFOW startup_info{};
    startup_info.cb = sizeof(startup_info);
    PROCESS_INFORMATION process_info{};
    DWORD creation_flags = 0;
    if (options.create_no_window) {
        creation_flags |= CREATE_NO_WINDOW;
    }

    std::wstring mutable_command_line = command_line;
    const BOOL ok = CreateProcessW(options.file_path.c_str(),
                                   mutable_command_line.data(),
                                   nullptr,
                                   nullptr,
                                   FALSE,
                                   creation_flags,
                                   nullptr,
                                   options.working_directory.empty() ? nullptr
                                                                     : options.working_directory.c_str(),
                                   &startup_info,
                                   &process_info);
    if (!ok) {
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"CreateProcessW");
        }
        return false;
    }

    handle->Close();
    handle->process_ = process_info.hProcess;
    handle->thread_ = process_info.hThread;
    handle->process_id_ = process_info.dwProcessId;
    return true;
}

}  // namespace mytools
