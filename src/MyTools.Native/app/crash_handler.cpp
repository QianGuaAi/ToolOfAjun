#include "app/crash_handler.h"

#include <sstream>

#include <windows.h>

#include "services/logger.h"

namespace mytools {
namespace {

Logger* g_logger = nullptr;

LONG WINAPI HandleUnhandledException(EXCEPTION_POINTERS* exception_info) {
    if (g_logger != nullptr && exception_info != nullptr && exception_info->ExceptionRecord != nullptr) {
        std::wstringstream message;
        message << L"Unhandled exception code=0x" << std::hex
                << exception_info->ExceptionRecord->ExceptionCode;
        g_logger->Crash(message.str());
    }
    return EXCEPTION_EXECUTE_HANDLER;
}

}  // namespace

void InstallCrashHandler(Logger* logger) {
    g_logger = logger;
    SetUnhandledExceptionFilter(HandleUnhandledException);
}

}  // namespace mytools
