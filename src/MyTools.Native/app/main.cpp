#include <chrono>
#include <sstream>

#include <windows.h>

#include "app/app_context.h"
#include "app/crash_handler.h"
#include "app/single_instance.h"
#include "services/logger.h"
#include "services/secret_store_dpapi.h"
#include "ui/main_window.h"

namespace {

void EnableDpiAwareness() {
    HMODULE user32 = LoadLibraryW(L"user32.dll");
    if (user32 == nullptr) {
        return;
    }

    using SetDpiAwarenessContextProc = BOOL(WINAPI*)(DPI_AWARENESS_CONTEXT);
    auto set_process_dpi_awareness_context =
        reinterpret_cast<SetDpiAwarenessContextProc>(GetProcAddress(user32, "SetProcessDpiAwarenessContext"));
    if (set_process_dpi_awareness_context != nullptr) {
        set_process_dpi_awareness_context(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    }

    FreeLibrary(user32);
}

std::wstring ElapsedText(std::chrono::steady_clock::time_point start) {
    const auto elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(
        std::chrono::steady_clock::now() - start);
    std::wstringstream text;
    text << elapsed.count() << L" ms";
    return text.str();
}

}  // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int command_show) {
    const auto startup_start = std::chrono::steady_clock::now();
    EnableDpiAwareness();

    mytools::SingleInstance single_instance;
    if (!single_instance.IsPrimary()) {
        MessageBoxW(nullptr,
                    L"AJun Tools Native is already running.",
                    L"AJun Tools Native",
                    MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    mytools::AppContext context = mytools::AppContext::Create(instance);
    mytools::Logger logger(context.exe_dir());
    mytools::InstallCrashHandler(&logger);
    logger.Startup(L"Native shell startup begin.");

    mytools::SecretStoreDpapi secret_store;
    std::wstring dpapi_error;
    if (secret_store.SmokeTest(&dpapi_error)) {
        logger.Startup(L"DPAPI smoke test passed.");
    } else {
        logger.Error(L"DPAPI smoke test failed: " + dpapi_error);
        MessageBoxW(nullptr, dpapi_error.c_str(), L"DPAPI smoke test failed", MB_OK | MB_ICONERROR);
        return 2;
    }

    mytools::MainWindow window(context, &logger);
    if (!window.Create()) {
        MessageBoxW(nullptr,
                    L"Failed to create the native main window. See MyToolsNative.log.",
                    L"AJun Tools Native",
                    MB_OK | MB_ICONERROR);
        return 3;
    }

    logger.Startup(L"Native shell first window created in " + ElapsedText(startup_start) + L".");
    window.Show(command_show);
    return window.RunMessageLoop();
}
