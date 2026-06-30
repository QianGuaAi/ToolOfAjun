#pragma once

#include <chrono>
#include <string>

#include <windows.h>

#include "app/app_context.h"
#include "app/tray_host.h"
#include "modules/codex_profiles/codex_profile_module.h"
#include "modules/module_registry.h"
#include "services/logger.h"
#include "ui/renderer_d2d.h"

namespace mytools {

class MainWindow {
public:
    MainWindow(const AppContext& context, Logger* logger);

    bool Create();
    void Show(int command_show);
    int RunMessageLoop();

private:
    static LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam);
    LRESULT HandleMessage(UINT message, WPARAM wparam, LPARAM lparam);

    void CreateMenuBar();
    void SwitchModule(ModuleId module_id);
    void HandleCodexProfileAction(CodexProfileUiAction action);
    bool PrepareCodexProfileActionOptions(CodexProfileUiAction action,
                                          CodexProfileActionOptions* options) const;
    bool PromptCodexProfileTarget(CodexProfileUiAction action,
                                  CodexProfileActionOptions* options) const;
    bool PromptCodexBoxPassword(CodexProfileUiAction action, std::wstring* password) const;
    bool PromptCodexBoxConflictPolicy(CodexProfileBoxConflictPolicy* policy) const;
    bool PromptCodexProfileDisplayName(std::wstring* display_name) const;
    bool PromptCodexProfileMetadataText(CodexProfileUiAction action, std::wstring* value) const;
    bool ConfirmCodexProfileWrite(CodexProfileUiAction action) const;
    void RestoreFromTray();
    void HideToTray();
    void ExitApplication();
    float CurrentDpiScale() const;

    const AppContext& context_;
    Logger* logger_ = nullptr;
    CodexProfileModule codex_profiles_;
    RendererD2D renderer_;
    TrayHost tray_;
    HWND window_ = nullptr;
    bool exiting_ = false;
    UINT dpi_ = 96;
    ModuleInfo current_module_;
};

}  // namespace mytools
