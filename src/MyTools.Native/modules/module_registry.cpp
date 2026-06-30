#include "modules/module_registry.h"

namespace mytools {

ModuleInfo BuildHomeModuleInfo() {
    ModuleInfo info;
    info.id = ModuleId::Home;
    info.title = L"AJun Tools Native";
    info.subtitle =
        L"Native rewrite shell: Win32 + Direct2D + DirectWrite, ready for phased module migration.";
    info.status = L"Current location: Home";
    info.bullets = {
        L"Phase 1: native shell, tray, DPI, logs, DPAPI smoke test.",
        L"Phase 2: config, file system, scan, task, process, network, hotkey services.",
        L"Phase 3 begins with Codex profile module scaffolding.",
        L"Removed features stay removed: SQL export, rotation pool, Ollama import, sensors, bundled FFmpeg."};
    return info;
}

}  // namespace mytools
