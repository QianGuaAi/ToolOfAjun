#include "services/hotkey_service.h"

#include "services/file_system.h"

namespace mytools {

HotkeyService::HotkeyService(HWND hwnd) : hwnd_(hwnd) {}

HotkeyService::~HotkeyService() {
    UnregisterAll();
}

bool HotkeyService::Register(int id, UINT modifiers, UINT virtual_key, std::wstring* error_message) {
    if (hwnd_ == nullptr) {
        if (error_message != nullptr) {
            *error_message = L"HotkeyService requires a valid window handle.";
        }
        return false;
    }

    Unregister(id);
    if (!RegisterHotKey(hwnd_, id, modifiers, virtual_key)) {
        if (error_message != nullptr) {
            *error_message = FormatLastErrorMessage(L"RegisterHotKey");
        }
        return false;
    }

    registered_ids_.insert(id);
    return true;
}

void HotkeyService::Unregister(int id) {
    const auto found = registered_ids_.find(id);
    if (found == registered_ids_.end()) {
        return;
    }

    UnregisterHotKey(hwnd_, id);
    registered_ids_.erase(found);
}

void HotkeyService::UnregisterAll() {
    for (int id : registered_ids_) {
        UnregisterHotKey(hwnd_, id);
    }
    registered_ids_.clear();
}

}  // namespace mytools
