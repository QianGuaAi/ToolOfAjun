#pragma once

#include <set>
#include <string>

#include <windows.h>

namespace mytools {

class HotkeyService {
public:
    explicit HotkeyService(HWND hwnd);
    ~HotkeyService();

    HotkeyService(const HotkeyService&) = delete;
    HotkeyService& operator=(const HotkeyService&) = delete;

    bool Register(int id, UINT modifiers, UINT virtual_key, std::wstring* error_message);
    void Unregister(int id);
    void UnregisterAll();

private:
    HWND hwnd_ = nullptr;
    std::set<int> registered_ids_;
};

}  // namespace mytools
