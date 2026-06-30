#pragma once

#include <windows.h>

namespace mytools {

constexpr UINT WM_MYTOOLS_TRAY = WM_APP + 0x42;
constexpr UINT ID_TRAY_SHOW = 40101;
constexpr UINT ID_TRAY_EXIT = 40102;

class TrayHost {
public:
    TrayHost() = default;
    ~TrayHost();

    TrayHost(const TrayHost&) = delete;
    TrayHost& operator=(const TrayHost&) = delete;

    bool Add(HWND window, HINSTANCE instance);
    void Remove();
    void ShowMenu(HWND owner);

private:
    NOTIFYICONDATAW data_{};
    bool added_ = false;
};

}  // namespace mytools
