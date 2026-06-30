#include "app/tray_host.h"

#include <shellapi.h>

#include "resources/resource.h"

namespace mytools {

TrayHost::~TrayHost() {
    Remove();
}

bool TrayHost::Add(HWND window, HINSTANCE instance) {
    if (added_) {
        return true;
    }

    data_ = {};
    data_.cbSize = sizeof(data_);
    data_.hWnd = window;
    data_.uID = 1;
    data_.uCallbackMessage = WM_MYTOOLS_TRAY;
    data_.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    data_.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
    if (data_.hIcon == nullptr) {
        data_.hIcon = LoadIconW(nullptr, IDI_APPLICATION);
    }
    wcscpy_s(data_.szTip, L"AJun Tools Native");

    added_ = Shell_NotifyIconW(NIM_ADD, &data_) == TRUE;
    return added_;
}

void TrayHost::Remove() {
    if (!added_) {
        return;
    }
    Shell_NotifyIconW(NIM_DELETE, &data_);
    added_ = false;
}

void TrayHost::ShowMenu(HWND owner) {
    HMENU menu = CreatePopupMenu();
    if (menu == nullptr) {
        return;
    }

    AppendMenuW(menu, MF_STRING, ID_TRAY_SHOW, L"Show window");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT, L"Exit");

    POINT cursor{};
    GetCursorPos(&cursor);
    SetForegroundWindow(owner);
    TrackPopupMenu(menu, TPM_RIGHTBUTTON, cursor.x, cursor.y, 0, owner, nullptr);
    DestroyMenu(menu);
}

}  // namespace mytools
