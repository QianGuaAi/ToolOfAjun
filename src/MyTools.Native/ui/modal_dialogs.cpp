#include "ui/modal_dialogs.h"

#include <vector>

#include <commdlg.h>
#include <objbase.h>
#include <shlobj.h>

namespace mytools {
namespace {

constexpr wchar_t kPasswordDialogClass[] = L"AJunToolsNativePasswordDialog";
constexpr wchar_t kConflictDialogClass[] = L"AJunToolsNativeConflictDialog";
constexpr int kEditId = 1001;
constexpr int kButtonOkId = IDOK;
constexpr int kButtonCancelId = IDCANCEL;
constexpr int kButtonRenameId = 2001;
constexpr int kButtonSkipId = 2002;
constexpr int kButtonReplaceId = 2003;

struct PasswordDialogState {
    HWND owner = nullptr;
    HWND window = nullptr;
    HWND edit = nullptr;
    std::wstring title;
    std::wstring prompt;
    std::wstring value;
    bool password_mode = true;
    bool accepted = false;
};

struct ConflictDialogState {
    HWND owner = nullptr;
    HWND window = nullptr;
    CodexProfileBoxConflictPolicy selected = CodexProfileBoxConflictPolicy::Rename;
    bool accepted = false;
};

void ApplyDefaultFont(HWND control) {
    HFONT font = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    if (font != nullptr && control != nullptr) {
        SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
}

void CenterWindow(HWND window, HWND owner) {
    RECT dialog_rect{};
    RECT owner_rect{};
    GetWindowRect(window, &dialog_rect);
    if (owner != nullptr && IsWindow(owner)) {
        GetWindowRect(owner, &owner_rect);
    } else {
        owner_rect.left = 0;
        owner_rect.top = 0;
        owner_rect.right = GetSystemMetrics(SM_CXSCREEN);
        owner_rect.bottom = GetSystemMetrics(SM_CYSCREEN);
    }

    const int width = dialog_rect.right - dialog_rect.left;
    const int height = dialog_rect.bottom - dialog_rect.top;
    const int owner_width = owner_rect.right - owner_rect.left;
    const int owner_height = owner_rect.bottom - owner_rect.top;
    const int x = owner_rect.left + (owner_width - width) / 2;
    const int y = owner_rect.top + (owner_height - height) / 2;
    SetWindowPos(window, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
}

void FinishDialog(HWND window, bool accepted) {
    auto* state = reinterpret_cast<PasswordDialogState*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    if (state != nullptr) {
        state->accepted = accepted;
        if (accepted && state->edit != nullptr) {
            const int length = GetWindowTextLengthW(state->edit);
            if (length > 0) {
                std::wstring value(static_cast<size_t>(length) + 1, L'\0');
                GetWindowTextW(state->edit, value.data(), length + 1);
                value.resize(static_cast<size_t>(length));
                state->value = std::move(value);
            } else {
                state->value.clear();
            }
            SetWindowTextW(state->edit, L"");
        }
    }
    DestroyWindow(window);
}

LRESULT CALLBACK PasswordDialogProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    auto* state = reinterpret_cast<PasswordDialogState*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    switch (message) {
        case WM_CREATE: {
            auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
            state = static_cast<PasswordDialogState*>(create->lpCreateParams);
            SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
            state->window = window;

            HWND prompt = CreateWindowExW(0,
                                          L"STATIC",
                                          state->prompt.c_str(),
                                          WS_CHILD | WS_VISIBLE,
                                          16,
                                          16,
                                          420,
                                          36,
                                          window,
                                          nullptr,
                                          nullptr,
                                          nullptr);
            state->edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                          L"EDIT",
                                          L"",
                                          WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL |
                                              (state->password_mode ? ES_PASSWORD : 0),
                                          16,
                                          58,
                                          420,
                                          24,
                                          window,
                                          reinterpret_cast<HMENU>(static_cast<INT_PTR>(kEditId)),
                                          nullptr,
                                          nullptr);
            if (state->password_mode) {
                SendMessageW(state->edit, EM_SETPASSWORDCHAR, static_cast<WPARAM>(L'*'), 0);
            }

            HWND ok = CreateWindowExW(0,
                                      L"BUTTON",
                                      L"OK",
                                      WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                                      250,
                                      102,
                                      86,
                                      28,
                                      window,
                                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(kButtonOkId)),
                                      nullptr,
                                      nullptr);
            HWND cancel = CreateWindowExW(0,
                                          L"BUTTON",
                                          L"Cancel",
                                          WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                          350,
                                          102,
                                          86,
                                          28,
                                          window,
                                          reinterpret_cast<HMENU>(static_cast<INT_PTR>(kButtonCancelId)),
                                          nullptr,
                                          nullptr);
            ApplyDefaultFont(prompt);
            ApplyDefaultFont(state->edit);
            ApplyDefaultFont(ok);
            ApplyDefaultFont(cancel);
            SetFocus(state->edit);
            return 0;
        }

        case WM_COMMAND:
            if (LOWORD(wparam) == kButtonOkId) {
                FinishDialog(window, true);
                return 0;
            }
            if (LOWORD(wparam) == kButtonCancelId) {
                FinishDialog(window, false);
                return 0;
            }
            break;

        case WM_CLOSE:
            FinishDialog(window, false);
            return 0;

        case WM_DESTROY:
            SetWindowLongPtrW(window, GWLP_USERDATA, 0);
            return 0;
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

void FinishConflictDialog(HWND window,
                          bool accepted,
                          CodexProfileBoxConflictPolicy selected) {
    auto* state = reinterpret_cast<ConflictDialogState*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    if (state != nullptr) {
        state->accepted = accepted;
        state->selected = selected;
    }
    DestroyWindow(window);
}

LRESULT CALLBACK ConflictDialogProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    auto* state = reinterpret_cast<ConflictDialogState*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    switch (message) {
        case WM_CREATE: {
            auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
            state = static_cast<ConflictDialogState*>(create->lpCreateParams);
            SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
            state->window = window;

            HWND prompt = CreateWindowExW(
                0,
                L"STATIC",
                L"Choose how to handle .codexbox profiles whose names already exist locally:",
                WS_CHILD | WS_VISIBLE,
                16,
                16,
                468,
                38,
                window,
                nullptr,
                nullptr,
                nullptr);
            HWND rename = CreateWindowExW(0,
                                          L"BUTTON",
                                          L"Rename",
                                          WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                                          16,
                                          70,
                                          102,
                                          28,
                                          window,
                                          reinterpret_cast<HMENU>(
                                              static_cast<INT_PTR>(kButtonRenameId)),
                                          nullptr,
                                          nullptr);
            HWND skip = CreateWindowExW(0,
                                        L"BUTTON",
                                        L"Skip",
                                        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                        130,
                                        70,
                                        102,
                                        28,
                                        window,
                                        reinterpret_cast<HMENU>(
                                            static_cast<INT_PTR>(kButtonSkipId)),
                                        nullptr,
                                        nullptr);
            HWND replace = CreateWindowExW(0,
                                           L"BUTTON",
                                           L"Replace",
                                           WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                           244,
                                           70,
                                           102,
                                           28,
                                           window,
                                           reinterpret_cast<HMENU>(
                                               static_cast<INT_PTR>(kButtonReplaceId)),
                                           nullptr,
                                           nullptr);
            HWND cancel = CreateWindowExW(0,
                                          L"BUTTON",
                                          L"Cancel",
                                          WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                          382,
                                          70,
                                          102,
                                          28,
                                          window,
                                          reinterpret_cast<HMENU>(
                                              static_cast<INT_PTR>(kButtonCancelId)),
                                          nullptr,
                                          nullptr);
            ApplyDefaultFont(prompt);
            ApplyDefaultFont(rename);
            ApplyDefaultFont(skip);
            ApplyDefaultFont(replace);
            ApplyDefaultFont(cancel);
            SetFocus(rename);
            return 0;
        }

        case WM_COMMAND:
            switch (LOWORD(wparam)) {
                case kButtonRenameId:
                    FinishConflictDialog(window, true, CodexProfileBoxConflictPolicy::Rename);
                    return 0;
                case kButtonSkipId:
                    FinishConflictDialog(window, true, CodexProfileBoxConflictPolicy::Skip);
                    return 0;
                case kButtonReplaceId:
                    FinishConflictDialog(window, true, CodexProfileBoxConflictPolicy::Replace);
                    return 0;
                case kButtonCancelId:
                    FinishConflictDialog(window, false, CodexProfileBoxConflictPolicy::Rename);
                    return 0;
                default:
                    break;
            }
            break;

        case WM_CLOSE:
            FinishConflictDialog(window, false, CodexProfileBoxConflictPolicy::Rename);
            return 0;

        case WM_DESTROY:
            SetWindowLongPtrW(window, GWLP_USERDATA, 0);
            return 0;
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

bool EnsurePasswordDialogClass(HINSTANCE instance) {
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = PasswordDialogProc;
    wc.hInstance = instance;
    wc.lpszClassName = kPasswordDialogClass;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    if (RegisterClassExW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS) {
        return true;
    }
    return false;
}

bool EnsureConflictDialogClass(HINSTANCE instance) {
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = ConflictDialogProc;
    wc.hInstance = instance;
    wc.lpszClassName = kConflictDialogClass;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    if (RegisterClassExW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS) {
        return true;
    }
    return false;
}

bool RunTextDialog(HWND owner,
                   const std::wstring& title,
                   const std::wstring& prompt,
                   bool password_mode,
                   std::wstring* value) {
    if (value == nullptr) {
        return false;
    }
    value->clear();

    HINSTANCE instance = GetModuleHandleW(nullptr);
    if (!EnsurePasswordDialogClass(instance)) {
        return false;
    }

    PasswordDialogState state;
    state.owner = owner;
    state.title = title;
    state.prompt = prompt;
    state.password_mode = password_mode;

    RECT rect{0, 0, 468, 168};
    AdjustWindowRectEx(&rect, WS_POPUP | WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
    HWND dialog = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
                                  kPasswordDialogClass,
                                  title.c_str(),
                                  WS_POPUP | WS_CAPTION | WS_SYSMENU,
                                  CW_USEDEFAULT,
                                  CW_USEDEFAULT,
                                  rect.right - rect.left,
                                  rect.bottom - rect.top,
                                  owner,
                                  nullptr,
                                  instance,
                                  &state);
    if (dialog == nullptr) {
        return false;
    }

    CenterWindow(dialog, owner);
    if (owner != nullptr && IsWindow(owner)) {
        EnableWindow(owner, FALSE);
    }
    ShowWindow(dialog, SW_SHOW);
    UpdateWindow(dialog);

    MSG message{};
    while (IsWindow(dialog)) {
        const BOOL get_result = GetMessageW(&message, nullptr, 0, 0);
        if (get_result == -1) {
            break;
        }
        if (get_result == 0) {
            PostQuitMessage(static_cast<int>(message.wParam));
            break;
        }
        if (!IsDialogMessageW(dialog, &message)) {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    if (owner != nullptr && IsWindow(owner)) {
        EnableWindow(owner, TRUE);
        SetForegroundWindow(owner);
    }

    if (!state.accepted) {
        SecureClearWideString(&state.value);
        return false;
    }

    *value = state.value;
    SecureClearWideString(&state.value);
    return true;
}

bool RunConflictDialog(HWND owner, CodexProfileBoxConflictPolicy* policy) {
    if (policy == nullptr) {
        return false;
    }
    *policy = CodexProfileBoxConflictPolicy::Rename;

    HINSTANCE instance = GetModuleHandleW(nullptr);
    if (!EnsureConflictDialogClass(instance)) {
        return false;
    }

    ConflictDialogState state;
    state.owner = owner;

    RECT rect{0, 0, 516, 138};
    AdjustWindowRectEx(&rect, WS_POPUP | WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
    HWND dialog = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
                                  kConflictDialogClass,
                                  L".codexbox import conflicts",
                                  WS_POPUP | WS_CAPTION | WS_SYSMENU,
                                  CW_USEDEFAULT,
                                  CW_USEDEFAULT,
                                  rect.right - rect.left,
                                  rect.bottom - rect.top,
                                  owner,
                                  nullptr,
                                  instance,
                                  &state);
    if (dialog == nullptr) {
        return false;
    }

    CenterWindow(dialog, owner);
    if (owner != nullptr && IsWindow(owner)) {
        EnableWindow(owner, FALSE);
    }
    ShowWindow(dialog, SW_SHOW);
    UpdateWindow(dialog);

    MSG message{};
    while (IsWindow(dialog)) {
        const BOOL get_result = GetMessageW(&message, nullptr, 0, 0);
        if (get_result == -1) {
            break;
        }
        if (get_result == 0) {
            PostQuitMessage(static_cast<int>(message.wParam));
            break;
        }
        if (!IsDialogMessageW(dialog, &message)) {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    if (owner != nullptr && IsWindow(owner)) {
        EnableWindow(owner, TRUE);
        SetForegroundWindow(owner);
    }

    if (!state.accepted) {
        return false;
    }
    *policy = state.selected;
    return true;
}

bool PickCodexBoxPath(HWND owner, bool save, std::wstring* path) {
    if (path == nullptr) {
        return false;
    }
    path->clear();

    std::vector<wchar_t> buffer(32768, L'\0');
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = L"Codex profile package (*.codexbox)\0*.codexbox\0All files (*.*)\0*.*\0";
    ofn.lpstrFile = buffer.data();
    ofn.nMaxFile = static_cast<DWORD>(buffer.size());
    ofn.lpstrDefExt = L"codexbox";
    ofn.Flags = OFN_EXPLORER | OFN_NOCHANGEDIR | OFN_PATHMUSTEXIST;
    if (save) {
        ofn.Flags |= OFN_OVERWRITEPROMPT;
        if (!GetSaveFileNameW(&ofn)) {
            return false;
        }
    } else {
        ofn.Flags |= OFN_FILEMUSTEXIST;
        if (!GetOpenFileNameW(&ofn)) {
            return false;
        }
    }

    *path = buffer.data();
    return !path->empty();
}

bool PickFolder(HWND owner, const std::wstring& title, std::wstring* path) {
    if (path == nullptr) {
        return false;
    }
    path->clear();

    const HRESULT com_result = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool com_initialized = SUCCEEDED(com_result);

    BROWSEINFOW browse{};
    browse.hwndOwner = owner;
    browse.lpszTitle = title.empty() ? L"Select folder" : title.c_str();
    browse.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_USENEWUI;

    PIDLIST_ABSOLUTE item = SHBrowseForFolderW(&browse);
    bool ok = false;
    if (item != nullptr) {
        std::vector<wchar_t> buffer(MAX_PATH, L'\0');
        if (SHGetPathFromIDListW(item, buffer.data()) != FALSE) {
            *path = buffer.data();
            ok = !path->empty();
        }
        CoTaskMemFree(item);
    }

    if (com_initialized) {
        CoUninitialize();
    }
    return ok;
}

}  // namespace

void SecureClearWideString(std::wstring* value) {
    if (value != nullptr && !value->empty()) {
        SecureZeroMemory(value->data(), value->size() * sizeof(wchar_t));
        value->clear();
    }
}

bool PromptPassword(HWND owner,
                    const std::wstring& title,
                    const std::wstring& prompt,
                    std::wstring* password) {
    return RunTextDialog(owner, title, prompt, true, password);
}

bool PromptText(HWND owner,
                const std::wstring& title,
                const std::wstring& prompt,
                std::wstring* value) {
    return RunTextDialog(owner, title, prompt, false, value);
}

bool PickCodexBoxSavePath(HWND owner, std::wstring* path) {
    return PickCodexBoxPath(owner, true, path);
}

bool PickCodexBoxOpenPath(HWND owner, std::wstring* path) {
    return PickCodexBoxPath(owner, false, path);
}

bool PickFolderPath(HWND owner, const std::wstring& title, std::wstring* path) {
    return PickFolder(owner, title, path);
}

bool ChooseCodexBoxConflictPolicy(HWND owner, CodexProfileBoxConflictPolicy* policy) {
    return RunConflictDialog(owner, policy);
}

}  // namespace mytools
