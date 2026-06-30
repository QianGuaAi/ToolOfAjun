#include "ui/main_window.h"

#include <cwctype>
#include <sstream>

#include <windowsx.h>

#include "resources/resource.h"
#include "ui/modal_dialogs.h"

namespace mytools {
namespace {

constexpr wchar_t kWindowClassName[] = L"AJunToolsNativeMainWindow";
constexpr UINT ID_FILE_EXIT = 1001;
constexpr UINT ID_HELP_ABOUT = 1002;
constexpr UINT ID_NAV_HOME = 1101;
constexpr UINT ID_NAV_CODEX_PROFILES = 1102;
constexpr UINT ID_CODEX_REFRESH = 1201;
constexpr UINT ID_CODEX_DIFF_FIRST = 1202;
constexpr UINT ID_CODEX_BACKUP_CURRENT = 1203;
constexpr UINT ID_CODEX_APPLY_FIRST = 1204;
constexpr UINT ID_CODEX_IMPORT_CURRENT = 1205;
constexpr UINT ID_CODEX_RESTORE_BACKUP = 1206;
constexpr UINT ID_CODEX_EXPORT_BOX = 1207;
constexpr UINT ID_CODEX_IMPORT_BOX = 1208;
constexpr UINT ID_CODEX_EXPORT_FIRST_FILES = 1209;
constexpr UINT ID_CODEX_RENAME_FIRST = 1210;
constexpr UINT ID_CODEX_EDIT_FIRST_NOTE = 1211;
constexpr UINT ID_CODEX_EDIT_FIRST_REMARK = 1212;
constexpr UINT ID_CODEX_EDIT_FIRST_TAGS = 1213;

std::wstring TrimWide(const std::wstring& value) {
    size_t first = 0;
    while (first < value.size() && std::iswspace(value[first]) != 0) {
        ++first;
    }

    size_t last = value.size();
    while (last > first && std::iswspace(value[last - 1]) != 0) {
        --last;
    }

    return value.substr(first, last - first);
}

bool ContainsControlCharacter(const std::wstring& value) {
    for (const wchar_t ch : value) {
        const unsigned int code_point = static_cast<unsigned int>(ch);
        if (code_point < 0x20 || (code_point >= 0x7F && code_point <= 0x9F)) {
            return true;
        }
    }
    return false;
}

}  // namespace

MainWindow::MainWindow(const AppContext& context, Logger* logger)
    : context_(context),
      logger_(logger),
      codex_profiles_(logger),
      current_module_(BuildHomeModuleInfo()) {}

bool MainWindow::Create() {
    if (!renderer_.Initialize()) {
        if (logger_ != nullptr) {
            logger_->Error(L"Direct2D/DirectWrite initialization failed.");
        }
        return false;
    }

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = MainWindow::WindowProc;
    wc.hInstance = context_.instance();
    wc.lpszClassName = kWindowClassName;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = LoadIconW(context_.instance(), MAKEINTRESOURCEW(IDI_APPICON));
    wc.hIconSm = wc.hIcon;
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);

    if (RegisterClassExW(&wc) == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        if (logger_ != nullptr) {
            logger_->Error(L"RegisterClassExW failed.");
        }
        return false;
    }

    window_ = CreateWindowExW(0,
                              kWindowClassName,
                              L"AJun Tools Native",
                              WS_OVERLAPPEDWINDOW,
                              CW_USEDEFAULT,
                              CW_USEDEFAULT,
                              1280,
                              800,
                              nullptr,
                              nullptr,
                              context_.instance(),
                              this);
    if (window_ == nullptr) {
        if (logger_ != nullptr) {
            logger_->Error(L"CreateWindowExW failed.");
        }
        return false;
    }

    CreateMenuBar();
    tray_.Add(window_, context_.instance());
    return true;
}

void MainWindow::Show(int command_show) {
    ShowWindow(window_, command_show);
    UpdateWindow(window_);
}

int MainWindow::RunMessageLoop() {
    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    return static_cast<int>(message.wParam);
}

LRESULT CALLBACK MainWindow::WindowProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    auto* self = reinterpret_cast<MainWindow*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    if (message == WM_NCCREATE) {
        auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
        self = static_cast<MainWindow*>(create->lpCreateParams);
        SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }

    if (self != nullptr) {
        return self->HandleMessage(message, wparam, lparam);
    }
    return DefWindowProcW(window, message, wparam, lparam);
}

LRESULT MainWindow::HandleMessage(UINT message, WPARAM wparam, LPARAM lparam) {
    switch (message) {
        case WM_CREATE:
            dpi_ = GetDpiForWindow(window_);
            return 0;

        case WM_COMMAND:
            switch (LOWORD(wparam)) {
                case ID_FILE_EXIT:
                case ID_TRAY_EXIT:
                    ExitApplication();
                    return 0;
                case ID_TRAY_SHOW:
                    RestoreFromTray();
                    return 0;
                case ID_NAV_HOME:
                    SwitchModule(ModuleId::Home);
                    return 0;
                case ID_NAV_CODEX_PROFILES:
                    SwitchModule(ModuleId::CodexProfiles);
                    return 0;
                case ID_CODEX_REFRESH:
                    HandleCodexProfileAction(CodexProfileUiAction::Refresh);
                    return 0;
                case ID_CODEX_DIFF_FIRST:
                    HandleCodexProfileAction(CodexProfileUiAction::DiffFirstProfile);
                    return 0;
                case ID_CODEX_BACKUP_CURRENT:
                    HandleCodexProfileAction(CodexProfileUiAction::BackupCurrentFolder);
                    return 0;
                case ID_CODEX_APPLY_FIRST:
                    HandleCodexProfileAction(CodexProfileUiAction::ApplyFirstProfile);
                    return 0;
                case ID_CODEX_IMPORT_CURRENT:
                    HandleCodexProfileAction(CodexProfileUiAction::ImportCurrentFolder);
                    return 0;
                case ID_CODEX_RESTORE_BACKUP:
                    HandleCodexProfileAction(CodexProfileUiAction::RestoreLatestBackup);
                    return 0;
                case ID_CODEX_EXPORT_FIRST_FILES:
                    HandleCodexProfileAction(CodexProfileUiAction::ExportFirstProfileFiles);
                    return 0;
                case ID_CODEX_RENAME_FIRST:
                    HandleCodexProfileAction(CodexProfileUiAction::RenameFirstProfile);
                    return 0;
                case ID_CODEX_EDIT_FIRST_NOTE:
                    HandleCodexProfileAction(CodexProfileUiAction::EditFirstProfileNote);
                    return 0;
                case ID_CODEX_EDIT_FIRST_REMARK:
                    HandleCodexProfileAction(CodexProfileUiAction::EditFirstProfileRemark);
                    return 0;
                case ID_CODEX_EDIT_FIRST_TAGS:
                    HandleCodexProfileAction(CodexProfileUiAction::EditFirstProfileTags);
                    return 0;
                case ID_CODEX_EXPORT_BOX:
                    HandleCodexProfileAction(CodexProfileUiAction::ExportBox);
                    return 0;
                case ID_CODEX_IMPORT_BOX:
                    HandleCodexProfileAction(CodexProfileUiAction::ImportBox);
                    return 0;
                case ID_HELP_ABOUT:
                    MessageBoxW(window_,
                                L"AJun Tools Native\nPhase 1/2 shell and foundation services.\nStage 3 Codex Profiles explicit menu actions are being migrated.",
                                L"About",
                                MB_OK | MB_ICONINFORMATION);
                    return 0;
                default:
                    return DefWindowProcW(window_, message, wparam, lparam);
            }

        case WM_SIZE:
            renderer_.Resize(window_);
            return 0;

        case WM_DPICHANGED: {
            dpi_ = HIWORD(wparam);
            const RECT* suggested = reinterpret_cast<RECT*>(lparam);
            SetWindowPos(window_,
                         nullptr,
                         suggested->left,
                         suggested->top,
                         suggested->right - suggested->left,
                         suggested->bottom - suggested->top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            InvalidateRect(window_, nullptr, FALSE);
            return 0;
        }

        case WM_PAINT: {
            PAINTSTRUCT paint{};
            BeginPaint(window_, &paint);
            renderer_.Render(window_, CurrentDpiScale(), current_module_);
            EndPaint(window_, &paint);
            return 0;
        }

        case WM_ERASEBKGND:
            return 1;

        case WM_CLOSE:
            if (exiting_) {
                DestroyWindow(window_);
            } else {
                HideToTray();
            }
            return 0;

        case WM_MYTOOLS_TRAY:
            if (LOWORD(lparam) == WM_LBUTTONDBLCLK) {
                RestoreFromTray();
            } else if (LOWORD(lparam) == WM_RBUTTONUP) {
                tray_.ShowMenu(window_);
            }
            return 0;

        case WM_DESTROY:
            tray_.Remove();
            PostQuitMessage(0);
            return 0;

        default:
            return DefWindowProcW(window_, message, wparam, lparam);
    }
}

void MainWindow::CreateMenuBar() {
    HMENU menu_bar = CreateMenu();
    HMENU file_menu = CreatePopupMenu();
    HMENU system_menu = CreatePopupMenu();
    HMENU tools_menu = CreatePopupMenu();
    HMENU codex_menu = CreatePopupMenu();
    HMENU help_menu = CreatePopupMenu();

    AppendMenuW(file_menu, MF_STRING, ID_FILE_EXIT, L"Exit");
    AppendMenuW(tools_menu, MF_STRING, ID_NAV_HOME, L"Home");
    AppendMenuW(codex_menu, MF_STRING, ID_NAV_CODEX_PROFILES, L"Open overview");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_REFRESH, L"Refresh summaries");
    AppendMenuW(codex_menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_DIFF_FIRST, L"Diff first profile");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_BACKUP_CURRENT, L"Backup current .codex folder");
    AppendMenuW(codex_menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_APPLY_FIRST, L"Apply first profile...");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_IMPORT_CURRENT, L"Import current folder...");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_RESTORE_BACKUP, L"Restore latest backup...");
    AppendMenuW(codex_menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_EXPORT_FIRST_FILES, L"Export first profile files...");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_RENAME_FIRST, L"Rename first profile...");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_EDIT_FIRST_NOTE, L"Edit first profile note...");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_EDIT_FIRST_REMARK, L"Edit first profile remark...");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_EDIT_FIRST_TAGS, L"Edit first profile tags...");
    AppendMenuW(codex_menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_EXPORT_BOX, L"Export .codexbox...");
    AppendMenuW(codex_menu, MF_STRING, ID_CODEX_IMPORT_BOX, L"Import .codexbox...");
    AppendMenuW(tools_menu, MF_POPUP, reinterpret_cast<UINT_PTR>(codex_menu), L"Codex Profiles");
    AppendMenuW(tools_menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(tools_menu, MF_STRING | MF_GRAYED, 0, L"FRP Tunnel (pending)");
    AppendMenuW(tools_menu, MF_STRING | MF_GRAYED, 0, L"Screenshot (pending)");
    AppendMenuW(system_menu, MF_STRING | MF_GRAYED, 0, L"Native modules pending");
    AppendMenuW(help_menu, MF_STRING, ID_HELP_ABOUT, L"About");

    AppendMenuW(menu_bar, MF_POPUP, reinterpret_cast<UINT_PTR>(file_menu), L"File");
    AppendMenuW(menu_bar, MF_POPUP, reinterpret_cast<UINT_PTR>(system_menu), L"System");
    AppendMenuW(menu_bar, MF_POPUP, reinterpret_cast<UINT_PTR>(tools_menu), L"Tools");
    AppendMenuW(menu_bar, MF_POPUP, reinterpret_cast<UINT_PTR>(help_menu), L"Help");

    SetMenu(window_, menu_bar);
}

void MainWindow::SwitchModule(ModuleId module_id) {
    if (module_id == ModuleId::CodexProfiles) {
        current_module_ = codex_profiles_.BuildModuleInfo();
    } else {
        current_module_ = BuildHomeModuleInfo();
    }

    if (logger_ != nullptr) {
        logger_->Info(L"Native module switched.");
    }
    InvalidateRect(window_, nullptr, FALSE);
}

void MainWindow::HandleCodexProfileAction(CodexProfileUiAction action) {
    SwitchModule(ModuleId::CodexProfiles);

    if (!ConfirmCodexProfileWrite(action)) {
        return;
    }

    CodexProfileActionOptions options;
    if (!PrepareCodexProfileActionOptions(action, &options)) {
        return;
    }

    const CodexProfileActionResult result = codex_profiles_.RunUiAction(action, options);
    SecureClearWideString(&options.password);
    const UINT icon = result.ok ? MB_ICONINFORMATION : MB_ICONWARNING;
    MessageBoxW(window_,
                result.message.empty() ? L"No detail returned." : result.message.c_str(),
                result.title.empty() ? L"Codex Profiles" : result.title.c_str(),
                MB_OK | icon);

    if (result.changed_state) {
        current_module_ = codex_profiles_.BuildModuleInfo();
        InvalidateRect(window_, nullptr, FALSE);
    }
}

bool MainWindow::PrepareCodexProfileActionOptions(CodexProfileUiAction action,
                                                  CodexProfileActionOptions* options) const {
    if (options == nullptr) {
        return false;
    }
    *options = CodexProfileActionOptions{};

    switch (action) {
        case CodexProfileUiAction::ExportBox:
            if (!PickCodexBoxSavePath(window_, &options->box_path)) {
                return false;
            }
            return PromptCodexBoxPassword(action, &options->password);

        case CodexProfileUiAction::ImportBox:
            if (!PickCodexBoxOpenPath(window_, &options->box_path)) {
                return false;
            }
            if (!PromptCodexBoxConflictPolicy(&options->box_conflict_policy)) {
                return false;
            }
            return PromptCodexBoxPassword(action, &options->password);

        case CodexProfileUiAction::ExportFirstProfileFiles:
            return PickFolderPath(window_,
                                  L"Select a folder for config.toml and auth.json export",
                                  &options->output_directory);

        case CodexProfileUiAction::RenameFirstProfile:
            return PromptCodexProfileDisplayName(&options->new_display_name);

        case CodexProfileUiAction::EditFirstProfileNote:
            return PromptCodexProfileMetadataText(action, &options->note);

        case CodexProfileUiAction::EditFirstProfileRemark:
            return PromptCodexProfileMetadataText(action, &options->remark);

        case CodexProfileUiAction::EditFirstProfileTags:
            return PromptCodexProfileMetadataText(action, &options->tags);

        default:
            return true;
    }
}

bool MainWindow::PromptCodexBoxPassword(CodexProfileUiAction action, std::wstring* password) const {
    if (password == nullptr) {
        return false;
    }
    password->clear();

    if (action == CodexProfileUiAction::ExportBox) {
        std::wstring first;
        std::wstring second;
        const bool first_ok = PromptPassword(window_,
                                             L"Export .codexbox",
                                             L"Enter a password for the encrypted .codexbox package:",
                                             &first);
        if (!first_ok) {
            SecureClearWideString(&first);
            return false;
        }
        const bool second_ok = PromptPassword(window_,
                                              L"Confirm .codexbox password",
                                              L"Enter the same password again:",
                                              &second);
        if (!second_ok) {
            SecureClearWideString(&first);
            SecureClearWideString(&second);
            return false;
        }
        if (first.empty() || first != second) {
            SecureClearWideString(&first);
            SecureClearWideString(&second);
            MessageBoxW(window_,
                        L"The .codexbox passwords were empty or did not match.",
                        L"Codex Profiles",
                        MB_OK | MB_ICONWARNING);
            return false;
        }

        *password = first;
        SecureClearWideString(&first);
        SecureClearWideString(&second);
        return true;
    }

    const bool ok = PromptPassword(window_,
                                   L"Import .codexbox",
                                   L"Enter the password for the selected .codexbox package:",
                                   password);
    if (!ok || password->empty()) {
        SecureClearWideString(password);
        if (ok) {
            MessageBoxW(window_,
                        L"The .codexbox import password cannot be empty.",
                        L"Codex Profiles",
                        MB_OK | MB_ICONWARNING);
        }
        return false;
    }
    return true;
}

bool MainWindow::PromptCodexBoxConflictPolicy(CodexProfileBoxConflictPolicy* policy) const {
    return ChooseCodexBoxConflictPolicy(window_, policy);
}

bool MainWindow::PromptCodexProfileDisplayName(std::wstring* display_name) const {
    if (display_name == nullptr) {
        return false;
    }
    display_name->clear();

    std::wstring raw_display_name;
    const bool ok = PromptText(window_,
                               L"Rename first Codex profile",
                               L"Enter the new display name for the first readable profile:",
                               &raw_display_name);
    if (!ok) {
        return false;
    }

    *display_name = TrimWide(raw_display_name);
    SecureClearWideString(&raw_display_name);
    if (display_name->empty()) {
        MessageBoxW(window_,
                    L"The Codex profile display name cannot be empty.",
                    L"Codex Profiles",
                    MB_OK | MB_ICONWARNING);
        return false;
    }
    if (display_name->size() > 120 || ContainsControlCharacter(*display_name)) {
        MessageBoxW(window_,
                    L"The Codex profile display name must be 120 characters or fewer and cannot contain control characters.",
                    L"Codex Profiles",
                    MB_OK | MB_ICONWARNING);
        return false;
    }
    return true;
}

bool MainWindow::PromptCodexProfileMetadataText(CodexProfileUiAction action,
                                                std::wstring* value) const {
    if (value == nullptr) {
        return false;
    }
    value->clear();

    const wchar_t* title = L"Edit first Codex profile metadata";
    const wchar_t* prompt = L"Enter the new value. Leave empty to clear the field:";
    size_t max_length = 500;
    switch (action) {
        case CodexProfileUiAction::EditFirstProfileNote:
            title = L"Edit first profile note";
            prompt = L"Enter note text for the first readable profile. Leave empty to clear it:";
            max_length = 500;
            break;
        case CodexProfileUiAction::EditFirstProfileRemark:
            title = L"Edit first profile remark";
            prompt = L"Enter remark text for the first readable profile. Leave empty to clear it:";
            max_length = 500;
            break;
        case CodexProfileUiAction::EditFirstProfileTags:
            title = L"Edit first profile tags";
            prompt = L"Enter tags for the first readable profile. Leave empty to clear them:";
            max_length = 200;
            break;
        default:
            return false;
    }

    std::wstring raw_value;
    const bool ok = PromptText(window_, title, prompt, &raw_value);
    if (!ok) {
        return false;
    }

    *value = TrimWide(raw_value);
    SecureClearWideString(&raw_value);
    if (value->size() > max_length || ContainsControlCharacter(*value)) {
        std::wstringstream message;
        message << L"The Codex profile metadata value must be " << max_length
                << L" characters or fewer and cannot contain control characters.";
        MessageBoxW(window_, message.str().c_str(), L"Codex Profiles", MB_OK | MB_ICONWARNING);
        return false;
    }
    return true;
}

bool MainWindow::ConfirmCodexProfileWrite(CodexProfileUiAction action) const {
    const wchar_t* message = nullptr;
    switch (action) {
        case CodexProfileUiAction::ApplyFirstProfile:
            message =
                L"This will back up the current .codex folder, then overwrite config.toml and auth.json with the first readable profile. Continue?";
            break;
        case CodexProfileUiAction::ImportCurrentFolder:
            message =
                L"This will read the current .codex config.toml/auth.json and append a DPAPI-protected profile to the local library. Continue?";
            break;
        case CodexProfileUiAction::RestoreLatestBackup:
            message =
                L"This will restore config.toml/auth.json from the latest native backup and overwrite the current .codex files. Continue?";
            break;
        case CodexProfileUiAction::ExportFirstProfileFiles:
            message =
                L"This will decrypt the first readable Codex profile in memory and export config.toml/auth.json to a folder you choose. Continue?";
            break;
        case CodexProfileUiAction::RenameFirstProfile:
            message =
                L"This will update the display name of the first readable Codex profile in the local DPAPI-protected library. Continue?";
            break;
        case CodexProfileUiAction::EditFirstProfileNote:
            message =
                L"This will update the note of the first readable Codex profile in the local DPAPI-protected library. Continue?";
            break;
        case CodexProfileUiAction::EditFirstProfileRemark:
            message =
                L"This will update the remark of the first readable Codex profile in the local DPAPI-protected library. Continue?";
            break;
        case CodexProfileUiAction::EditFirstProfileTags:
            message =
                L"This will update the tags of the first readable Codex profile in the local DPAPI-protected library. Continue?";
            break;
        case CodexProfileUiAction::ExportBox:
            message =
                L"This will decrypt saved Codex profiles in memory and write an encrypted .codexbox package. Continue?";
            break;
        case CodexProfileUiAction::ImportBox:
            message =
                L"This will import profiles from an encrypted .codexbox package into the local DPAPI-protected library. You will choose how name conflicts are handled before the package is read. Continue?";
            break;
        default:
            return true;
    }

    return MessageBoxW(window_,
                       message,
                       L"Confirm Codex profile action",
                       MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2) == IDYES;
}

void MainWindow::RestoreFromTray() {
    ShowWindow(window_, SW_SHOW);
    ShowWindow(window_, SW_RESTORE);
    SetForegroundWindow(window_);
    if (logger_ != nullptr) {
        logger_->Info(L"Window restored from tray.");
    }
}

void MainWindow::HideToTray() {
    ShowWindow(window_, SW_HIDE);
    if (logger_ != nullptr) {
        logger_->Info(L"Window hidden to tray.");
    }
}

void MainWindow::ExitApplication() {
    exiting_ = true;
    if (logger_ != nullptr) {
        logger_->Info(L"Exit requested.");
    }
    DestroyWindow(window_);
}

float MainWindow::CurrentDpiScale() const {
    return static_cast<float>(dpi_) / 96.0f;
}

}  // namespace mytools
