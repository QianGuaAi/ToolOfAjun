#pragma once

#include <string>

#include <windows.h>

#include "services/codex_profile_box_service.h"

namespace mytools {

void SecureClearWideString(std::wstring* value);

bool PromptPassword(HWND owner,
                    const std::wstring& title,
                    const std::wstring& prompt,
                    std::wstring* password);

bool PromptText(HWND owner,
                const std::wstring& title,
                const std::wstring& prompt,
                std::wstring* value);

bool PickCodexBoxSavePath(HWND owner, std::wstring* path);
bool PickCodexBoxOpenPath(HWND owner, std::wstring* path);
bool PickFolderPath(HWND owner, const std::wstring& title, std::wstring* path);
bool ChooseCodexBoxConflictPolicy(HWND owner, CodexProfileBoxConflictPolicy* policy);

}  // namespace mytools
