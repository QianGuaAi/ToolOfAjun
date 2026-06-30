param(
    [switch]$Quick,
    [switch]$Build,
    [switch]$Unit,
    [switch]$Smoke,
    [switch]$Installer
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$nativeRoot = Join-Path $repoRoot "src\MyTools.Native"
$buildRoot = Join-Path $repoRoot "artifacts\native\build"
$runAll = -not ($Quick -or $Build -or $Unit -or $Smoke -or $Installer)
$failed = $false

function Invoke-Step {
    param(
        [string]$Name,
        [scriptblock]$Command
    )

    Write-Host ""
    Write-Host "== $Name =="
    try {
        & $Command
        Write-Host "[PASS] $Name"
    } catch {
        Write-Host "[FAIL] $Name"
        Write-Host $_
        $script:failed = $true
    }
}

function Require-File {
    param([string]$RelativePath)

    $path = Join-Path $repoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing required file: $RelativePath"
    }
}

function Assert-NativeCodexMetadataEditGuards {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MainWindow,
        [Parameter(Mandatory = $true)]
        [string]$Modules,
        [Parameter(Mandatory = $true)]
        [string]$CodexEdit
    )

    $metadataFields = @(
        @{
            Label = "note";
            Action = "EditFirstProfileNote";
            Id = "ID_CODEX_EDIT_FIRST_NOTE";
            MenuText = "Edit first profile note...";
            PromptTitle = "Edit first profile note";
            StructToken = "std::wstring note";
            OptionArrow = "options->note";
            OptionDot = "options.note";
            UpdateFlag = "update_note";
            RequestField = "note";
            JsonProperty = '"Note"';
            Validated = "validated_note";
            Max = 500
        },
        @{
            Label = "remark";
            Action = "EditFirstProfileRemark";
            Id = "ID_CODEX_EDIT_FIRST_REMARK";
            MenuText = "Edit first profile remark...";
            PromptTitle = "Edit first profile remark";
            StructToken = "std::wstring remark";
            OptionArrow = "options->remark";
            OptionDot = "options.remark";
            UpdateFlag = "update_remark";
            RequestField = "remark";
            JsonProperty = '"Remark"';
            Validated = "validated_remark";
            Max = 500
        },
        @{
            Label = "tags";
            Action = "EditFirstProfileTags";
            Id = "ID_CODEX_EDIT_FIRST_TAGS";
            MenuText = "Edit first profile tags...";
            PromptTitle = "Edit first profile tags";
            StructToken = "std::wstring tags";
            OptionArrow = "options->tags";
            OptionDot = "options.tags";
            UpdateFlag = "update_tags";
            RequestField = "tags";
            JsonProperty = '"Tags"';
            Validated = "validated_tags";
            Max = 200
        }
    )

    foreach ($field in $metadataFields) {
        $uiTokens = @(
            $field.Id,
            $field.MenuText,
            "CodexProfileUiAction::$($field.Action)",
            "PromptCodexProfileMetadataText",
            "PromptCodexProfileMetadataText(action, &$($field.OptionArrow))",
            $field.PromptTitle,
            "characters or fewer",
            $field.OptionArrow
        )
        foreach ($token in $uiTokens) {
            if (-not $MainWindow.Contains($token)) {
                throw "Native Codex metadata UI missing $($field.Label) token: $token"
            }
        }
        $promptCasePattern = "case\s+CodexProfileUiAction::" + [regex]::Escape($field.Action) +
            "\s*:(?:(?!case\s+CodexProfileUiAction::).)*" +
            [regex]::Escape($field.PromptTitle) +
            "(?:(?!case\s+CodexProfileUiAction::).)*max_length\s*=\s*" + $field.Max
        if (-not [regex]::IsMatch($MainWindow, $promptCasePattern, [Text.RegularExpressions.RegexOptions]::Singleline)) {
            throw "Native Codex metadata prompt must bind $($field.Label) to $($field.Max) characters."
        }

        $moduleTokens = @(
            "enum class CodexProfileUiAction",
            $field.Action,
            $field.StructToken,
            $field.OptionDot,
            "case CodexProfileUiAction::$($field.Action)",
            "request.$($field.UpdateFlag) = true",
            "request.$($field.RequestField) = $($field.OptionDot)",
            "UpdateProfileMetadata"
        )
        foreach ($token in $moduleTokens) {
            if (-not $Modules.Contains($token)) {
                throw "Native Codex metadata module route missing $($field.Label) token: $token"
            }
        }
        $moduleRoutePattern = "case\s+CodexProfileUiAction::" + [regex]::Escape($field.Action) +
            "\s*:\s*request\." + [regex]::Escape($field.UpdateFlag) +
            "\s*=\s*true;\s*request\." + [regex]::Escape($field.RequestField) +
            "\s*=\s*options\." + [regex]::Escape($field.RequestField) + "\s*;"
        if (-not [regex]::IsMatch($Modules, $moduleRoutePattern, [Text.RegularExpressions.RegexOptions]::Singleline)) {
            throw "Native Codex metadata module route must bind $($field.Action) to $($field.UpdateFlag)/$($field.RequestField)."
        }

        $requestToken = "request.$($field.RequestField)"
        $trimToken = "TrimWide($requestToken)"
        $validateCallToken = "ValidateMetadataText($($field.Validated), $($field.Max)"
        $hasDirectValidation = $CodexEdit.Contains($field.Validated) `
            -and $CodexEdit.Contains($trimToken) `
            -and $CodexEdit.Contains("$($field.Validated).size() > $($field.Max)") `
            -and $CodexEdit.Contains("ContainsControlCharacter($($field.Validated))")
        $hasSharedValidation = $CodexEdit.Contains($field.Validated) `
            -and $CodexEdit.Contains($trimToken) `
            -and $CodexEdit.Contains($validateCallToken) `
            -and $CodexEdit.Contains("value.size() > max_length") `
            -and $CodexEdit.Contains("ContainsControlCharacter(value)") `
            -and $CodexEdit.Contains("characters or fewer")
        if (-not ($hasDirectValidation -or $hasSharedValidation)) {
            throw "Native Codex edit service must trim, length-check, and reject control characters for $($field.Label) metadata."
        }

        $editTokens = @(
            $field.UpdateFlag,
            $requestToken,
            $field.JsonProperty,
            "SetJsonStringProperty",
            "WideToUtf8($($field.Validated))"
        )
        foreach ($token in $editTokens) {
            if (-not $CodexEdit.Contains($token)) {
                throw "Native Codex edit service missing $($field.Label) token: $token"
            }
        }
    }

    if ($MainWindow.Contains("CodexProfileEditService") -or $MainWindow.Contains("UpdateProfileMetadata")) {
        throw "Native main window must route Codex metadata edits through CodexProfileModule::RunUiAction, not CodexProfileEditService."
    }
    foreach ($blockedText in @(
        "Renaming the active Codex profile is blocked",
        "active.json rename synchronization is migrated"
    )) {
        if ($MainWindow.Contains($blockedText) -or $Modules.Contains($blockedText) -or $CodexEdit.Contains($blockedText)) {
            throw "Native Codex profile rename must synchronize active.json instead of keeping old blocked wording: $blockedText"
        }
    }
    foreach ($forbiddenUiToken in @("active_json_path", "FileSystem::WriteUtf8FileAtomic", "SaveActiveDisplayName")) {
        if ($MainWindow.Contains($forbiddenUiToken)) {
            throw "Native main window must not write or coordinate active.json directly: $forbiddenUiToken"
        }
    }

    $activeRenameModuleTokens = @(
        "request.active_json_path = probe.paths.active_json",
        "request.active_display_name = probe.active_display_name",
        "result.active_json_updated",
        "Active profile marker updated in active.json"
    )
    foreach ($token in $activeRenameModuleTokens) {
        if (-not $Modules.Contains($token)) {
            throw "Native Codex active rename module route missing token: $token"
        }
    }

    $activeRenameServiceTokens = @(
        "std::wstring active_json_path",
        "std::wstring active_display_name",
        "bool active_json_updated",
        "SaveActiveDisplayName",
        "ReadUtf8File(active_json_path",
        '"ActiveDisplayName"',
        "WriteUtf8FileAtomic(active_json_path",
        "target_is_active",
        "request.active_json_path.empty()",
        "request.active_display_name"
    )
    foreach ($token in $activeRenameServiceTokens) {
        if (-not $CodexEdit.Contains($token)) {
            throw "Native Codex edit service missing active rename sync token: $token"
        }
    }
}

function Assert-NativeCodexBoxConflictGuards {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MainWindow,
        [Parameter(Mandatory = $true)]
        [string]$ModalDialogs,
        [Parameter(Mandatory = $true)]
        [string]$Modules,
        [Parameter(Mandatory = $true)]
        [string]$CodexBox
    )

    foreach ($token in @(
        "PromptCodexBoxConflictPolicy",
        "ChooseCodexBoxConflictPolicy(window_, policy)",
        "options->box_conflict_policy",
        "You will choose how name conflicts are handled"
    )) {
        if (-not $MainWindow.Contains($token)) {
            throw "Native Codex .codexbox conflict UI missing main-window token: $token"
        }
    }
    foreach ($token in @(
        "ChooseCodexBoxConflictPolicy",
        "RunConflictDialog",
        ".codexbox import conflicts",
        "kButtonRenameId",
        "kButtonSkipId",
        "kButtonReplaceId",
        "CodexProfileBoxConflictPolicy::Rename",
        "CodexProfileBoxConflictPolicy::Skip",
        "CodexProfileBoxConflictPolicy::Replace"
    )) {
        if (-not $ModalDialogs.Contains($token)) {
            throw "Native Codex .codexbox conflict dialog missing token: $token"
        }
    }
    foreach ($token in @(
        "CodexProfileBoxConflictPolicy box_conflict_policy",
        "request.conflict_policy = options.box_conflict_policy",
        "import asks how to handle name conflicts"
    )) {
        if (-not $Modules.Contains($token)) {
            throw "Native Codex .codexbox conflict module route missing token: $token"
        }
    }
    foreach ($token in @(
        "CodexProfileBoxConflictPolicy::Skip",
        "CodexProfileBoxConflictPolicy::Replace",
        "CodexProfileBoxConflictPolicy::Rename",
        "skipped_count",
        "replaced_count",
        "renamed_count"
    )) {
        if (-not $CodexBox.Contains($token)) {
            throw "Native Codex .codexbox service missing conflict policy token: $token"
        }
    }

    foreach ($forbidden in @(
        ".codexbox conflict choice dialog pending",
        "Name conflicts will be renamed automatically",
        "request.conflict_policy = CodexProfileBoxConflictPolicy::Rename;"
    )) {
        if ($MainWindow.Contains($forbidden) -or $Modules.Contains($forbidden)) {
            throw "Native Codex .codexbox import must use the selected conflict policy, not pending UI or hard-coded rename: $forbidden"
        }
    }
}

function Assert-NativeSourceLayout {
    $required = @(
        "src\MyTools.Native\CMakeLists.txt",
        "src\MyTools.Native\README.md",
        "src\MyTools.Native\app\main.cpp",
        "src\MyTools.Native\app\app_context.h",
        "src\MyTools.Native\app\app_context.cpp",
        "src\MyTools.Native\app\single_instance.h",
        "src\MyTools.Native\app\single_instance.cpp",
        "src\MyTools.Native\app\crash_handler.h",
        "src\MyTools.Native\app\crash_handler.cpp",
        "src\MyTools.Native\app\tray_host.h",
        "src\MyTools.Native\app\tray_host.cpp",
        "src\MyTools.Native\ccore\include\mt_result.h",
        "src\MyTools.Native\ccore\json\mt_json_schema.h",
        "src\MyTools.Native\ccore\json\mt_json_schema.c",
        "src\MyTools.Native\modules\module_registry.h",
        "src\MyTools.Native\modules\module_registry.cpp",
        "src\MyTools.Native\modules\codex_profiles\codex_profile_module.h",
        "src\MyTools.Native\modules\codex_profiles\codex_profile_module.cpp",
        "src\MyTools.Native\services\config_store.h",
        "src\MyTools.Native\services\config_store.cpp",
        "src\MyTools.Native\services\codex_profile_backup_service.h",
        "src\MyTools.Native\services\codex_profile_backup_service.cpp",
        "src\MyTools.Native\services\codex_profile_box_service.h",
        "src\MyTools.Native\services\codex_profile_box_service.cpp",
        "src\MyTools.Native\services\codex_profile_diff_service.h",
        "src\MyTools.Native\services\codex_profile_diff_service.cpp",
        "src\MyTools.Native\services\codex_profile_edit_service.h",
        "src\MyTools.Native\services\codex_profile_edit_service.cpp",
        "src\MyTools.Native\services\codex_profile_export_service.h",
        "src\MyTools.Native\services\codex_profile_export_service.cpp",
        "src\MyTools.Native\services\codex_profile_import_service.h",
        "src\MyTools.Native\services\codex_profile_import_service.cpp",
        "src\MyTools.Native\services\codex_profile_switch_service.h",
        "src\MyTools.Native\services\codex_profile_switch_service.cpp",
        "src\MyTools.Native\services\file_scanner.h",
        "src\MyTools.Native\services\file_scanner.cpp",
        "src\MyTools.Native\services\file_system.h",
        "src\MyTools.Native\services\file_system.cpp",
        "src\MyTools.Native\services\hotkey_service.h",
        "src\MyTools.Native\services\hotkey_service.cpp",
        "src\MyTools.Native\services\logger.h",
        "src\MyTools.Native\services\logger.cpp",
        "src\MyTools.Native\services\network_service.h",
        "src\MyTools.Native\services\network_service.cpp",
        "src\MyTools.Native\services\process_runner.h",
        "src\MyTools.Native\services\process_runner.cpp",
        "src\MyTools.Native\services\secret_store_dpapi.h",
        "src\MyTools.Native\services\secret_store_dpapi.cpp",
        "src\MyTools.Native\services\task_runner.h",
        "src\MyTools.Native\services\task_runner.cpp",
        "src\MyTools.Native\ui\main_window.h",
        "src\MyTools.Native\ui\main_window.cpp",
        "src\MyTools.Native\ui\modal_dialogs.h",
        "src\MyTools.Native\ui\modal_dialogs.cpp",
        "src\MyTools.Native\ui\renderer_d2d.h",
        "src\MyTools.Native\ui\renderer_d2d.cpp",
        "src\MyTools.Native\resources\app.rc.in",
        "src\MyTools.Native\resources\resource.h"
    )

    foreach ($file in $required) {
        Require-File $file
    }
}

function Assert-NativeSourceRules {
    $cmake = Get-Content -LiteralPath (Join-Path $nativeRoot "CMakeLists.txt") -Raw -Encoding UTF8
    $main = Get-Content -LiteralPath (Join-Path $nativeRoot "app\main.cpp") -Raw -Encoding UTF8
    $renderer = Get-Content -LiteralPath (Join-Path $nativeRoot "ui\renderer_d2d.cpp") -Raw -Encoding UTF8
    $secret = Get-Content -LiteralPath (Join-Path $nativeRoot "services\secret_store_dpapi.cpp") -Raw -Encoding UTF8
    $tray = Get-Content -LiteralPath (Join-Path $nativeRoot "app\tray_host.cpp") -Raw -Encoding UTF8
    $config = Get-Content -LiteralPath (Join-Path $nativeRoot "services\config_store.cpp") -Raw -Encoding UTF8
    $codexBackup = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_backup_service.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_backup_service.cpp") -Raw -Encoding UTF8)
    $codexBox = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_box_service.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_box_service.cpp") -Raw -Encoding UTF8)
    $codexDiff = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_diff_service.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_diff_service.cpp") -Raw -Encoding UTF8)
    $codexEdit = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_edit_service.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_edit_service.cpp") -Raw -Encoding UTF8)
    $codexExport = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_export_service.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_export_service.cpp") -Raw -Encoding UTF8)
    $codexImport = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_import_service.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_import_service.cpp") -Raw -Encoding UTF8)
    $codexSwitch = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_switch_service.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\codex_profile_switch_service.cpp") -Raw -Encoding UTF8)
    $fileSystem = Get-Content -LiteralPath (Join-Path $nativeRoot "services\file_system.cpp") -Raw -Encoding UTF8
    $scanner = Get-Content -LiteralPath (Join-Path $nativeRoot "services\file_scanner.cpp") -Raw -Encoding UTF8
    $hotkey = Get-Content -LiteralPath (Join-Path $nativeRoot "services\hotkey_service.cpp") -Raw -Encoding UTF8
    $network = Get-Content -LiteralPath (Join-Path $nativeRoot "services\network_service.cpp") -Raw -Encoding UTF8
    $processRunner = Get-Content -LiteralPath (Join-Path $nativeRoot "services\process_runner.cpp") -Raw -Encoding UTF8
    $taskRunner = (Get-Content -LiteralPath (Join-Path $nativeRoot "services\task_runner.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "services\task_runner.cpp") -Raw -Encoding UTF8)
    $jsonSchema = Get-Content -LiteralPath (Join-Path $nativeRoot "ccore\json\mt_json_schema.c") -Raw -Encoding UTF8
    $mainWindow = (Get-Content -LiteralPath (Join-Path $nativeRoot "ui\main_window.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "ui\main_window.cpp") -Raw -Encoding UTF8)
    $modalDialogs = (Get-Content -LiteralPath (Join-Path $nativeRoot "ui\modal_dialogs.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "ui\modal_dialogs.cpp") -Raw -Encoding UTF8)
    $rendererHeader = Get-Content -LiteralPath (Join-Path $nativeRoot "ui\renderer_d2d.h") -Raw -Encoding UTF8
    $modules = (Get-Content -LiteralPath (Join-Path $nativeRoot "modules\module_registry.cpp") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "modules\codex_profiles\codex_profile_module.h") -Raw -Encoding UTF8) +
        (Get-Content -LiteralPath (Join-Path $nativeRoot "modules\codex_profiles\codex_profile_module.cpp") -Raw -Encoding UTF8)

    $mustContain = @{
        "CMakeLists.txt" = @("LANGUAGES C CXX", "cxx_std_20", "d2d1", "dwrite", "bcrypt", "comdlg32", "crypt32", "ole32", "WIN32", "configure_file", "ws2_32", "ccore/json/mt_json_schema.c", "modules/codex_profiles/codex_profile_module.cpp", "modules/module_registry.cpp", "services/codex_profile_backup_service.cpp", "services/codex_profile_box_service.cpp", "services/codex_profile_diff_service.cpp", "services/codex_profile_edit_service.cpp", "services/codex_profile_export_service.cpp", "services/codex_profile_import_service.cpp", "services/codex_profile_switch_service.cpp", "ui/modal_dialogs.cpp")
        "app\main.cpp" = @("wWinMain", "SetProcessDpiAwarenessContext", "DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2")
        "ui\renderer_d2d.cpp" = @("D2D1CreateFactory", "DWriteCreateFactory")
        "services\secret_store_dpapi.cpp" = @("CryptProtectData", "CryptUnprotectData", "CryptBinaryToStringA", "CryptStringToBinaryA", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8")
        "app\tray_host.cpp" = @("Shell_NotifyIconW", "NIM_ADD", "NIM_DELETE")
        "services\config_store.cpp" = @("mt_json_has_schema_version", "schema_version", "WriteUtf8FileAtomic")
        "services\codex_profile_backup_service.cpp" = @("CreateCurrentFolderBackup", "RestoreLatestBackup", "native-codex-current-folder-before-switch", "ConfigTomlBase64", "AuthJsonBase64", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8", "CryptStringToBinaryA", "FindFirstFileW", "CompareFileTime", "WriteUtf8FileAtomic", "WriteFileBytesAtomic", "SecureZeroMemory", ".bak.dpapi")
        "services\codex_profile_box_service.cpp" = @("CodexProfileBoxService", "ExportBox", "ImportBox", "CodexProfileBoxConflictPolicy", "CDXB", "portable-codex-profiles-v2", "PBKDF2", "BCryptDeriveKeyPBKDF2", "BCRYPT_CHAIN_MODE_CBC", "BCRYPT_BLOCK_PADDING", "BCryptGenRandom", "HmacSha256", "ConstantTimeEquals", "200000", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlBase64", "AuthJsonBase64", "WriteFileBytesAtomic", "WriteUtf8FileAtomic", "SecureZeroMemory")
        "services\codex_profile_diff_service.cpp" = @("BuildProfileDiffSummary", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "UnprotectBase64ToUtf8", "ReadFileBytes", "FileSystem::Exists", "current_available", "BCryptOpenAlgorithmProvider", "BCRYPT_SHA256_ALGORITHM", "profile_sha256_hex", "current_sha256_hex", "profile_line_count", "current_line_count", "SecureZeroMemory")
        "services\codex_profile_edit_service.cpp" = @("UpdateProfileMetadata", "update_display_name", "update_note", "update_remark", "update_tags", "DisplayName", "Name", "Note", "Remark", "Tags", "UnprotectBase64ToUtf8", "ProtectUtf8ToBase64", "WriteUtf8FileAtomic", "CompareStringOrdinal", "SetJsonStringProperty", "TrimWide", "ContainsControlCharacter", "120 characters", "control characters", "0x9F", "SecureZeroMemory")
        "services\codex_profile_export_service.cpp" = @("ExportProfileByDisplayName", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "UnprotectBase64ToUtf8", "WriteUtf8FileAtomic", "config.toml", "auth.json", "SecureZeroMemory")
        "services\codex_profile_import_service.cpp" = @("ImportCurrentFolderProfile", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8", "ReadFileBytes", "WriteUtf8FileAtomic", "schemaVersion", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "LastImportedAt", "SecureZeroMemory", "config.toml", "auth.json")
        "services\codex_profile_switch_service.cpp" = @("ApplyProfileByDisplayName", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "CreateCurrentFolderBackup", "UnprotectBase64ToUtf8", "WriteUtf8FileAtomic", "ActiveDisplayName", "SwitchedAtUtc", "config.toml", "auth.json")
        "services\file_system.cpp" = @("MoveFileExW", "MOVEFILE_REPLACE_EXISTING", "FlushFileBuffers", "ReadFileBytes", "WriteFileBytesAtomic")
        "services\file_scanner.cpp" = @("FindFirstFileW", "FILE_ATTRIBUTE_REPARSE_POINT", "IsCancellationRequested")
        "services\hotkey_service.cpp" = @("RegisterHotKey", "UnregisterHotKey")
        "services\network_service.cpp" = @("WSAStartup", "GetAddrInfoW", "select")
        "services\process_runner.cpp" = @("CreateProcessW", "TerminateProcess")
        "services\task_runner.cpp" = @("std::condition_variable", "std::thread", "WorkerLoop")
        "ccore\json\mt_json_schema.c" = @("mt_json_has_schema_version", '\"schema_version\"')
        "ui\main_window.cpp" = @("ID_NAV_CODEX_PROFILES", "ID_CODEX_REFRESH", "ID_CODEX_DIFF_FIRST", "ID_CODEX_BACKUP_CURRENT", "ID_CODEX_APPLY_FIRST", "ID_CODEX_IMPORT_CURRENT", "ID_CODEX_RESTORE_BACKUP", "ID_CODEX_EXPORT_BOX", "ID_CODEX_IMPORT_BOX", "ID_CODEX_EXPORT_FIRST_FILES", "ID_CODEX_RENAME_FIRST", "HandleCodexProfileAction", "PrepareCodexProfileActionOptions", "PromptCodexBoxPassword", "PromptCodexBoxConflictPolicy", "PromptCodexProfileDisplayName", "PromptText", "PickCodexBoxSavePath", "PickCodexBoxOpenPath", "PickFolderPath", "TrimWide", "ContainsControlCharacter", "SecureClearWideString", "ConfirmCodexProfileWrite", "MB_YESNO", "MB_DEFBUTTON2", "SwitchModule", "Codex Profiles", "Export first profile files...", "Rename first profile...", "CodexProfileUiAction::ExportFirstProfileFiles", "CodexProfileUiAction::RenameFirstProfile", "options->output_directory", "options->new_display_name", "options->box_conflict_policy", "120 characters", "control characters", "0x9F", "FRP Tunnel (pending)")
        "ui\modal_dialogs.cpp" = @("PromptPassword", "PromptText", "PickFolderPath", "ChooseCodexBoxConflictPolicy", "SecureClearWideString", "GetSaveFileNameW", "GetOpenFileNameW", "SHBrowseForFolderW", "SHGetPathFromIDListW", "CoInitializeEx", "COINIT_APARTMENTTHREADED", "CoTaskMemFree", "CoUninitialize", "OFN_OVERWRITEPROMPT", "OFN_FILEMUSTEXIST", "ES_PASSWORD", "EM_SETPASSWORDCHAR", "EnableWindow", "Codex profile package (*.codexbox)")
        "ui\renderer_d2d.h" = @("ModuleInfo")
        "modules\codex_profiles" = @("LOCALAPPDATA", "USERPROFILE", "profiles.json", "active.json", "Backups", "config.toml", "auth.json", "UnprotectBase64ToUtf8", "ExtractJsonStringValues", "ExtractJsonObjectArray", "ReadHexCodeUnit", "AppendUtf8CodePoint", "ActiveDisplayName", "active profile", "Switch backup directory", "explicit menu actions", ".codexbox", "DisplayName", "Status", "LastImportedAt", "first_profile_display_name", "CodexProfileUiAction", "RunUiAction", "DiffFirstProfile", "BackupCurrentFolder", "ApplyFirstProfile", "ImportCurrentFolder", "RestoreLatestBackup", "ExportBox", "ImportBox", "ExportFirstProfileFiles", "RenameFirstProfile", "std::wstring output_directory", "std::wstring new_display_name", "options.output_directory", "options.new_display_name", "options.box_conflict_policy", "request.output_directory", "request.update_display_name", "request.new_display_name", "request.active_json_path", "request.active_display_name", "request.conflict_policy = options.box_conflict_policy", "SameDirectoryPath(options.output_directory, probe.paths.codex_home)", "Active profile marker updated", "CodexProfileExportService", "CodexProfileEditService", "ExportProfileByDisplayName", "UpdateProfileMetadata", "without exposing embedded auth/config secrets")
    }

    foreach ($token in $mustContain["CMakeLists.txt"]) {
        if (-not $cmake.Contains($token)) { throw "Native CMake missing token: $token" }
    }
    foreach ($token in $mustContain["app\main.cpp"]) {
        if (-not $main.Contains($token)) { throw "Native main missing token: $token" }
    }
    foreach ($token in $mustContain["ui\renderer_d2d.cpp"]) {
        if (-not $renderer.Contains($token)) { throw "Native renderer missing token: $token" }
    }
    foreach ($token in $mustContain["services\secret_store_dpapi.cpp"]) {
        if (-not $secret.Contains($token)) { throw "Native DPAPI store missing token: $token" }
    }
    foreach ($token in $mustContain["app\tray_host.cpp"]) {
        if (-not $tray.Contains($token)) { throw "Native tray host missing token: $token" }
    }
    foreach ($token in $mustContain["services\config_store.cpp"]) {
        if (-not $config.Contains($token)) { throw "Native config store missing token: $token" }
    }
    foreach ($token in $mustContain["services\codex_profile_backup_service.cpp"]) {
        if (-not $codexBackup.Contains($token)) { throw "Native Codex backup service missing token: $token" }
    }
    foreach ($token in $mustContain["services\codex_profile_box_service.cpp"]) {
        if (-not $codexBox.Contains($token)) { throw "Native Codex .codexbox service missing token: $token" }
    }
    foreach ($token in $mustContain["services\codex_profile_diff_service.cpp"]) {
        if (-not $codexDiff.Contains($token)) { throw "Native Codex diff service missing token: $token" }
    }
    foreach ($token in $mustContain["services\codex_profile_edit_service.cpp"]) {
        if (-not $codexEdit.Contains($token)) { throw "Native Codex edit service missing token: $token" }
    }
    foreach ($token in $mustContain["services\codex_profile_export_service.cpp"]) {
        if (-not $codexExport.Contains($token)) { throw "Native Codex export service missing token: $token" }
    }
    foreach ($token in $mustContain["services\codex_profile_import_service.cpp"]) {
        if (-not $codexImport.Contains($token)) { throw "Native Codex import service missing token: $token" }
    }
    foreach ($token in $mustContain["services\codex_profile_switch_service.cpp"]) {
        if (-not $codexSwitch.Contains($token)) { throw "Native Codex switch service missing token: $token" }
    }
    foreach ($token in $mustContain["services\file_system.cpp"]) {
        if (-not $fileSystem.Contains($token)) { throw "Native file system missing token: $token" }
    }
    foreach ($token in $mustContain["services\file_scanner.cpp"]) {
        if (-not $scanner.Contains($token)) { throw "Native file scanner missing token: $token" }
    }
    foreach ($token in $mustContain["services\hotkey_service.cpp"]) {
        if (-not $hotkey.Contains($token)) { throw "Native hotkey service missing token: $token" }
    }
    foreach ($token in $mustContain["services\network_service.cpp"]) {
        if (-not $network.Contains($token)) { throw "Native network service missing token: $token" }
    }
    foreach ($token in $mustContain["services\process_runner.cpp"]) {
        if (-not $processRunner.Contains($token)) { throw "Native process runner missing token: $token" }
    }
    foreach ($token in $mustContain["services\task_runner.cpp"]) {
        if (-not $taskRunner.Contains($token)) { throw "Native task runner missing token: $token" }
    }
    foreach ($token in $mustContain["ccore\json\mt_json_schema.c"]) {
        if (-not $jsonSchema.Contains($token)) { throw "Native C json schema helper missing token: $token" }
    }
    foreach ($token in $mustContain["ui\main_window.cpp"]) {
        if (-not $mainWindow.Contains($token)) { throw "Native main window module navigation missing token: $token" }
    }
    foreach ($token in $mustContain["ui\modal_dialogs.cpp"]) {
        if (-not $modalDialogs.Contains($token)) { throw "Native modal dialog helper missing token: $token" }
    }
    foreach ($token in $mustContain["ui\renderer_d2d.h"]) {
        if (-not $rendererHeader.Contains($token)) { throw "Native renderer module contract missing token: $token" }
    }
    foreach ($token in $mustContain["modules\codex_profiles"]) {
        if (-not $modules.Contains($token)) { throw "Native Codex profile scaffold missing token: $token" }
    }
    Assert-NativeCodexMetadataEditGuards -MainWindow $mainWindow -Modules $modules -CodexEdit $codexEdit
    Assert-NativeCodexBoxConflictGuards -MainWindow $mainWindow -ModalDialogs $modalDialogs -Modules $modules -CodexBox $codexBox

    $combined = $cmake + $main + $renderer + $secret + $tray + $config + $codexBackup + $codexBox + $codexDiff + $codexEdit + $codexExport + $codexImport + $codexSwitch + $fileSystem + $scanner + $hotkey + $network + $processRunner + $taskRunner + $jsonSchema + $mainWindow + $modalDialogs + $rendererHeader + $modules
    foreach ($blocked in @("Qt", "Electron", "WebView2", "WindowsAppSDK", "Microsoft.NET")) {
        if ($combined.Contains($blocked)) {
            throw "Native phase 1/2 must not introduce blocked runtime token: $blocked"
        }
    }
    foreach ($forbidden in @("ProtectedAuthJsonBase64", "AuthJsonContentProtected", "ProtectedConfigTomlBase64", "ConfigTomlContentProtected", "access_token", "refresh_token", "WinHttp", "HttpSendRequest", "WriteUtf8FileAtomic", "CreateProcessW", "ProbeTcp")) {
        if ($modules.Contains($forbidden)) {
            throw "Native Codex profile UI action layer must not expose secrets, network, or low-level file writes directly: $forbidden"
        }
    }
    foreach ($forbidden in @("access_token", "refresh_token", "WinHttp", "HttpSendRequest", "GetAddrInfoW", "ProbeTcp", "CreateProcessW", "Logger", "AppLog")) {
        if ($codexBackup.Contains($forbidden)) {
            throw "Native Codex backup/restore service must not parse tokens, start processes, contact networks, or log secret-bearing payloads: $forbidden"
        }
    }
    foreach ($forbidden in @("access_token", "refresh_token", "WinHttp", "HttpSendRequest", "GetAddrInfoW", "ProbeTcp", "CreateProcessW", "Logger", "AppLog", "ZipArchive", "miniz", "AesGcm")) {
        if ($codexBox.Contains($forbidden)) {
            throw "Native Codex .codexbox service must not parse tokens, start processes, contact networks, log secret-bearing payloads, or change to a ZIP/GCM package format: $forbidden"
        }
    }
    foreach ($forbidden in @("access_token", "refresh_token", "WinHttp", "HttpSendRequest", "GetAddrInfoW", "ProbeTcp", "CreateProcessW", "Logger", "AppLog", "WriteUtf8FileAtomic", "WriteFileBytesAtomic")) {
        if ($codexDiff.Contains($forbidden)) {
            throw "Native Codex diff service must not parse tokens, write files, start processes, contact networks, or log secret-bearing payloads: $forbidden"
        }
    }
    foreach ($forbidden in @("access_token", "refresh_token", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "WinHttp", "HttpSendRequest", "GetAddrInfoW", "ProbeTcp", "CreateProcessW", "Logger", "AppLog", "WriteFileBytesAtomic")) {
        if ($codexEdit.Contains($forbidden)) {
            throw "Native Codex edit service must not parse or name protected config/auth fields, start processes, contact networks, log secret-bearing payloads, or overwrite current Codex files: $forbidden"
        }
    }
    foreach ($forbidden in @("access_token", "refresh_token", "WinHttp", "HttpSendRequest", "GetAddrInfoW", "ProbeTcp", "CreateProcessW", "Logger", "AppLog")) {
        if ($codexExport.Contains($forbidden)) {
            throw "Native Codex export service must not parse tokens, start processes, contact networks, or log secret-bearing payloads: $forbidden"
        }
    }
    foreach ($forbidden in @("access_token", "refresh_token", "WinHttp", "HttpSendRequest", "GetAddrInfoW", "ProbeTcp", "CreateProcessW", "Logger", "AppLog", "WriteFileBytesAtomic")) {
        if ($codexImport.Contains($forbidden)) {
            throw "Native Codex import service must not parse tokens, start processes, contact networks, log secret-bearing payloads, or overwrite current Codex binaries: $forbidden"
        }
    }
    foreach ($forbidden in @("access_token", "refresh_token", "WinHttp", "HttpSendRequest", "GetAddrInfoW", "ProbeTcp", "CreateProcessW", "Logger", "AppLog")) {
        if ($codexSwitch.Contains($forbidden)) {
            throw "Native Codex switch service must not parse tokens, start processes, contact networks, or log secret-bearing payloads: $forbidden"
        }
    }
    if ($mainWindow.Contains("ApplyProfileByDisplayName")) {
        throw "Native main window must route Codex profile switching through CodexProfileModule explicit actions, not service calls."
    }
    if ($mainWindow.Contains("CodexProfileBackupService") -or $mainWindow.Contains("RestoreLatestBackup(")) {
        throw "Native main window must route Codex backup restore through CodexProfileModule explicit actions, not service calls."
    }
    if ($mainWindow.Contains("ExportProfileByDisplayName")) {
        throw "Native main window must route Codex profile export through CodexProfileModule explicit actions, not service calls."
    }
    if ($mainWindow.Contains("UpdateProfileMetadata")) {
        throw "Native main window must route Codex metadata edits through CodexProfileModule explicit actions, not service calls."
    }
    if ($mainWindow.Contains("ImportCurrentFolderProfile")) {
        throw "Native main window must route Codex profile import through CodexProfileModule explicit actions, not service calls."
    }
    if ($mainWindow.Contains("CodexProfileBoxService") -or $mainWindow.Contains(".ExportBox(") -or $mainWindow.Contains(".ImportBox(")) {
        throw "Native main window must route Codex .codexbox import/export through CodexProfileModule explicit actions, not service calls."
    }
    if ($mainWindow.Contains("BuildProfileDiffSummary")) {
        throw "Native main window navigation must not diff Codex profiles until an explicit reviewed UI action exists."
    }
}

function Resolve-CMake {
    $command = Get-Command cmake -ErrorAction SilentlyContinue
    if (-not $command) {
        throw "cmake was not found. Install CMake and MSVC Build Tools before running native build validation."
    }
    return $command.Source
}

function Invoke-NativeBuild {
    $cmake = Resolve-CMake
    New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null

    & $cmake -S $nativeRoot -B $buildRoot -DCMAKE_BUILD_TYPE=Release
    if ($LASTEXITCODE -ne 0) {
        throw "CMake configure failed with exit code $LASTEXITCODE"
    }

    & $cmake --build $buildRoot --config Release
    if ($LASTEXITCODE -ne 0) {
        throw "Native build failed with exit code $LASTEXITCODE"
    }
}

function Invoke-NativeSmoke {
    $candidates = @(
        Join-Path $buildRoot "Release\MyToolsNative.exe",
        Join-Path $buildRoot "MyToolsNative.exe"
    )
    $exe = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $exe) {
        throw "Native executable not found. Run -Build first."
    }

    $item = Get-Item -LiteralPath $exe
    Write-Host "Native exe: $($item.FullName)"
    Write-Host "Size: $($item.Length) bytes"
}

if ($runAll -or $Quick) {
    Invoke-Step "native source layout" { Assert-NativeSourceLayout }
    Invoke-Step "native source rules" { Assert-NativeSourceRules }
}

if ($runAll -or $Build) {
    Invoke-Step "native build" { Invoke-NativeBuild }
}

if ($runAll -or $Unit) {
    Invoke-Step "native unit placeholder" {
        Assert-NativeSourceLayout
        Write-Host "No standalone native unit tests exist yet; native phase 1/2/3 scaffold uses source guards and build validation."
    }
}

if ($runAll -or $Smoke) {
    Invoke-Step "native smoke" { Invoke-NativeSmoke }
}

if ($Installer) {
    Invoke-Step "native installer placeholder" {
        throw "Native installer is planned for phase 10 and is not implemented in phase 1."
    }
}

if ($failed) {
    throw "native-eval: failed"
}

Write-Host ""
Write-Host "native-eval: passed"
