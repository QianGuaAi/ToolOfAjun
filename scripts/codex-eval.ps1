param(
    [switch]$Quick,
    [switch]$Build,
    [switch]$Installer,
    [string]$InstallerVersion = "2026.6.30.0"
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$runAll = -not ($Quick -or $Build -or $Installer)
$failed = $false

function Invoke-Step {
    param(
        [string]$Name,
        [string]$WorkingDirectory,
        [scriptblock]$Command
    )

    Write-Host ""
    Write-Host "== $Name =="
    Push-Location $WorkingDirectory
    try {
        & $Command
        if ($LASTEXITCODE -ne $null -and $LASTEXITCODE -ne 0) {
            throw "Command exited with code $LASTEXITCODE"
        }
        Write-Host "[PASS] $Name"
    } catch {
        Write-Host "[FAIL] $Name"
        Write-Host $_
        $script:failed = $true
    } finally {
        Pop-Location
    }
}

function Resolve-Dotnet {
    $localDotnet = Join-Path $repoRoot ".dotnet\dotnet.exe"
    if (Test-Path -LiteralPath $localDotnet) {
        $env:DOTNET_ROOT = Join-Path $repoRoot ".dotnet"
        $env:PATH = "$env:DOTNET_ROOT;$env:PATH"
        return $localDotnet
    }

    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($dotnetCommand) {
        return $dotnetCommand.Source
    }

    throw "dotnet not found. Use repo-local .dotnet or add dotnet to PATH."
}

function Test-FileContainsAsciiText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $bytes = [IO.File]::ReadAllBytes($Path)
    $needle = [Text.Encoding]::ASCII.GetBytes($Text)
    if ($bytes.Length -lt $needle.Length) {
        return $false
    }

    for ($i = 0; $i -le $bytes.Length - $needle.Length; $i++) {
        $matched = $true
        for ($j = 0; $j -lt $needle.Length; $j++) {
            if ($bytes[$i + $j] -ne $needle[$j]) {
                $matched = $false
                break
            }
        }

        if ($matched) {
            return $true
        }
    }

    return $false
}

function Assert-InstallerArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedVersion
    )

    Add-Type -AssemblyName System.IO.Compression

    $setupPath = Join-Path $repoRoot "artifacts\installer\MyToolsSetup.exe"
    $payloadZipPath = Join-Path $repoRoot "artifacts\installer\MyToolsPayload.zip"
    if (-not (Test-Path -LiteralPath $setupPath)) {
        throw "Installer artifact missing: $setupPath"
    }
    if (-not (Test-Path -LiteralPath $payloadZipPath)) {
        throw "Payload zip missing: $payloadZipPath"
    }

    $setupItem = Get-Item -LiteralPath $setupPath
    if ($setupItem.VersionInfo.FileVersion -ne $ExpectedVersion) {
        throw "Installer version mismatch. Expected $ExpectedVersion, got $($setupItem.VersionInfo.FileVersion)."
    }

    $payloadZipItem = Get-Item -LiteralPath $payloadZipPath
    if ($setupItem.Length -le 0 -or $payloadZipItem.Length -le 0) {
        throw "Installer artifact or payload zip is empty."
    }

    $setupAssembly = [Reflection.Assembly]::LoadFrom($setupPath)
    $resources = $setupAssembly.GetManifestResourceNames() | Sort-Object
    $requiredResources = @(
        "MyToolsPayload.zip",
        "MyTools.Uninstaller.exe"
    )
    foreach ($resource in $requiredResources) {
        if ($resources -notcontains $resource) {
            throw "Installer missing embedded resource: $resource"
        }

        $resourceStream = $setupAssembly.GetManifestResourceStream($resource)
        try {
            if ($null -eq $resourceStream -or $resourceStream.Length -le 0) {
                throw "Installer embedded resource is empty: $resource"
            }
        } finally {
            if ($null -ne $resourceStream) {
                $resourceStream.Dispose()
            }
        }
    }

    $embeddedPayloadStream = $setupAssembly.GetManifestResourceStream("MyToolsPayload.zip")
    try {
        $externalPayloadSha = (Get-FileHash -Algorithm SHA256 -LiteralPath $payloadZipPath).Hash
        $embeddedPayloadSha = Get-StreamSha256Hex -Stream $embeddedPayloadStream
        if ($externalPayloadSha -ne $embeddedPayloadSha) {
            throw "Embedded payload hash mismatch. External=$externalPayloadSha Embedded=$embeddedPayloadSha."
        }
    } finally {
        if ($null -ne $embeddedPayloadStream) {
            $embeddedPayloadStream.Dispose()
        }
    }

    $payloadStream = $setupAssembly.GetManifestResourceStream("MyToolsPayload.zip")
    $payloadArchive = $null
    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ("MyToolsInstallerEval\" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    try {
        $payloadArchive = [IO.Compression.ZipArchive]::new($payloadStream, [IO.Compression.ZipArchiveMode]::Read, $false)
        $entries = $payloadArchive.Entries | ForEach-Object { $_.FullName.Replace("/", "\") } | Sort-Object
        $requiredEntries = @(
            "MyTools.exe",
            "MyTools.exe.config",
            "LockWin10_22H2.ps1",
            "NativeBinaries\README.txt",
            "NativeBinaries\ffmpeg\README.txt"
        )

        foreach ($entry in $requiredEntries) {
            if ($entries -notcontains $entry) {
                throw "Embedded payload missing required file: $entry"
            }
        }
        if ($entries -contains "NativeBinaries\ffmpeg\ffmpeg.exe") {
            throw "Embedded payload must not include bundled ffmpeg.exe; FFmpeg is an external dependency."
        }

        $mainEntry = $payloadArchive.GetEntry("MyTools.exe")
        if ($null -eq $mainEntry -or $mainEntry.Length -le 0) {
            throw "Embedded payload MyTools.exe is empty."
        }

        $mainExtractPath = Join-Path $tempDir "MyTools.exe"
        $inputStream = $mainEntry.Open()
        $outputStream = [IO.File]::Create($mainExtractPath)
        try {
            $inputStream.CopyTo($outputStream)
        } finally {
            $outputStream.Dispose()
            $inputStream.Dispose()
        }

        $requiredTypeNames = @(
            "CodexLocalRelayCompatibility",
            "CodexLocalRelayService"
        )
        foreach ($typeName in $requiredTypeNames) {
            if (-not (Test-FileContainsAsciiText -Path $mainExtractPath -Text $typeName)) {
                throw "Embedded payload MyTools.exe does not contain expected type name: $typeName"
            }
        }
    } finally {
        if ($null -ne $payloadArchive) {
            $payloadArchive.Dispose()
        } elseif ($null -ne $payloadStream) {
            $payloadStream.Dispose()
        }

        if (Test-Path -LiteralPath $tempDir) {
            Remove-Item -LiteralPath $tempDir -Recurse -Force
        }
    }
}

function Get-StreamSha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [IO.Stream]$Stream
    )

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        if ($Stream.CanSeek) {
            $Stream.Position = 0
        }

        $hashBytes = $sha.ComputeHash($Stream)
        return ([BitConverter]::ToString($hashBytes)).Replace("-", "")
    } finally {
        $sha.Dispose()
    }
}

function Assert-LoopMemoryDocs {
    function Join-CodePointString {
        param([int[]]$CodePoints)
        return -join ($CodePoints | ForEach-Object { [string][char]$_ })
    }

    $memoryIndexName = (Join-CodePointString @(0x667A, 0x80FD, 0x52A9, 0x624B, 0x8BB0, 0x5FC6, 0x7D22, 0x5F15)) + ".md"
    $memoryDepositName = "LoopEngineering" + (Join-CodePointString @(0x8BB0, 0x5FC6, 0x6C89, 0x6DC0)) + ".md"
    $memoryIndexPath = Join-Path "docs" $memoryIndexName
    $memoryDepositPath = Join-Path "docs" $memoryDepositName
    $memoryCandidateText = Join-CodePointString @(0x8BB0, 0x5FC6, 0x5019, 0x9009)

    $requiredPaths = @(
        $memoryIndexPath,
        $memoryDepositPath,
        ".agents\skills\mytools-loop-engineering\SKILL.md",
        ".codex\agents\mytools-reviewer.toml"
    )

    foreach ($relativePath in $requiredPaths) {
        $path = Join-Path $repoRoot $relativePath
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Loop Engineering memory document missing: $relativePath"
        }
    }

    $skill = Get-Content -LiteralPath (Join-Path $repoRoot ".agents\skills\mytools-loop-engineering\SKILL.md") -Raw -Encoding UTF8
    if (-not $skill.Contains($memoryIndexPath) -or -not $skill.Contains($memoryDepositName.Replace(".md", ""))) {
        throw "MyTools Loop Engineering skill must route Inspect/Deposit through the memory index and deposit guide."
    }
    Assert-LoopMultiAgentPolicy $skill (Get-Content -LiteralPath (Join-Path $repoRoot "AGENTS.md") -Raw -Encoding UTF8)

    $reviewer = Get-Content -LiteralPath (Join-Path $repoRoot ".codex\agents\mytools-reviewer.toml") -Raw -Encoding UTF8
    if (-not $reviewer.Contains($memoryCandidateText)) {
        throw "mytools-reviewer must report a memory candidate line."
    }

    Assert-MemoryIndexReferencesExist
}

function Assert-StartupResponsivenessGuards {
    $mainViewModelPath = Join-Path $repoRoot "src\MyTools\ViewModels\MainViewModel.cs"
    $source = Get-Content -LiteralPath $mainViewModelPath -Raw -Encoding UTF8

    $startupMethodMatch = [regex]::Match(
        $source,
        "private\s+void\s+ScheduleStartupBackgroundLoads\s*\(\s*\)\s*\{(?<body>.*?)^\s*\}",
        [Text.RegularExpressions.RegexOptions]::Singleline -bor [Text.RegularExpressions.RegexOptions]::Multiline)
    if (-not $startupMethodMatch.Success) {
        throw "Could not locate ScheduleStartupBackgroundLoads in MainViewModel.cs."
    }

    $startupBody = $startupMethodMatch.Groups["body"].Value
    if ($startupBody.Contains("HardwareSensor") -or $startupBody.Contains("StartSensorsBackground")) {
        throw "Startup background loads must not start hardware sensors; the sensor module has been removed."
    }

    $switchModuleMatch = [regex]::Match(
        $source,
        "private\s+void\s+SwitchModule\s*\(\s*string\s+module\s*\)\s*\{(?<body>.*?)^\s*\}\s*public\s+void\s+SwitchToHomeFromMultimedia",
        [Text.RegularExpressions.RegexOptions]::Singleline -bor [Text.RegularExpressions.RegexOptions]::Multiline)
    if (-not $switchModuleMatch.Success) {
        throw "Could not locate SwitchModule in MainViewModel.cs."
    }
    $hasSensorRuntimeReference = $switchModuleMatch.Groups["body"].Value.Contains("StartSensorsBackgroundAsync") `
        -or $source.Contains("HardwareSensorService") `
        -or $source.Contains("RefreshSensorsAsync")
    if ($hasSensorRuntimeReference) {
        throw "Hardware sensor service and polling methods must stay removed."
    }
}

function Assert-SqlExportRemoved {
    $deletedFiles = @(
        "src\MyTools\Services\SqlConnectionHistoryService.cs",
        "src\MyTools\Services\SqlExportModels.cs",
        "src\MyTools\Services\SqlExportProvider.cs",
        "src\MyTools\Services\SqlExportService.cs",
        "src\MyTools\Services\SqlServerExportProvider.cs"
    )
    foreach ($relativePath in $deletedFiles) {
        $path = Join-Path $repoRoot $relativePath
        if (Test-Path -LiteralPath $path) {
            throw "SQL export tool must stay removed; found deleted file: $relativePath"
        }
    }

    $checks = @{
        "src\MyTools\MainWindow.xaml" = @("ShowSqlExportCommand", 'Value="SqlExport"', "SQL 导出")
        "src\MyTools\MainWindow.xaml.cs" = @("CurrentSqlPasswordBox", "CurrentSqlQueryResultGrid", "SqlQuery")
        "src\MyTools\Views\MainModulesView.xaml" = @("SqlExportHost", 'Value="SqlExport"', "SQL 导出工具", "SQL 查询")
        "src\MyTools\Views\MainModulesView.xaml.cs" = @("CurrentSqlPasswordBox", "CurrentSqlQueryResultGrid", "SqlQuery")
        "src\MyTools\ViewModels\MainViewModel.cs" = @("ShowSqlExportCommand", "SqlExport", "SqlConnectionHistoryService", "SqlExportService", "SqlProviderKind", "SqlServerAddress", "ExecuteQuery")
        "src\MyTools\Services\SystemBackupService.cs" = @("MyTools.sqlhistory.json")
        "src\MyTools\Resources\Illustrations.xaml" = @("SqlEmpty")
    }

    foreach ($entry in $checks.GetEnumerator()) {
        $path = Join-Path $repoRoot $entry.Key
        if (-not (Test-Path -LiteralPath $path)) {
            throw "SQL export removal guard source missing: $($entry.Key)"
        }

        $source = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        foreach ($token in $entry.Value) {
            if ($source.Contains($token)) {
                throw "SQL export tool must stay removed; found token '$token' in $($entry.Key)."
            }
        }
    }
}

function Assert-LoopMultiAgentPolicy {
    param(
        [string]$Skill,
        [string]$Agents
    )

    function Join-CodePointString {
        param([int[]]$CodePoints)
        return -join ($CodePoints | ForEach-Object { [string][char]$_ })
    }

    $reviewRoleText = Join-CodePointString @(0x68C0, 0x67E5, 0x2F, 0x534F, 0x8C03, 0x2F, 0x9A8C, 0x6536)
    $multiAgentSplitText = Join-CodePointString @(0x591A, 0x4EE3, 0x7406, 0x5206, 0x5DE5)
    $microSingleText = (Join-CodePointString @(0x5FAE, 0x5C0F, 0x4EFB, 0x52A1, 0x53EF, 0x5355)) + " agent"
    $nonTrivialMultiText = Join-CodePointString @(0x975E, 0x5FAE, 0x5C0F, 0x4EE3, 0x7801, 0x3001, 0x89C4, 0x5219, 0x3001, 0x53D1, 0x5E03, 0x3001, 0x8DE8, 0x6587, 0x4EF6, 0x6216, 0x8DE8, 0x6A21, 0x5757, 0x4EFB, 0x52A1, 0x9ED8, 0x8BA4, 0x591A, 0x4EE3, 0x7406, 0x6267, 0x884C)
    $primaryIntegrationText = (Join-CodePointString @(0x4E3B)) + " agent " + (Join-CodePointString @(0x8D1F, 0x8D23, 0x6700, 0x7EC8, 0x6574, 0x5408))

    $requiredSkillTokens = @(
        "multi-agent role split",
        "Micro tasks",
        "Non-trivial code, rule, release, cross-file, or cross-module tasks default to multi-agent execution",
        $reviewRoleText,
        "primary agent owns final integration"
    )
    foreach ($token in $requiredSkillTokens) {
        if (-not $Skill.Contains($token)) {
            throw "MyTools Loop Engineering skill missing multi-agent policy token: $token"
        }
    }

    $requiredAgentsTokens = @(
        $multiAgentSplitText,
        $microSingleText,
        $nonTrivialMultiText,
        $reviewRoleText,
        $primaryIntegrationText
    )
    foreach ($token in $requiredAgentsTokens) {
        if (-not $Agents.Contains($token)) {
            throw "AGENTS.md missing multi-agent policy token: $token"
        }
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
            MenuText = "Edit profile note...";
            PromptTitle = "Edit profile note";
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
            MenuText = "Edit profile remark...";
            PromptTitle = "Edit profile remark";
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
            MenuText = "Edit profile tags...";
            PromptTitle = "Edit profile tags";
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

function Assert-MemoryIndexReferencesExist {
    function Join-CodePointString {
        param([int[]]$CodePoints)
        return -join ($CodePoints | ForEach-Object { [string][char]$_ })
    }

    $memoryIndexName = (Join-CodePointString @(0x667A, 0x80FD, 0x52A9, 0x624B, 0x8BB0, 0x5FC6, 0x7D22, 0x5F15)) + ".md"
    $memoryIndexPath = Join-Path $repoRoot (Join-Path "docs" $memoryIndexName)
    $memoryIndex = Get-Content -LiteralPath $memoryIndexPath -Raw -Encoding UTF8

    # 只校验记忆索引的长期路径引用：反引号包裹、且以 docs/、.agents/、.codex/、scripts/ 开头。
    # 裸文件名、章节名、通配符和源码 glob 故意跳过，避免把说明文字误判为路径。
    $backtick = [string][char]96
    $referencePattern = $backtick + "((?:docs[\\/]|\.agents[\\/]|\.codex[\\/]|scripts[\\/])[^" + $backtick + "]+)" + $backtick
    $matches = [regex]::Matches($memoryIndex, $referencePattern)
    $missing = @()
    $seen = @{}

    foreach ($match in $matches) {
        $reference = $match.Groups[1].Value.Trim()
        if ($reference.Contains("*")) {
            continue
        }
        $normalized = $reference.TrimEnd([char[]]@("/", "\")).Replace("/", "\")
        if ($seen.ContainsKey($normalized)) {
            continue
        }
        $seen[$normalized] = $true

        $candidate = Join-Path $repoRoot $normalized
        if (-not (Test-Path -LiteralPath $candidate)) {
            $missing += $reference
        }
    }

    if ($missing.Count -gt 0) {
        throw ("docs memory index references missing path(s): " + ($missing -join ", "))
    }
}

function Assert-AutoUpdateStopReadbackCoverage {
    $sourcePath = Join-Path $repoRoot "src\MyTools\Services\WindowsSecurityService.cs"
    $source = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8

    $requiredServices = @(
        "wuauserv",
        "UsoSvc",
        "BITS",
        "DoSvc",
        "WaaSMedicSvc"
    )
    foreach ($service in $requiredServices) {
        $count = [regex]::Matches($source, [regex]::Escape($service)).Count
        if ($count -lt 2) {
            throw "Windows Update stop service $service must appear in both stop handling and readback coverage."
        }
    }

    if (-not $source.Contains("AreExistingServicesDisabled(AutoUpdateBlockingServices)")) {
        throw "Auto-update stop success must read back disabled core service state."
    }
    if (-not $source.Contains("AreExistingScheduledTasksDisabled(AutoUpdateScheduledTasks)")) {
        throw "Auto-update stop success must read back disabled scheduled task state."
    }
    if (-not $source.Contains("AreExistingScheduledTasksEnabled(AutoUpdateScheduledTasks)")) {
        throw "Auto-update restore success must read back enabled scheduled task state."
    }

    $requiredTasks = @(
        "\Microsoft\Windows\WindowsUpdate\Scheduled Start",
        "\Microsoft\Windows\WindowsUpdate\sih",
        "\Microsoft\Windows\WindowsUpdate\sihboot",
        "\Microsoft\Windows\UpdateOrchestrator\Schedule Scan",
        "\Microsoft\Windows\UpdateOrchestrator\Schedule Scan Static Task",
        "\Microsoft\Windows\UpdateOrchestrator\USO_UxBroker",
        "\Microsoft\Windows\UpdateOrchestrator\Maintenance Install",
        "\Microsoft\Windows\UpdateOrchestrator\Reboot",
        "\Microsoft\Windows\UpdateOrchestrator\Reboot_AC",
        "\Microsoft\Windows\UpdateOrchestrator\Reboot_Battery",
        "\Microsoft\Windows\UpdateOrchestrator\Refresh Settings"
    )
    foreach ($task in $requiredTasks) {
        $count = [regex]::Matches($source, [regex]::Escape($task)).Count
        if ($count -lt 2) {
            throw "Windows Update scheduled task $task must appear in both stop handling and readback coverage."
        }
    }
}

function Assert-JunkCleanupScanSafety {
    function Join-CodePointString {
        param([int[]]$CodePoints)
        return -join ($CodePoints | ForEach-Object { [string][char]$_ })
    }

    $servicePath = Join-Path $repoRoot "src\MyTools\Services\JunkCleanupService.cs"
    $viewModelPath = Join-Path $repoRoot "src\MyTools\ViewModels\MainViewModel.cs"
    $viewPath = Join-Path $repoRoot "src\MyTools\Views\SystemOptimizationView.xaml"

    $service = Get-Content -LiteralPath $servicePath -Raw -Encoding UTF8
    $viewModel = Get-Content -LiteralPath $viewModelPath -Raw -Encoding UTF8
    $view = Get-Content -LiteralPath $viewPath -Raw -Encoding UTF8

    $exportReportText = Join-CodePointString @(0x5BFC, 0x51FA, 0x62A5, 0x544A)
    $preflightReportText = Join-CodePointString @(0x9884, 0x68C0, 0x62A5, 0x544A)
    $beforeCleanupReportText = Join-CodePointString @(0x6E05, 0x7406, 0x524D, 0x62A5, 0x544A)
    foreach ($forbiddenToken in @("ExportJunkCleanupPlanCommand", "BuildJunkCleanupPlanText", $exportReportText, $preflightReportText, $beforeCleanupReportText)) {
        if ($viewModel.Contains($forbiddenToken) -or $view.Contains($forbiddenToken)) {
            throw ("Junk cleanup report export entry must stay removed; found " + $forbiddenToken + ".")
        }
    }

    foreach ($token in @("JunkScanStepCount = 11", "WindowsErrorReports", "CrashDumps", "SafetyBoundary", "Evidence", "BuildBrowserCacheRoots", "RunAdminCleanupScriptAsync(pair.Key, pair.Value", "BuildPowerShellStringArray", "`$targets = ", "Invoke-SafeFileCleanup -Root `$target -Targets `$targets", "foreach (`$targetPath in `$Targets)", "Test-PathInsideRoot", "Test-ReparsePoint", "Invoke-SafeFileCleanup", "FileAttributes.ReparsePoint")) {
        if (-not $service.Contains($token)) {
            throw ("Junk cleanup scan safety token missing from service: " + $token)
        }
    }

    foreach ($bindingToken in @("SafetyBoundary", "Evidence", "ScrollViewer.HorizontalScrollBarVisibility")) {
        if (-not $view.Contains($bindingToken)) {
            throw ("Junk cleanup scan grid missing binding token: " + $bindingToken)
        }
    }

    foreach ($forbiddenToken in @("Remove-Item -LiteralPath `$_.FullName -Recurse", "Get-ChildItem -LiteralPath `$target -Force -Recurse")) {
        if ($service.Contains($forbiddenToken)) {
            throw ("Junk cleanup elevated script contains unsafe old recursive token: " + $forbiddenToken)
        }
    }
}

function Assert-ScreenshotSettingsPersistence {
    $settingsPath = Join-Path $repoRoot "src\MyTools\Services\AppSettingsService.cs"
    $viewModelPath = Join-Path $repoRoot "src\MyTools\ViewModels\MainViewModel.cs"
    $settings = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8
    $viewModel = Get-Content -LiteralPath $viewModelPath -Raw -Encoding UTF8

    foreach ($token in @("IsGifRecordingMode", "RecordingFrameRateOption", "RecordingQualityOption")) {
        if (-not $settings.Contains($token)) {
            throw "AppSettings is missing screenshot recording setting: $token"
        }
    }

    foreach ($token in @(
        "_suppressRecordingBehaviorAutoSave",
        "PersistRecordingBehaviorAsync",
        "IsGifRecordingMode = settings.IsGifRecordingMode",
        "RecordingFrameRateOption = settings.RecordingFrameRateOption",
        "RecordingQualityOption = settings.RecordingQualityOption",
        "settings.IsGifRecordingMode = IsGifRecordingMode",
        "settings.RecordingFrameRateOption = RecordingFrameRateOption",
        "settings.RecordingQualityOption = RecordingQualityOption",
        "settings.AudioRecordingFormat = AudioRecordingFormat.FromExtension(AudioRecordingFormatOption).Extension"
    )) {
        if (-not $viewModel.Contains($token)) {
            throw "Screenshot settings persistence path missing token: $token"
        }
    }
}

function Assert-CodexRotationAndOllamaImportRemoved {
    $deletedFiles = @(
        "src\MyTools\Services\CodexOllamaProfileService.cs",
        "src\MyTools\Services\CodexRotationService.cs",
        "src\MyTools\Services\CodexHotTokenService.cs"
    )
    foreach ($relativePath in $deletedFiles) {
        $path = Join-Path $repoRoot $relativePath
        if (Test-Path -LiteralPath $path) {
            throw "Codex rotation/Ollama import removal guard found deleted file: $relativePath"
        }
    }

    function Join-CodePointString {
        param([int[]]$CodePoints)
        return -join ($CodePoints | ForEach-Object { [string][char]$_ })
    }

    $rotateNowText = Join-CodePointString @(0x7ACB, 0x5373, 0x8F6E, 0x6362)
    $selectAllRotateText = Join-CodePointString @(0x5168, 0x9009, 0x8F6E, 0x6362)
    $invertRotateText = Join-CodePointString @(0x53CD, 0x9009, 0x8F6E, 0x6362)
    $notifyAfterRotateText = Join-CodePointString @(0x8F6E, 0x6362, 0x540E, 0x901A, 0x77E5)
    $noRotateTargetText = Join-CodePointString @(0x65E0, 0x53EF, 0x7528, 0x8F6E, 0x6362, 0x76EE, 0x6807)
    $joinRotateText = Join-CodePointString @(0x52A0, 0x5165, 0x8F6E, 0x6362)
    $priorityText = (Join-CodePointString @(0x4F18, 0x5148, 0x7EA7)) + ":"
    $importOllamaText = (Join-CodePointString @(0x5BFC, 0x5165)) + " Ollama " + (Join-CodePointString @(0x6A21, 0x578B))
    $oldExpiredRotateText = Join-CodePointString @(0x5DF2, 0x8FC7, 0x671F, 0xFF0C, 0x8BF7, 0x5237, 0x65B0, 0x540E, 0x518D, 0x8F6E, 0x6362)

    $checks = @{
        "src\MyTools\Views\MainModulesView.xaml" = @(
            $rotateNowText,
            $selectAllRotateText,
            $invertRotateText,
            $notifyAfterRotateText,
            $noRotateTargetText,
            $joinRotateText,
            $priorityText,
            $importOllamaText,
            "ImportCodexOllamaProfilesCommand",
            "RotateToNextCodexProfileCommand",
            "CodexNextSwitchPreview"
        )
        "src\MyTools\ViewModels\MainViewModel.cs" = @(
            "ImportCodexOllamaProfilesCommand",
            "CodexOllamaProfileService",
            "CodexRotationService",
            "CodexHotTokenService",
            "RotateToNextCodexProfileCommand",
            "SelectAllCodexRotationCommand",
            "InvertCodexRotationCommand",
            "CodexRotationSettings",
            "CodexNextSwitchPreview",
            "IsCodexRotationAvailable",
            "UpdateCodexRotationState",
            $noRotateTargetText,
            $oldExpiredRotateText
        )
    }
    foreach ($entry in $checks.GetEnumerator()) {
        $path = Join-Path $repoRoot $entry.Key
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Codex removal guard source missing: $($entry.Key)"
        }

        $source = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        foreach ($token in $entry.Value) {
            if ($source.Contains($token)) {
                throw "Codex rotation/Ollama import UI must stay removed; found token '$token' in $($entry.Key)."
            }
        }
    }
}

function Assert-NativeShellScaffold {
    $nativeRoot = Join-Path $repoRoot "src\MyTools.Native"
    if (-not (Test-Path -LiteralPath $nativeRoot)) {
        throw "Native rewrite scaffold missing: src\MyTools.Native"
    }

    $requiredFiles = @(
        "CMakeLists.txt",
        "README.md",
        "app\main.cpp",
        "app\app_context.h",
        "app\app_context.cpp",
        "app\single_instance.h",
        "app\single_instance.cpp",
        "app\crash_handler.h",
        "app\crash_handler.cpp",
        "app\tray_host.h",
        "app\tray_host.cpp",
        "ccore\include\mt_result.h",
        "ccore\json\mt_json_schema.h",
        "ccore\json\mt_json_schema.c",
        "modules\module_registry.h",
        "modules\module_registry.cpp",
        "modules\codex_profiles\codex_profile_module.h",
        "modules\codex_profiles\codex_profile_module.cpp",
        "services\config_store.h",
        "services\config_store.cpp",
        "services\codex_profile_backup_service.h",
        "services\codex_profile_backup_service.cpp",
        "services\codex_profile_box_service.h",
        "services\codex_profile_box_service.cpp",
        "services\codex_profile_diff_service.h",
        "services\codex_profile_diff_service.cpp",
        "services\codex_profile_edit_service.h",
        "services\codex_profile_edit_service.cpp",
        "services\codex_profile_export_service.h",
        "services\codex_profile_export_service.cpp",
        "services\codex_profile_import_service.h",
        "services\codex_profile_import_service.cpp",
        "services\codex_profile_switch_service.h",
        "services\codex_profile_switch_service.cpp",
        "services\file_scanner.h",
        "services\file_scanner.cpp",
        "services\file_system.h",
        "services\file_system.cpp",
        "services\hotkey_service.h",
        "services\hotkey_service.cpp",
        "services\logger.h",
        "services\logger.cpp",
        "services\network_service.h",
        "services\network_service.cpp",
        "services\process_runner.h",
        "services\process_runner.cpp",
        "services\secret_store_dpapi.h",
        "services\secret_store_dpapi.cpp",
        "services\task_runner.h",
        "services\task_runner.cpp",
        "ui\main_window.h",
        "ui\main_window.cpp",
        "ui\modal_dialogs.h",
        "ui\modal_dialogs.cpp",
        "ui\renderer_d2d.h",
        "ui\renderer_d2d.cpp",
        "resources\app.rc.in",
        "resources\resource.h"
    )
    foreach ($relativePath in $requiredFiles) {
        $path = Join-Path $nativeRoot $relativePath
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Native shell scaffold missing required file: src\MyTools.Native\$relativePath"
        }
    }

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

    $requiredTokens = @(
        @{ Name = "CMake"; Source = $cmake; Tokens = @("LANGUAGES C CXX", "cxx_std_20", "d2d1", "dwrite", "bcrypt", "comdlg32", "crypt32", "ole32", "WIN32", "configure_file", "ws2_32", "ccore/json/mt_json_schema.c", "modules/codex_profiles/codex_profile_module.cpp", "modules/module_registry.cpp", "services/codex_profile_backup_service.cpp", "services/codex_profile_box_service.cpp", "services/codex_profile_diff_service.cpp", "services/codex_profile_edit_service.cpp", "services/codex_profile_export_service.cpp", "services/codex_profile_import_service.cpp", "services/codex_profile_switch_service.cpp", "ui/modal_dialogs.cpp") },
        @{ Name = "main"; Source = $main; Tokens = @("wWinMain", "SetProcessDpiAwarenessContext", "DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2") },
        @{ Name = "renderer"; Source = $renderer; Tokens = @("D2D1CreateFactory", "DWriteCreateFactory") },
        @{ Name = "dpapi"; Source = $secret; Tokens = @("CryptProtectData", "CryptUnprotectData", "CryptBinaryToStringA", "CryptStringToBinaryA", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8") },
        @{ Name = "tray"; Source = $tray; Tokens = @("Shell_NotifyIconW", "NIM_ADD", "NIM_DELETE") },
        @{ Name = "config"; Source = $config; Tokens = @("mt_json_has_schema_version", "schema_version", "WriteUtf8FileAtomic") },
        @{ Name = "codex-backup"; Source = $codexBackup; Tokens = @("CreateCurrentFolderBackup", "RestoreLatestBackup", "native-codex-current-folder-before-switch", "ConfigTomlBase64", "AuthJsonBase64", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8", "CryptStringToBinaryA", "FindFirstFileW", "CompareFileTime", "WriteUtf8FileAtomic", "WriteFileBytesAtomic", "SecureZeroMemory", ".bak.dpapi") },
        @{ Name = "codex-box"; Source = $codexBox; Tokens = @("CodexProfileBoxService", "ExportBox", "ImportBox", "CodexProfileBoxConflictPolicy", "CDXB", "portable-codex-profiles-v2", "PBKDF2", "BCryptDeriveKeyPBKDF2", "BCRYPT_CHAIN_MODE_CBC", "BCRYPT_BLOCK_PADDING", "BCryptGenRandom", "HmacSha256", "ConstantTimeEquals", "200000", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlBase64", "AuthJsonBase64", "WriteFileBytesAtomic", "WriteUtf8FileAtomic", "SecureZeroMemory") },
        @{ Name = "codex-diff"; Source = $codexDiff; Tokens = @("BuildProfileDiffSummary", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "UnprotectBase64ToUtf8", "ReadFileBytes", "FileSystem::Exists", "current_available", "BCryptOpenAlgorithmProvider", "BCRYPT_SHA256_ALGORITHM", "profile_sha256_hex", "current_sha256_hex", "profile_line_count", "current_line_count", "SecureZeroMemory") },
        @{ Name = "codex-edit"; Source = $codexEdit; Tokens = @("UpdateProfileMetadata", "update_display_name", "update_note", "update_remark", "update_tags", "DisplayName", "Name", "Note", "Remark", "Tags", "UnprotectBase64ToUtf8", "ProtectUtf8ToBase64", "WriteUtf8FileAtomic", "CompareStringOrdinal", "SetJsonStringProperty", "TrimWide", "ContainsControlCharacter", "120 characters", "control characters", "0x9F", "SecureZeroMemory") },
        @{ Name = "codex-export"; Source = $codexExport; Tokens = @("ExportProfileByDisplayName", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "UnprotectBase64ToUtf8", "WriteUtf8FileAtomic", "config.toml", "auth.json", "SecureZeroMemory") },
        @{ Name = "codex-import"; Source = $codexImport; Tokens = @("ImportCurrentFolderProfile", "ProtectUtf8ToBase64", "UnprotectBase64ToUtf8", "ReadFileBytes", "WriteUtf8FileAtomic", "schemaVersion", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "LastImportedAt", "SecureZeroMemory", "config.toml", "auth.json") },
        @{ Name = "codex-switch"; Source = $codexSwitch; Tokens = @("ApplyProfileByDisplayName", "ProtectedConfigTomlBase64", "ProtectedAuthJsonBase64", "ConfigTomlContentProtected", "AuthJsonContentProtected", "CreateCurrentFolderBackup", "UnprotectBase64ToUtf8", "WriteUtf8FileAtomic", "ActiveDisplayName", "SwitchedAtUtc", "config.toml", "auth.json") },
        @{ Name = "file-system"; Source = $fileSystem; Tokens = @("MoveFileExW", "MOVEFILE_REPLACE_EXISTING", "FlushFileBuffers", "ReadFileBytes", "WriteFileBytesAtomic") },
        @{ Name = "file-scanner"; Source = $scanner; Tokens = @("FindFirstFileW", "FILE_ATTRIBUTE_REPARSE_POINT", "IsCancellationRequested") },
        @{ Name = "hotkey"; Source = $hotkey; Tokens = @("RegisterHotKey", "UnregisterHotKey") },
        @{ Name = "network"; Source = $network; Tokens = @("WSAStartup", "GetAddrInfoW", "select") },
        @{ Name = "process"; Source = $processRunner; Tokens = @("CreateProcessW", "TerminateProcess") },
        @{ Name = "task-runner"; Source = $taskRunner; Tokens = @("std::condition_variable", "std::thread", "WorkerLoop") },
        @{ Name = "c-json"; Source = $jsonSchema; Tokens = @("mt_json_has_schema_version", '\"schema_version\"') },
        @{ Name = "main-window-nav"; Source = $mainWindow; Tokens = @("ID_NAV_CODEX_PROFILES", "ID_CODEX_REFRESH", "ID_CODEX_DIFF_FIRST", "ID_CODEX_BACKUP_CURRENT", "ID_CODEX_APPLY_FIRST", "ID_CODEX_IMPORT_CURRENT", "ID_CODEX_RESTORE_BACKUP", "ID_CODEX_EXPORT_BOX", "ID_CODEX_IMPORT_BOX", "ID_CODEX_EXPORT_FIRST_FILES", "ID_CODEX_RENAME_FIRST", "HandleCodexProfileAction", "PrepareCodexProfileActionOptions", "PromptCodexProfileTarget", "ChooseCodexProfileDisplayName", "PromptCodexBoxPassword", "PromptCodexBoxConflictPolicy", "PromptCodexProfileDisplayName", "PromptText", "PickCodexBoxSavePath", "PickCodexBoxOpenPath", "PickFolderPath", "TrimWide", "ContainsControlCharacter", "SecureClearWideString", "ConfirmCodexProfileWrite", "MB_YESNO", "MB_DEFBUTTON2", "SwitchModule", "Codex Profiles", "Diff profile", "Apply profile...", "Export profile files...", "Rename profile...", "Edit profile note...", "CodexProfileUiAction::ExportFirstProfileFiles", "CodexProfileUiAction::RenameFirstProfile", "options->target_display_name", "options->output_directory", "options->new_display_name", "options->box_conflict_policy", "120 characters", "control characters", "0x9F", "FRP Tunnel (pending)") },
        @{ Name = "modal-dialogs"; Source = $modalDialogs; Tokens = @("PromptPassword", "PromptText", "PickFolderPath", "ChooseCodexBoxConflictPolicy", "ChooseCodexProfileDisplayName", "SecureClearWideString", "GetSaveFileNameW", "GetOpenFileNameW", "SHBrowseForFolderW", "SHGetPathFromIDListW", "CoInitializeEx", "COINIT_APARTMENTTHREADED", "CoTaskMemFree", "CoUninitialize", "OFN_OVERWRITEPROMPT", "OFN_FILEMUSTEXIST", "ES_PASSWORD", "EM_SETPASSWORDCHAR", "EnableWindow", "Codex profile package (*.codexbox)", "Choose Codex profile", "LISTBOX", "LB_ADDSTRING", "LB_GETCURSEL") },
        @{ Name = "renderer-contract"; Source = $rendererHeader; Tokens = @("ModuleInfo") },
        @{ Name = "codex-profile-scaffold"; Source = $modules; Tokens = @("LOCALAPPDATA", "USERPROFILE", "profiles.json", "active.json", "Backups", "config.toml", "auth.json", "UnprotectBase64ToUtf8", "ExtractJsonStringValues", "ExtractJsonObjectArray", "ReadHexCodeUnit", "AppendUtf8CodePoint", "ActiveDisplayName", "active profile", "Switch backup directory", "explicit menu actions", ".codexbox", "DisplayName", "Status", "LastImportedAt", "first_profile_display_name", "profile_display_names", "CodexProfileUiAction", "RunUiAction", "DiffFirstProfile", "BackupCurrentFolder", "ApplyFirstProfile", "ImportCurrentFolder", "RestoreLatestBackup", "ExportBox", "ImportBox", "ExportFirstProfileFiles", "RenameFirstProfile", "std::wstring target_display_name", "std::wstring output_directory", "std::wstring new_display_name", "options.target_display_name", "options.output_directory", "options.new_display_name", "options.box_conflict_policy", "SelectedProfileTarget", "request.output_directory", "request.update_display_name", "request.new_display_name", "request.active_json_path", "request.active_display_name", "request.conflict_policy = options.box_conflict_policy", "SameDirectoryPath(options.output_directory, probe.paths.codex_home)", "Active profile marker updated", "CodexProfileExportService", "CodexProfileEditService", "ExportProfileByDisplayName", "UpdateProfileMetadata", "without exposing embedded auth/config secrets") }
    )
    foreach ($group in $requiredTokens) {
        foreach ($token in $group.Tokens) {
            if (-not $group.Source.Contains($token)) {
                throw "Native shell $($group.Name) missing required token: $token"
            }
        }
    }
    Assert-NativeCodexMetadataEditGuards -MainWindow $mainWindow -Modules $modules -CodexEdit $codexEdit
    Assert-NativeCodexBoxConflictGuards -MainWindow $mainWindow -ModalDialogs $modalDialogs -Modules $modules -CodexBox $codexBox

    $combined = $cmake + $main + $renderer + $secret + $tray + $config + $codexBackup + $codexBox + $codexDiff + $codexEdit + $codexExport + $codexImport + $codexSwitch + $fileSystem + $scanner + $hotkey + $network + $processRunner + $taskRunner + $jsonSchema + $mainWindow + $modalDialogs + $rendererHeader + $modules
    foreach ($blocked in @("Qt", "Electron", "WebView2", "WindowsAppSDK", "Microsoft.NET")) {
        if ($combined.Contains($blocked)) {
            throw "Native shell phase 1/2 must not introduce blocked runtime token: $blocked"
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

    $nativeEval = Join-Path $repoRoot "scripts\native-eval.ps1"
    if (-not (Test-Path -LiteralPath $nativeEval)) {
        throw "Native validation entry missing: scripts\native-eval.ps1"
    }
}

if ($runAll -or $Quick) {
    Invoke-Step "loop memory docs" $repoRoot {
        Assert-LoopMemoryDocs
    }

    Invoke-Step "startup responsiveness guards" $repoRoot {
        Assert-StartupResponsivenessGuards
    }

    Invoke-Step "sql export removal guards" $repoRoot {
        Assert-SqlExportRemoved
    }

    Invoke-Step "codex rotation and ollama import removal guards" $repoRoot {
        Assert-CodexRotationAndOllamaImportRemoved
    }

    Invoke-Step "native shell scaffold" $repoRoot {
        Assert-NativeShellScaffold
    }

    Invoke-Step "auto update stop readback coverage" $repoRoot {
        Assert-AutoUpdateStopReadbackCoverage
    }

    Invoke-Step "auto optimize plan safety" $repoRoot {
        $servicePath = Join-Path $repoRoot "src\MyTools\Services\SystemOptimizationService.cs"
        $viewModelPath = Join-Path $repoRoot "src\MyTools\ViewModels\MainViewModel.cs"
        $viewPath = Join-Path $repoRoot "src\MyTools\Views\SystemOptimizationView.xaml"

        $service = Get-Content -LiteralPath $servicePath -Raw -Encoding UTF8
        $viewModel = Get-Content -LiteralPath $viewModelPath -Raw -Encoding UTF8
        $view = Get-Content -LiteralPath $viewPath -Raw -Encoding UTF8

        if (-not $service.Contains("const int totalSteps = 14")) {
            throw "Auto optimize scan must keep the documented 14-step plan."
        }

        foreach ($stepKey in @("StepKeyDeliveryOptimizationCache", "StepKeyWindowsErrorReports", "StepKeyCrashDumps", "StepKeyComponentStoreCleanup")) {
            if (-not $service.Contains($stepKey)) {
                throw "Auto optimize scan is missing deeper step key: $stepKey"
            }
        }

        $doubleAmpersand = ([string][char]38) + ([string][char]38)
        $defaultSelectionToken = "IsSelected = canOptimize " + $doubleAmpersand + " defaultSelected"
        foreach ($token in @("RiskLevel", "SafetyBoundary", "Evidence", $defaultSelectionToken, "FileAttributes.ReparsePoint", "IsPathInsideRoot(file, safeRoot)", "ElevatedScriptTimeoutMs", "process.ExitCode != 0", "BuildSafeDeleteScriptPrelude", "Test-PathInsideRoot", "Test-ReparsePoint")) {
            if (-not $service.Contains($token)) {
                throw "Auto optimize safety token missing from service: $token"
            }
        }

        foreach ($forbiddenToken in @("-Recurse -File", "-Recurse -Directory", "Remove-Item -LiteralPath `$_.FullName -Recurse")) {
            if ($service.Contains($forbiddenToken)) {
                throw "Auto optimize elevated cleanup script contains unsafe or PS2-incompatible token: $forbiddenToken"
            }
        }

        $quote = [char]34
        $highRiskDefaultPattern = "RiskHigh,\s*$quote[^$quote]*$quote,\s*[^,]+,\s*true"
        $highRiskDefaultSelected = [regex]::Matches($service, $highRiskDefaultPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($highRiskDefaultSelected.Count -gt 0) {
            throw "High-risk auto optimize items must not be selected by default."
        }

        $selectedFilterToken = "Where(x => x.CanOptimize " + $doubleAmpersand + " x.IsSelected)"
        if (-not $viewModel.Contains($selectedFilterToken)) {
            throw "Auto optimize selected-plan ViewModel guard missing: selected item filter"
        }
        if (-not $viewModel.Contains("BuildAutoOptimizeConfirmMessage")) {
            throw "Auto optimize selected-plan ViewModel guard missing: confirm message builder"
        }
        if (-not $viewModel.Contains("DetachAutoOptimizePlanItems")) {
            throw "Auto optimize selected-plan ViewModel guard missing: plan item detach"
        }

        foreach ($bindingToken in @("IsSelected", "RiskLevel", "SafetyBoundary", "Evidence")) {
            if (-not $view.Contains($bindingToken)) {
                throw "Auto optimize scan grid missing binding token: $bindingToken"
            }
        }
    }

    Invoke-Step "junk cleanup scan safety" $repoRoot {
        Assert-JunkCleanupScanSafety
    }

    Invoke-Step "screenshot settings persistence" $repoRoot {
        Assert-ScreenshotSettingsPersistence
    }
}

$dotnetExe = Resolve-Dotnet

if ($runAll -or $Quick -or $Build) {
    Invoke-Step "mytools release build" $repoRoot {
        & $dotnetExe build src\MyTools\MyTools.csproj -c Release
    }
}

if ($runAll -or $Quick) {
    Invoke-Step "codex local relay cursor compatibility" $repoRoot {
        $newtonsoftPath = Join-Path $env:USERPROFILE ".nuget\packages\newtonsoft.json\13.0.3\lib\net45\Newtonsoft.Json.dll"
        if (Test-Path -LiteralPath $newtonsoftPath) {
            [Reflection.Assembly]::LoadFrom($newtonsoftPath) | Out-Null
        }

        $appAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "src\MyTools\bin\Release\net48\MyTools.exe"))
        $compatType = $appAssembly.GetType("MyTools.Services.CodexLocalRelayCompatibility", $true)

        $modelsMethod = $compatType.GetMethod("BuildModelsResponseJson", [Reflection.BindingFlags]"Public,Static")
        $modelsJson = [string]$modelsMethod.Invoke($null, @("gpt-5.5"))
        $jsonQuote = [string][char]34
        $modelsObjectNeedle = $jsonQuote + "object" + $jsonQuote + ":" + $jsonQuote + "list" + $jsonQuote
        $modelsIdNeedle = $jsonQuote + "id" + $jsonQuote + ":" + $jsonQuote + "gpt-5.5" + $jsonQuote
        if (-not $modelsJson.Contains($modelsObjectNeedle) -or -not $modelsJson.Contains($modelsIdNeedle)) {
            throw "Cursor-compatible models fallback did not return an OpenAI-style model list."
        }

        $contentTypeMethod = $compatType.GetMethod("IsOpenAiCompatibleContentType", [Reflection.BindingFlags]"Public,Static")
        if ([bool]$contentTypeMethod.Invoke($null, @("text/html; charset=utf-8"))) {
            throw "Cursor compatibility should reject HTML success responses from upstream."
        }
        if ([bool]$contentTypeMethod.Invoke($null, @(""))) {
            throw "Cursor compatibility should reject empty upstream content-type."
        }
        if (-not [bool]$contentTypeMethod.Invoke($null, @("text/event-stream; charset=utf-8"))) {
            throw "Cursor compatibility should accept SSE upstream responses."
        }

        $relayType = $appAssembly.GetType("MyTools.Services.CodexLocalRelayService", $true)
        $responsesUrisMethod = $relayType.GetMethod("TryBuildResponsesCompatibilityUris", [Reflection.BindingFlags]"NonPublic,Static")
        [object[]]$uriArguments = @([Uri]"https://c-api.cc", "/v1", $null, $null, $null, $null)
        $urisBuilt = [bool]$responsesUrisMethod.Invoke($null, $uriArguments)
        if (-not $urisBuilt) {
            throw "Relay should build Responses compatibility URIs for root upstream base URL."
        }
        $primaryResponsesUri = [Uri]$uriArguments[3]
        $fallbackResponsesUri = [Uri]$uriArguments[5]
        if ($primaryResponsesUri.AbsolutePath -ne "/responses") {
            throw "Primary Responses compatibility URI should preserve existing root-base mapping; got $($primaryResponsesUri.AbsolutePath)."
        }
        if ($fallbackResponsesUri.AbsolutePath -ne "/v1/responses") {
            throw "Fallback Responses compatibility URI should add /v1 before /responses; got $($fallbackResponsesUri.AbsolutePath)."
        }

        $chatJson = '{"model":"gpt-5.5","messages":[{"role":"user","content":"ping"}],"tools":[{"type":"function","function":{"name":"lookup","description":"lookup data","parameters":{"type":"object","properties":{"q":{"type":"string"}}}}}],"tool_choice":"auto","max_tokens":7,"stream":true}'
        $chatBytes = [Text.Encoding]::UTF8.GetBytes($chatJson)
        $buildResponsesMethod = $compatType.GetMethod("TryBuildResponsesRequestBody", [Reflection.BindingFlags]"Public,Static")
        [object[]]$buildArguments = @($chatBytes, "fallback-model", $null, $false)
        $built = [bool]$buildResponsesMethod.Invoke($null, $buildArguments)
        if (-not $built) {
            throw "Cursor chat/completions request should convert to Responses API request body."
        }

        $responsesBody = [Text.Encoding]::UTF8.GetString([byte[]]$buildArguments[2])
        if (-not [bool]$buildArguments[3]) {
            throw "Cursor compatibility should preserve client streaming preference."
        }
        if (-not $responsesBody.Contains($jsonQuote + "input" + $jsonQuote) -or -not $responsesBody.Contains($jsonQuote + "tools" + $jsonQuote) -or -not $responsesBody.Contains($jsonQuote + "max_output_tokens" + $jsonQuote + ":7")) {
            throw "Converted Responses API request body lost input, tools, or token limit fields."
        }

        $responsesJson = '{"id":"resp_123","model":"gpt-5.5","output":[{"type":"message","content":[{"type":"output_text","text":"pong"}]}],"usage":{"input_tokens":1,"output_tokens":1,"total_tokens":2}}'
        $chatCompletionMethod = $compatType.GetMethod("TryBuildChatCompletionJson", [Reflection.BindingFlags]"Public,Static")
        [object[]]$completionArguments = @($responsesJson, $chatBytes, $null)
        $converted = [bool]$chatCompletionMethod.Invoke($null, $completionArguments)
        if (-not $converted) {
            throw "Responses API output should convert to chat/completions JSON."
        }

        $completionJson = [string]$completionArguments[2]
        if (-not $completionJson.Contains($jsonQuote + "object" + $jsonQuote + ":" + $jsonQuote + "chat.completion" + $jsonQuote) -or -not $completionJson.Contains($jsonQuote + "content" + $jsonQuote + ":" + $jsonQuote + "pong" + $jsonQuote) -or -not $completionJson.Contains($jsonQuote + "total_tokens" + $jsonQuote + ":2")) {
            throw "Converted chat/completions JSON lost object type, content, or usage fields."
        }

        $sseMethod = $compatType.GetMethod("TryBuildChatCompletionSsePayload", [Reflection.BindingFlags]"Public,Static")
        [object[]]$sseArguments = @($responsesJson, $chatBytes, $null)
        $sseConverted = [bool]$sseMethod.Invoke($null, $sseArguments)
        if (-not $sseConverted) {
            throw "Responses API output should convert to chat/completions SSE."
        }

        $ssePayload = [string]$sseArguments[2]
        if (-not $ssePayload.Contains($jsonQuote + "object" + $jsonQuote + ":" + $jsonQuote + "chat.completion.chunk" + $jsonQuote) -or -not $ssePayload.Contains($jsonQuote + "content" + $jsonQuote + ":" + $jsonQuote + "pong" + $jsonQuote) -or -not $ssePayload.Contains("data: [DONE]")) {
            throw "Converted chat/completions SSE lost chunk content or DONE marker."
        }
    }

    Invoke-Step "schedule excel big shift export" $repoRoot {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $outputDir = Join-Path ([IO.Path]::GetTempPath()) "MyToolsCodexEval\schedule-export-check"
        New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
        $xlsxPath = Join-Path $outputDir "schedule-big-shift.xlsx"
        if (Test-Path -LiteralPath $xlsxPath) {
            Remove-Item -LiteralPath $xlsxPath -Force
        }

        $newtonsoftPath = Join-Path $env:USERPROFILE ".nuget\packages\newtonsoft.json\13.0.3\lib\net45\Newtonsoft.Json.dll"
        if (Test-Path -LiteralPath $newtonsoftPath) {
            [Reflection.Assembly]::LoadFrom($newtonsoftPath) | Out-Null
        }

        $appAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "src\MyTools\bin\Release\net48\MyTools.exe"))
        $scheduleType = $appAssembly.GetType("MyTools.Services.ScheduleVersion", $true)
        $employeeType = $appAssembly.GetType("MyTools.Services.EmployeeRow", $true)
        $cellType = $appAssembly.GetType("MyTools.Services.ShiftCell", $true)
        $exporterType = $appAssembly.GetType("MyTools.Services.ScheduleExcelExporter", $true)
        $schedulePageType = $appAssembly.GetType("MyTools.Views.SchedulePage", $true)
        $shiftCodesType = $appAssembly.GetType("MyTools.Services.ShiftCodes", $true)
        $bigShiftCode = [string][char]0x5927
        $legacyNightCode = [string][char]0x591C
        $smallShiftCode = [string][char]0x5C0F
        $publicShiftCode = [string][char]0x516C
        $restShiftCode = [string][char]0x4F11
        $maternityShiftCode = [string][char]0x4EA7
        $weekendRestHeader = ([string][char]0x672B) + ([string][char]0x4F11)
        $availableWorkHeader = ([string][char]0x53EF) + ([string][char]0x4E0A)
        $maternityPickerLabel = $maternityShiftCode + ([string][char]0x5047)
        $halfShiftCode = [string][char]0x5348

        $schedule = [Activator]::CreateInstance($scheduleType)
        $schedule.Year = 2026
        $schedule.Month = 8
        $schedule.VersionName = "export-check"
        for ($day = 0; $day -lt $schedule.DayCount; $day++) {
            $schedule.DailyRestQuotas.Add(1.0)
        }

        $employee = [Activator]::CreateInstance($employeeType)
        $employee.Name = "export-check"
        for ($day = 0; $day -lt $schedule.DayCount; $day++) {
            $cell = [Activator]::CreateInstance($cellType)
            if ($day -eq 0) {
                $cell.Code = $smallShiftCode
            } elseif ($day -eq 1) {
                $cell.Code = $legacyNightCode
            } elseif ($day -eq 2) {
                $cell.Code = $publicShiftCode
            } elseif ($day -eq 3) {
                $cell.Code = $halfShiftCode
            } elseif ($day -eq 4) {
                $cell.Code = $maternityShiftCode
            } elseif ($day -eq 7) {
                $cell.Code = $restShiftCode
            } elseif ($day -eq 8) {
                $cell.Code = $halfShiftCode
            } else {
                $cell.Code = ""
            }
            $employee.Cells.Add($cell)
        }
        $schedule.Employees.Add($employee)

        $exportMethod = $exporterType.GetMethod("ExportAsync", [Reflection.BindingFlags]"Public,Static")
        [object[]]$exportArguments = @($schedule, [string]$xlsxPath)
        $task = $exportMethod.Invoke($null, $exportArguments)
        $null = $task.GetAwaiter().GetResult()

        $zip = [IO.Compression.ZipFile]::OpenRead($xlsxPath)
        try {
            $entry = $zip.GetEntry("xl/worksheets/sheet1.xml")
            $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
            try {
                $worksheetXml = $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $zip.Dispose()
        }

        if ($worksheetXml -notmatch [regex]::Escape($bigShiftCode)) {
            throw "Exported worksheet did not contain the big-shift display text."
        }

        if ($worksheetXml -match [regex]::Escape($legacyNightCode)) {
            throw "Big shift was exported as the legacy night text instead of the current big-shift text."
        }

        if ($worksheetXml -notmatch [regex]::Escape($smallShiftCode)) {
            throw "Exported worksheet did not contain the small-shift display text."
        }

        if ($worksheetXml -notmatch [regex]::Escape($maternityShiftCode)) {
            throw "Exported worksheet did not contain the maternity-leave display text."
        }

        $resolveCellBg = $schedulePageType.GetMethod("ResolveCellBg", [Reflection.BindingFlags]"NonPublic,Static")
        $restDaysMethod = $shiftCodesType.GetMethod("RestDays", [Reflection.BindingFlags]"Public,Static")
        $workDaysMethod = $shiftCodesType.GetMethod("WorkDays", [Reflection.BindingFlags]"Public,Static")
        $smallCell = [Activator]::CreateInstance($cellType)
        $smallCell.Code = $smallShiftCode
        $bigCell = [Activator]::CreateInstance($cellType)
        $bigCell.Code = $bigShiftCode
        $publicCell = [Activator]::CreateInstance($cellType)
        $publicCell.Code = $publicShiftCode
        $maternityCell = [Activator]::CreateInstance($cellType)
        $maternityCell.Code = $maternityShiftCode
        $halfCell = [Activator]::CreateInstance($cellType)
        $halfCell.Code = $halfShiftCode
        $smallBrush = $resolveCellBg.Invoke($null, @($smallCell, $true))
        $bigBrush = $resolveCellBg.Invoke($null, @($bigCell, $true))
        $publicBrush = $resolveCellBg.Invoke($null, @($publicCell, $false))
        $maternityBrush = $resolveCellBg.Invoke($null, @($maternityCell, $false))
        $halfBrush = $resolveCellBg.Invoke($null, @($halfCell, $false))
        $maternityRestDays = [double]$restDaysMethod.Invoke($null, @($maternityShiftCode))
        $maternityWorkDays = [double]$workDaysMethod.Invoke($null, @($maternityShiftCode))
        if ($smallBrush.Color.ToString() -ne "#FF475569") {
            throw "Holiday small-shift UI background should be #475569, got $($smallBrush.Color)."
        }
        if ($bigBrush.Color.ToString() -ne "#FF1E293B") {
            throw "Holiday big-shift UI background should be #1E293B, got $($bigBrush.Color)."
        }
        if ($publicBrush.Color.ToString() -ne "#FFFEF3C7") {
            throw "Public-shift UI background should be #FEF3C7, got $($publicBrush.Color)."
        }
        if ($maternityBrush.Color.ToString() -ne "#FFFED7A8") {
            throw "Maternity-leave UI background should be #FED7A8, got $($maternityBrush.Color)."
        }
        if ($halfBrush.Color.ToString() -ne "#FFFEF3C7") {
            throw "Half-shift UI background should remain #FEF3C7, got $($halfBrush.Color)."
        }
        if ($maternityRestDays -ne 1.0 -or $maternityWorkDays -ne 0.0) {
            throw "Maternity leave should count as 1 rest day and 0 work days, got rest=$maternityRestDays work=$maternityWorkDays."
        }

        $pickerLabelsMethod = $schedulePageType.GetMethod("GetShiftPickerVisibleLabels", [Reflection.BindingFlags]"NonPublic,Static")
        $pickerLabels = [string[]]$pickerLabelsMethod.Invoke($null, @())
        if ($pickerLabels -contains $halfShiftCode) {
            throw "Shift picker should hide the half-day option."
        }
        if ($pickerLabels -notcontains $publicShiftCode) {
            throw "Shift picker should still show the public-rest option."
        }
        if ($pickerLabels -notcontains $maternityPickerLabel) {
            throw "Shift picker should show the maternity-leave option."
        }

        $stylesXml = $null
        $zip = [IO.Compression.ZipFile]::OpenRead($xlsxPath)
        try {
            $entry = $zip.GetEntry("xl/styles.xml")
            $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
            try {
                $stylesXml = $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $zip.Dispose()
        }

        [xml]$sheetDoc = $worksheetXml
        [xml]$stylesDoc = $stylesXml
        $ns = New-Object Xml.XmlNamespaceManager($sheetDoc.NameTable)
        $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $styleNs = New-Object Xml.XmlNamespaceManager($stylesDoc.NameTable)
        $styleNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        function Get-CellText {
            param(
                [xml]$Sheet,
                [Xml.XmlNamespaceManager]$SheetNs,
                [string]$CellRef
            )

            $cellXPath = "//x:c[@r=" + [char]39 + $CellRef + [char]39 + "]"
            $cellNode = $Sheet.SelectSingleNode($cellXPath, $SheetNs)
            if ($null -eq $cellNode) {
                throw "Cell $CellRef was not found in worksheet."
            }

            $inlineTextNode = $cellNode.SelectSingleNode("x:is/x:t", $SheetNs)
            if ($null -ne $inlineTextNode) {
                return [string]$inlineTextNode.InnerText
            }

            $valueNode = $cellNode.SelectSingleNode("x:v", $SheetNs)
            if ($null -ne $valueNode) {
                return [string]$valueNode.InnerText
            }

            return ""
        }

        function Get-FillColorForCell {
            param(
                [xml]$Sheet,
                [System.Xml.XmlNamespaceManager]$SheetNs,
                [xml]$Styles,
                [System.Xml.XmlNamespaceManager]$StylesNs,
                [string]$CellRef
            )

            $cellXPath = "//x:c[@r=" + [char]39 + $CellRef + [char]39 + "]"
            $cellNode = $Sheet.SelectSingleNode($cellXPath, $SheetNs)
            if ($null -eq $cellNode) {
                throw "Cell $CellRef was not found in worksheet."
            }

            $styleIndex = [int]$cellNode.s
            $xfNode = $Styles.SelectSingleNode("//x:cellXfs/x:xf[$($styleIndex + 1)]", $StylesNs)
            $fillId = [int]$xfNode.fillId
            $fillNode = $Styles.SelectSingleNode("//x:fills/x:fill[$($fillId + 1)]/x:patternFill/x:fgColor", $StylesNs)
            return $fillNode.rgb
        }

        $weekendRestHeaderCell = Get-CellText $sheetDoc $ns "AJ1"
        if ($weekendRestHeaderCell -ne $weekendRestHeader) {
            throw "Schedule export should write weekend-rest header to AJ1, got $weekendRestHeaderCell."
        }

        $weekendRestValue = Get-CellText $sheetDoc $ns "AJ5"
        if ($weekendRestValue -ne "1.5") {
            throw "Schedule export should count one full and one half weekend rest day in AJ5, got $weekendRestValue."
        }

        $availableWorkHeaderCell = Get-CellText $sheetDoc $ns "AK1"
        if ($availableWorkHeaderCell -ne $availableWorkHeader) {
            throw "Schedule export should write available-work header to AK1, got $availableWorkHeaderCell."
        }

        $availableWorkValue = Get-CellText $sheetDoc $ns "AK5"
        if ($availableWorkValue -ne "26") {
            throw "Schedule export should count unset cells plus current work cells as available work in AK5, got $availableWorkValue."
        }

        $smallFill = Get-FillColorForCell $sheetDoc $ns $stylesDoc $styleNs "B5"
        $bigFill = Get-FillColorForCell $sheetDoc $ns $stylesDoc $styleNs "C5"
        $publicFill = Get-FillColorForCell $sheetDoc $ns $stylesDoc $styleNs "D5"
        $halfFill = Get-FillColorForCell $sheetDoc $ns $stylesDoc $styleNs "E5"
        $maternityFill = Get-FillColorForCell $sheetDoc $ns $stylesDoc $styleNs "F5"
        if ($smallFill -ne "FF475569") {
            throw "Holiday small-shift Excel fill should be FF475569, got $smallFill."
        }
        if ($bigFill -ne "FF1E293B") {
            throw "Holiday big-shift Excel fill should be FF1E293B, got $bigFill."
        }
        if ($publicFill -ne "FFFEF3C7") {
            throw "Public-shift Excel fill should be FFFEF3C7, got $publicFill."
        }
        if ($maternityFill -ne "FFFED7A8") {
            throw "Maternity-leave Excel fill should be FFFED7A8, got $maternityFill."
        }
        if ($halfFill -ne "FFFEF3C7") {
            throw "Half-shift Excel fill should remain FFFEF3C7, got $halfFill."
        }
    }

    Invoke-Step "schedule excel import" $repoRoot {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $outputDir = Join-Path ([IO.Path]::GetTempPath()) "MyToolsCodexEval\schedule-import-check"
        New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
        $xlsxPath = Join-Path $outputDir "排班_2026-08_import-check.xlsx"
        if (Test-Path -LiteralPath $xlsxPath) {
            Remove-Item -LiteralPath $xlsxPath -Force
        }

        $newtonsoftPath = Join-Path $env:USERPROFILE ".nuget\packages\newtonsoft.json\13.0.3\lib\net45\Newtonsoft.Json.dll"
        if (Test-Path -LiteralPath $newtonsoftPath) {
            [Reflection.Assembly]::LoadFrom($newtonsoftPath) | Out-Null
        }

        $appAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "src\MyTools\bin\Release\net48\MyTools.exe"))
        $scheduleType = $appAssembly.GetType("MyTools.Services.ScheduleVersion", $true)
        $employeeType = $appAssembly.GetType("MyTools.Services.EmployeeRow", $true)
        $cellType = $appAssembly.GetType("MyTools.Services.ShiftCell", $true)
        $exporterType = $appAssembly.GetType("MyTools.Services.ScheduleExcelExporter", $true)
        $importerType = $appAssembly.GetType("MyTools.Services.ScheduleExcelImporter", $true)
        $dayShiftCode = [string][char]0x767D
        $bigShiftCode = [string][char]0x5927
        $smallShiftCode = [string][char]0x5C0F
        $publicShiftCode = [string][char]0x516C
        $maternityShiftCode = [string][char]0x4EA7
        $halfShiftCode = [string][char]0x5348

        $schedule = [Activator]::CreateInstance($scheduleType)
        $schedule.Year = 2026
        $schedule.Month = 8
        $schedule.VersionName = "import-check"
        for ($day = 0; $day -lt $schedule.DayCount; $day++) {
            $schedule.DailyRestQuotas.Add(0.0)
        }
        $schedule.DailyRestQuotas[2] = 1.0
        $schedule.DailyRestQuotas[3] = 0.5
        $schedule.DailyRestQuotas[4] = 1.0

        $employee = [Activator]::CreateInstance($employeeType)
        $employee.Name = "import-check"
        for ($day = 0; $day -lt $schedule.DayCount; $day++) {
            $cell = [Activator]::CreateInstance($cellType)
            if ($day -eq 0) {
                $cell.Code = $smallShiftCode
            } elseif ($day -eq 1) {
                $cell.Code = $bigShiftCode
            } elseif ($day -eq 2) {
                $cell.Code = $publicShiftCode
            } elseif ($day -eq 3) {
                $cell.Code = $halfShiftCode
            } elseif ($day -eq 4) {
                $cell.Code = $maternityShiftCode
            } elseif ($day -eq 5) {
                $cell.Code = $dayShiftCode
            } else {
                $cell.Code = ""
            }
            $employee.Cells.Add($cell)
        }
        $schedule.Employees.Add($employee)

        $exportMethod = $exporterType.GetMethod("ExportAsync", [Reflection.BindingFlags]"Public,Static")
        [object[]]$exportArguments = @($schedule, [string]$xlsxPath)
        $exportTask = $exportMethod.Invoke($null, $exportArguments)
        $null = $exportTask.GetAwaiter().GetResult()

        $importMethod = $importerType.GetMethod(
            "ImportAsync",
            [Reflection.BindingFlags]"Public,Static",
            $null,
            [Type[]]@([string]),
            $null)
        [object[]]$importArguments = @([string]$xlsxPath)
        $importTask = $importMethod.Invoke($null, $importArguments)
        $result = $importTask.GetAwaiter().GetResult()
        $imported = $result.GetType().GetProperty("Schedule").GetValue($result)
        $warningsEnumerable = [System.Collections.IEnumerable]$result.GetType().GetProperty("Warnings").GetValue($result)
        $warningCount = 0
        foreach ($warning in $warningsEnumerable) {
            $warningCount++
        }

        if ($imported.Year -ne 2026 -or $imported.Month -ne 8) {
            throw "Imported schedule year/month should be 2026-08, got $($imported.Year)-$($imported.Month)."
        }

        if ($imported.DayCount -ne 31 -or $imported.DailyRestQuotas.Count -ne 31) {
            throw "Imported schedule should contain 31 days and 31 rest quotas."
        }

        if ($imported.DailyRestQuotas[2] -ne 1.0 -or $imported.DailyRestQuotas[3] -ne 0.5 -or $imported.DailyRestQuotas[4] -ne 1.0) {
            throw "Imported rest quotas did not preserve full-day, half-day, and maternity-leave rest values."
        }

        if ($imported.Employees.Count -ne 1 -or $imported.Employees[0].Name -ne "import-check") {
            throw "Imported employees were not preserved."
        }

        $importedEmployee = $imported.Employees[0]
        if ($importedEmployee.Cells.Count -ne 31) {
            throw "Imported employee cell count should be 31, got $($importedEmployee.Cells.Count)."
        }

        if ($importedEmployee.Cells[0].Code -ne $smallShiftCode) {
            throw "Imported day 1 should preserve small shift."
        }
        if ($importedEmployee.Cells[1].Code -ne $bigShiftCode) {
            throw "Imported day 2 should preserve big shift."
        }
        if ($importedEmployee.Cells[2].Code -ne $publicShiftCode) {
            throw "Imported day 3 should preserve public rest."
        }
        if ($importedEmployee.Cells[3].Code -ne $halfShiftCode) {
            throw "Imported day 4 should preserve half-day rest."
        }
        if ($importedEmployee.Cells[4].Code -ne $maternityShiftCode) {
            throw "Imported day 5 should preserve maternity leave."
        }
        if ($importedEmployee.Cells[5].Code -ne $dayShiftCode) {
            throw "Imported day 6 should restore exported blank white shift from work statistics."
        }
        if (-not $importedEmployee.Cells[5].IsManual) {
            throw "Restored white shift should be marked manual after import."
        }
        if ($warningCount -lt 1) {
            throw "Importer should report that it restored at least one white shift from work statistics."
        }

        $negativePath = Join-Path $outputDir "排班_2026-08_no-work-stat.xlsx"
        Copy-Item -LiteralPath $xlsxPath -Destination $negativePath -Force
        $workHeader = ([string][char]0x4E0A) + ([string][char]0x73ED)
        $noteHeader = ([string][char]0x5907) + ([string][char]0x6CE8)
        $zip = [IO.Compression.ZipFile]::Open($negativePath, [IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.GetEntry("xl/worksheets/sheet1.xml")
            $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
            try {
                $worksheetXml = $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }

            if ($worksheetXml -notmatch [regex]::Escape($workHeader)) {
                throw "Negative import sample did not contain a work-stat header to replace."
            }

            $worksheetXml = $worksheetXml -replace [regex]::Escape($workHeader), $noteHeader
            $entry.Delete()
            $entry = $zip.CreateEntry("xl/worksheets/sheet1.xml", [IO.Compression.CompressionLevel]::Fastest)
            $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
            try {
                $writer.Write($worksheetXml)
            } finally {
                $writer.Dispose()
            }
        } finally {
            $zip.Dispose()
        }

        [object[]]$negativeImportArguments = @([string]$negativePath)
        $negativeTask = $importMethod.Invoke($null, $negativeImportArguments)
        $negativeResult = $negativeTask.GetAwaiter().GetResult()
        $negativeImported = $negativeResult.GetType().GetProperty("Schedule").GetValue($negativeResult)
        $negativeWarnings = [System.Collections.IEnumerable]$negativeResult.GetType().GetProperty("Warnings").GetValue($negativeResult)
        $missingWorkHeaderWarning = $false
        foreach ($warning in $negativeWarnings) {
            if ([string]$warning -match ([regex]::Escape($workHeader))) {
                $missingWorkHeaderWarning = $true
                break
            }
        }

        if ($negativeImported.Employees[0].Cells[5].Code -eq $dayShiftCode) {
            throw "Importer restored a blank white shift even though the work-stat header was not present."
        }
        if (-not $missingWorkHeaderWarning) {
            throw "Importer should warn when the work-stat header is not present."
        }
    }

    Invoke-Step "schedule export command availability" $repoRoot {
        $scheduleViewModelSource = Get-Content -Raw -Encoding UTF8 (Join-Path $repoRoot "src\MyTools\ViewModels\ScheduleViewModel.cs")
        $exportRejectedText = ([string][char]0x5BFC) + ([string][char]0x51FA) + ([string][char]0x88AB) + ([string][char]0x62D2) + ([string][char]0x7EDD)
        $incompleteAutoGenerateText = ([string][char]0x672A) + ([string][char]0x5B8C) + ([string][char]0x6210) + ([string][char]0x81EA) + ([string][char]0x52A8) + ([string][char]0x751F) + ([string][char]0x6210)
        $cannotExportExcelText = ([string][char]0x4E0D) + ([string][char]0x80FD) + ([string][char]0x5BFC) + ([string][char]0x51FA) + " Excel"
        $forbiddenExportGateTokens = @(
            "Current.HasGenerated",
            $exportRejectedText,
            $incompleteAutoGenerateText,
            $cannotExportExcelText
        )
        foreach ($token in $forbiddenExportGateTokens) {
            if ($scheduleViewModelSource.Contains($token)) {
                throw "Schedule export should not be gated by auto-generation state; found forbidden token $token in ScheduleViewModel.cs."
            }
        }

        $newtonsoftPath = Join-Path $env:USERPROFILE ".nuget\packages\newtonsoft.json\13.0.3\lib\net45\Newtonsoft.Json.dll"
        if (Test-Path -LiteralPath $newtonsoftPath) {
            [Reflection.Assembly]::LoadFrom($newtonsoftPath) | Out-Null
        }

        $appAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "src\MyTools\bin\Release\net48\MyTools.exe"))
        $viewModelType = $appAssembly.GetType("MyTools.ViewModels.ScheduleViewModel", $true)
        $scheduleType = $appAssembly.GetType("MyTools.Services.ScheduleVersion", $true)
        $employeeType = $appAssembly.GetType("MyTools.Services.EmployeeRow", $true)
        $cellType = $appAssembly.GetType("MyTools.Services.ShiftCell", $true)

        $schedule = [Activator]::CreateInstance($scheduleType)
        $schedule.Year = 2026
        $schedule.Month = 8
        $schedule.VersionName = "export-command-check"
        for ($day = 0; $day -lt $schedule.DayCount; $day++) {
            $schedule.DailyRestQuotas.Add(0.0)
        }

        $employee = [Activator]::CreateInstance($employeeType)
        $employee.Name = "export-command-check"
        for ($day = 0; $day -lt $schedule.DayCount; $day++) {
            $employee.Cells.Add([Activator]::CreateInstance($cellType))
        }
        $schedule.Employees.Add($employee)

        $hasGeneratedProperty = $scheduleType.GetProperty("HasGenerated", [Reflection.BindingFlags]"Public,Instance")
        if ([bool]$hasGeneratedProperty.GetValue($schedule)) {
            throw "Fresh schedule should have GeneratedAt=null and HasGenerated=false."
        }

        $viewModel = [Activator]::CreateInstance($viewModelType)
        $exportCommand = $viewModelType.GetProperty("ExportExcelCommand", [Reflection.BindingFlags]"Public,Instance").GetValue($viewModel)
        if ($exportCommand.CanExecute($null)) {
            throw "Export command should be disabled before a current schedule is loaded."
        }

        $currentProperty = $viewModelType.GetProperty("Current", [Reflection.BindingFlags]"Public,Instance")
        $currentProperty.GetSetMethod($true).Invoke($viewModel, @($schedule)) | Out-Null

        if (-not $exportCommand.CanExecute($null)) {
            throw "Export command should be enabled for a loaded schedule even when GeneratedAt is null."
        }
    }

    Invoke-Step "schedule cell fill and delete" $repoRoot {
        $schedulePageSource = Get-Content -Raw -Encoding UTF8 (Join-Path $repoRoot "src\MyTools\Views\SchedulePage.xaml.cs")
        $availableWorkHeader = ([string][char]0x53EF) + ([string][char]0x4E0A)
        $requiredScheduleGridTokens = @(
            "MakeHeaderCell(`"$availableWorkHeader`", 0, 5 + days",
            "FormatStat(stats.availableWork), row, 5 + days",
            "Grid.SetColumn(delBtn, 6 + days)",
            "SetStatText(_availableWorkStatTexts, empIdx, stats.availableWork)"
        )
        foreach ($token in $requiredScheduleGridTokens) {
            if (-not $schedulePageSource.Contains($token)) {
                throw "Schedule grid should place available-work after weekend-rest and move delete button after it; missing token $token."
            }
        }

        $mouseDownStart = $schedulePageSource.IndexOf("private void CellButton_PreviewMouseLeftButtonDown")
        $mouseMoveStart = $schedulePageSource.IndexOf("private void CellButton_PreviewMouseMove")
        if ($mouseDownStart -lt 0 -or $mouseMoveStart -le $mouseDownStart) {
            throw "Could not locate schedule cell mouse handlers in SchedulePage.xaml.cs."
        }

        $mouseDownBody = $schedulePageSource.Substring($mouseDownStart, $mouseMoveStart - $mouseDownStart)
        if ($mouseDownBody.Contains("CaptureMouse()")) {
            throw "Schedule cell normal click path must not capture the mouse on MouseDown; it prevents Button.Click from opening the shift picker."
        }

        $mouseUpStart = $schedulePageSource.IndexOf("private void CellButton_PreviewMouseLeftButtonUp")
        if ($mouseUpStart -lt 0) {
            throw "Could not locate schedule cell mouse-up handler in SchedulePage.xaml.cs."
        }

        $mouseMoveBody = $schedulePageSource.Substring($mouseMoveStart, $mouseUpStart - $mouseMoveStart)
        if (-not $mouseMoveBody.Contains("CaptureMouse()")) {
            throw "Schedule cell drag-selection path should capture the mouse only after moving into another cell."
        }

        $newtonsoftPath = Join-Path $env:USERPROFILE ".nuget\packages\newtonsoft.json\13.0.3\lib\net45\Newtonsoft.Json.dll"
        if (Test-Path -LiteralPath $newtonsoftPath) {
            [Reflection.Assembly]::LoadFrom($newtonsoftPath) | Out-Null
        }

        $appAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "src\MyTools\bin\Release\net48\MyTools.exe"))
        $viewModelType = $appAssembly.GetType("MyTools.ViewModels.ScheduleViewModel", $true)
        $scheduleType = $appAssembly.GetType("MyTools.Services.ScheduleVersion", $true)
        $employeeType = $appAssembly.GetType("MyTools.Services.EmployeeRow", $true)
        $cellType = $appAssembly.GetType("MyTools.Services.ShiftCell", $true)
        $shiftCodesType = $appAssembly.GetType("MyTools.Services.ShiftCodes", $true)
        $restShiftCode = [string][char]0x4F11

        $schedule = [Activator]::CreateInstance($scheduleType)
        $schedule.Year = 2026
        $schedule.Month = 8
        $schedule.VersionName = "fill-check"
        for ($day = 0; $day -lt $schedule.DayCount; $day++) {
            $schedule.DailyRestQuotas.Add(0.0)
        }

        for ($employeeIndex = 0; $employeeIndex -lt 2; $employeeIndex++) {
            $employee = [Activator]::CreateInstance($employeeType)
            $employee.Name = "fill-check-$employeeIndex"
            for ($day = 0; $day -lt $schedule.DayCount; $day++) {
                $employee.Cells.Add([Activator]::CreateInstance($cellType))
            }
            $schedule.Employees.Add($employee)
        }

        $schedule.Employees[0].Cells[0].Code = $restShiftCode
        $schedule.Employees[0].Cells[0].IsManual = $true

        $viewModel = [Activator]::CreateInstance($viewModelType)
        $currentProperty = $viewModelType.GetProperty("Current", [Reflection.BindingFlags]"Public,Instance")
        $currentProperty.GetSetMethod($true).Invoke($viewModel, @($schedule)) | Out-Null

        $tupleType = [System.ValueTuple[int,int]]
        $targets = [Array]::CreateInstance($tupleType, 3)
        $targets.SetValue([System.ValueTuple[int,int]]::new(0, 1), 0)
        $targets.SetValue([System.ValueTuple[int,int]]::new(1, 0), 1)
        $targets.SetValue([System.ValueTuple[int,int]]::new(1, 1), 2)

        $fillMethod = $viewModelType.GetMethod("FillCellsFromSource", [Reflection.BindingFlags]"Public,Instance")
        $fillArguments = [object[]]::new(3)
        $fillArguments[0] = [int]0
        $fillArguments[1] = [int]0
        $fillArguments[2] = $targets
        $filled = [int]$fillMethod.Invoke($viewModel, $fillArguments)
        if ($filled -ne 3) {
            throw "Bulk fill should affect 3 selected cells, got $filled."
        }

        foreach ($target in $targets) {
            $cell = $schedule.Employees[$target.Item1].Cells[$target.Item2]
            if ($cell.Code -ne $restShiftCode -or -not $cell.IsManual) {
                throw "Filled cell $($target.Item1),$($target.Item2) should contain rest shift and be marked manual."
            }
        }

        $restDaysMethod = $shiftCodesType.GetMethod("RestDays", [Reflection.BindingFlags]"Public,Static")
        $filledRestDays = [double]$restDaysMethod.Invoke($null, @($schedule.Employees[1].Cells[0].Code))
        if ($filledRestDays -ne 1.0) {
            throw "Filled rest shift should still count as 1 rest day, got $filledRestDays."
        }

        $computeRowStatsMethod = $viewModelType.GetMethod("ComputeRowStats", [Reflection.BindingFlags]"Public,Instance")
        $filledStats = $computeRowStatsMethod.Invoke($viewModel, @([int]1))
        $filledWeekendRest = [double]$filledStats.GetType().GetField("Item4").GetValue($filledStats)
        if ($filledWeekendRest -ne 2.0) {
            throw "Filled Saturday/Sunday rest shifts should count as 2 weekend rest days, got $filledWeekendRest."
        }
        $filledAvailableWork = [double]$filledStats.GetType().GetField("Item5").GetValue($filledStats)
        if ($filledAvailableWork -ne 29.0) {
            throw "Filled row available work should count unset cells plus work shifts only; expected 29, got $filledAvailableWork."
        }

        $clearMethod = $viewModelType.GetMethod("ClearCells", [Reflection.BindingFlags]"Public,Instance")
        $clearArguments = [object[]]::new(1)
        $clearArguments[0] = $targets
        $cleared = [int]$clearMethod.Invoke($viewModel, $clearArguments)
        if ($cleared -ne 3) {
            throw "Bulk clear should affect 3 selected cells, got $cleared."
        }

        foreach ($target in $targets) {
            $cell = $schedule.Employees[$target.Item1].Cells[$target.Item2]
            if ($cell.Code -ne "" -or $cell.IsManual) {
                throw "Cleared cell $($target.Item1),$($target.Item2) should be empty and not manual."
            }
        }

        if ($schedule.Employees[0].Cells[0].Code -ne $restShiftCode -or -not $schedule.Employees[0].Cells[0].IsManual) {
            throw "Bulk clear should not modify the unselected source cell."
        }

        $clearedStats = $computeRowStatsMethod.Invoke($viewModel, @([int]1))
        $clearedWeekendRest = [double]$clearedStats.GetType().GetField("Item4").GetValue($clearedStats)
        if ($clearedWeekendRest -ne 0.0) {
            throw "Cleared Saturday/Sunday rest shifts should reset weekend rest days, got $clearedWeekendRest."
        }
        $clearedAvailableWork = [double]$clearedStats.GetType().GetField("Item5").GetValue($clearedStats)
        if ($clearedAvailableWork -ne 31.0) {
            throw "Cleared row available work should treat unset cells as available work; expected 31, got $clearedAvailableWork."
        }
    }

    Invoke-Step "audio detection sample scan" $repoRoot {
        $outputDir = Join-Path ([IO.Path]::GetTempPath()) "MyToolsCodexEval\audio-detection-check"
        New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

        function Write-TestWav {
            param(
                [Parameter(Mandatory = $true)]
                [string]$Path,
                [bool]$Tone
            )

            $sampleRate = 44100
            $samples = $sampleRate
            $dataBytes = $samples * 2
            $bytes = New-Object byte[] (44 + $dataBytes)

            function Write-Ascii([int]$Offset, [string]$Text) {
                $value = [Text.Encoding]::ASCII.GetBytes($Text)
                [Array]::Copy($value, 0, $bytes, $Offset, $value.Length)
            }

            function Write-Int32([int]$Offset, [int]$Value) {
                [Array]::Copy([BitConverter]::GetBytes($Value), 0, $bytes, $Offset, 4)
            }

            function Write-Int16([int]$Offset, [int16]$Value) {
                [Array]::Copy([BitConverter]::GetBytes($Value), 0, $bytes, $Offset, 2)
            }

            Write-Ascii 0 "RIFF"
            Write-Int32 4 (36 + $dataBytes)
            Write-Ascii 8 "WAVE"
            Write-Ascii 12 "fmt "
            Write-Int32 16 16
            Write-Int16 20 1
            Write-Int16 22 1
            Write-Int32 24 $sampleRate
            Write-Int32 28 ($sampleRate * 2)
            Write-Int16 32 2
            Write-Int16 34 16
            Write-Ascii 36 "data"
            Write-Int32 40 $dataBytes

            for ($i = 0; $i -lt $samples; $i++) {
                [int16]$sample = 0
                if ($Tone) {
                    $sample = [int16]([Math]::Sin(2 * [Math]::PI * 440 * $i / $sampleRate) * 12000)
                }

                [Array]::Copy([BitConverter]::GetBytes($sample), 0, $bytes, 44 + ($i * 2), 2)
            }

            [IO.File]::WriteAllBytes($Path, $bytes)
        }

        $silentPath = Join-Path $outputDir "silent.wav"
        $tonePath = Join-Path $outputDir "tone.wav"
        Write-TestWav -Path $silentPath -Tone:$false
        Write-TestWav -Path $tonePath -Tone:$true

        $dependencyPaths = @(
            (Join-Path $env:USERPROFILE ".nuget\packages\naudio.core\2.2.1\lib\netstandard2.0\NAudio.Core.dll"),
            (Join-Path $env:USERPROFILE ".nuget\packages\naudio.wasapi\2.2.1\lib\netstandard2.0\NAudio.Wasapi.dll"),
            (Join-Path $env:USERPROFILE ".nuget\packages\naudio\2.2.1\lib\net472\NAudio.dll")
        )
        foreach ($dependencyPath in $dependencyPaths) {
            if (Test-Path -LiteralPath $dependencyPath) {
                [Reflection.Assembly]::LoadFrom($dependencyPath) | Out-Null
            }
        }

        $appAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "src\MyTools\bin\Release\net48\MyTools.exe"))
        $recorderType = $appAssembly.GetType("MyTools.Services.WasapiLoopbackAudioRecorder", $true)
        $scanMethod = $recorderType.GetMethod("ScanWaveFile", [Reflection.BindingFlags]"Public,Static")

        function Test-Audible {
            param([string]$Path)
            [object[]]$arguments = @([string]$Path)
            $stats = $scanMethod.Invoke($null, $arguments)
            return [bool]$stats.GetType().GetProperty("HasAudibleAudio").GetValue($stats)
        }

        if (Test-Audible $silentPath) {
            throw "Silent WAV was incorrectly detected as audible."
        }

        if (-not (Test-Audible $tonePath)) {
            throw "Tone WAV was incorrectly detected as silent."
        }
    }
}

if ($runAll -or $Installer) {
    Invoke-Step "installer pipeline" $repoRoot {
        & powershell -ExecutionPolicy Bypass -File scripts\Build-Installer.ps1 -Version $InstallerVersion
        Assert-InstallerArtifact -ExpectedVersion $InstallerVersion
    }
}

if ($failed) {
    Write-Host ""
    Write-Host "codex-eval: failed"
    exit 1
}

Write-Host ""
Write-Host "codex-eval: passed"
