param(
    [switch]$Quick,
    [switch]$Build,
    [switch]$Installer,
    [string]$InstallerVersion = "2026.6.24.2"
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
            "NativeBinaries\ffmpeg\README.txt",
            "NativeBinaries\ffmpeg\ffmpeg.exe"
        )

        foreach ($entry in $requiredEntries) {
            if ($entries -notcontains $entry) {
                throw "Embedded payload missing required file: $entry"
            }
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
    if ($startupBody.Contains("StartSensorsBackground")) {
        throw "Startup background loads must not start hardware sensors; sensors are loaded on SystemInfo entry only."
    }

    $switchModuleMatch = [regex]::Match(
        $source,
        "private\s+void\s+SwitchModule\s*\(\s*string\s+module\s*\)\s*\{(?<body>.*?)^\s*\}\s*public\s+void\s+SwitchToHomeFromMultimedia",
        [Text.RegularExpressions.RegexOptions]::Singleline -bor [Text.RegularExpressions.RegexOptions]::Multiline)
    if (-not $switchModuleMatch.Success -or -not $switchModuleMatch.Groups["body"].Value.Contains("StartSensorsBackgroundAsync")) {
        throw "SystemInfo module entry must start sensors on demand."
    }

    $refreshMethodMatch = [regex]::Match(
        $source,
        "private\s+async\s+Task\s+RefreshSensorsAsync\s*\(\s*\)\s*\{(?<body>.*?)^\s*\}\s*private\s+static\s+readonly",
        [Text.RegularExpressions.RegexOptions]::Singleline -bor [Text.RegularExpressions.RegexOptions]::Multiline)
    if (-not $refreshMethodMatch.Success -or -not $refreshMethodMatch.Groups["body"].Value.Contains("Task.Run(() => service.ReadAll())")) {
        throw "Sensor refresh must keep HardwareSensorService.ReadAll off the UI thread."
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
    $matches = [regex]::Matches($memoryIndex, '`((?:docs[\\/]|\.agents[\\/]|\.codex[\\/]|scripts[\\/])[^`]+)`')
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
        throw "docs memory index references missing path(s): $($missing -join ', ')"
    }
}

if ($runAll -or $Quick) {
    Invoke-Step "loop memory docs" $repoRoot {
        Assert-LoopMemoryDocs
    }

    Invoke-Step "startup responsiveness guards" $repoRoot {
        Assert-StartupResponsivenessGuards
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
        if (-not $modelsJson.Contains('"object":"list"') -or -not $modelsJson.Contains('"id":"gpt-5.5"')) {
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
            throw "Cursor compatibility should preserve the client's streaming preference."
        }
        if (-not $responsesBody.Contains('"input"') -or -not $responsesBody.Contains('"tools"') -or -not $responsesBody.Contains('"max_output_tokens":7')) {
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
        if (-not $completionJson.Contains('"object":"chat.completion"') -or -not $completionJson.Contains('"content":"pong"') -or -not $completionJson.Contains('"total_tokens":2')) {
            throw "Converted chat/completions JSON lost object type, content, or usage fields."
        }

        $sseMethod = $compatType.GetMethod("TryBuildChatCompletionSsePayload", [Reflection.BindingFlags]"Public,Static")
        [object[]]$sseArguments = @($responsesJson, $chatBytes, $null)
        $sseConverted = [bool]$sseMethod.Invoke($null, $sseArguments)
        if (-not $sseConverted) {
            throw "Responses API output should convert to chat/completions SSE."
        }

        $ssePayload = [string]$sseArguments[2]
        if (-not $ssePayload.Contains('"object":"chat.completion.chunk"') -or -not $ssePayload.Contains('"content":"pong"') -or -not $ssePayload.Contains('data: [DONE]')) {
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

            $cellNode = $Sheet.SelectSingleNode("//x:c[@r='$CellRef']", $SheetNs)
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

            $cellNode = $Sheet.SelectSingleNode("//x:c[@r='$CellRef']", $SheetNs)
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
            throw "Schedule export should write weekend-rest header to AJ1, got '$weekendRestHeaderCell'."
        }

        $weekendRestValue = Get-CellText $sheetDoc $ns "AJ5"
        if ($weekendRestValue -ne "1.5") {
            throw "Schedule export should count one full and one half weekend rest day in AJ5, got '$weekendRestValue'."
        }

        $availableWorkHeaderCell = Get-CellText $sheetDoc $ns "AK1"
        if ($availableWorkHeaderCell -ne $availableWorkHeader) {
            throw "Schedule export should write available-work header to AK1, got '$availableWorkHeaderCell'."
        }

        $availableWorkValue = Get-CellText $sheetDoc $ns "AK5"
        if ($availableWorkValue -ne "26") {
            throw "Schedule export should count unset cells plus current work cells as available work in AK5, got '$availableWorkValue'."
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
                throw "Schedule export should not be gated by auto-generation state; found forbidden token '$token' in ScheduleViewModel.cs."
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
                throw "Schedule grid should place available-work after weekend-rest and move delete button after it; missing token '$token'."
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
        $fillArguments = New-Object 'object[]' 3
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
        $clearArguments = New-Object 'object[]' 1
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
