param(
    [string]$Configuration = "Release",
    [string]$Version = "2026.6.30.0",
    [switch]$SkipFfmpeg
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$mainProject = Join-Path $repoRoot "src\MyTools\MyTools.csproj"
$uninstallerProject = Join-Path $repoRoot "src\MyTools.Uninstaller\MyTools.Uninstaller.csproj"
$installerProject = Join-Path $repoRoot "src\MyTools.Installer\MyTools.Installer.csproj"
$mainOutput = Join-Path $repoRoot "src\MyTools\bin\$Configuration\net48"
$uninstallerOutput = Join-Path $repoRoot "src\MyTools.Uninstaller\bin\$Configuration\net48\MyTools.Uninstaller.exe"
$artifactsRoot = Join-Path $repoRoot "artifacts\installer"
$payloadRoot = Join-Path $artifactsRoot "payload"
$payloadZip = Join-Path $artifactsRoot "MyToolsPayload.zip"
$setupOutput = Join-Path $repoRoot "src\MyTools.Installer\bin\$Configuration\net48\MyToolsSetup.exe"
$setupArtifact = Join-Path $artifactsRoot "MyToolsSetup.exe"

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

$dotnetExe = Resolve-Dotnet

function Invoke-CheckedDotnetBuild {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $script:dotnetExe @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed: $script:dotnetExe $($Arguments -join ' ')"
    }
}

New-Item -ItemType Directory -Force -Path $artifactsRoot | Out-Null
if (Test-Path $payloadRoot) {
    Remove-Item -LiteralPath $payloadRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $payloadRoot | Out-Null

Invoke-CheckedDotnetBuild @("build", $mainProject, "-c", $Configuration)
Invoke-CheckedDotnetBuild @("build", $uninstallerProject, "-c", $Configuration, "/p:Version=$Version")

$requiredFiles = @(
    "MyTools.exe",
    "MyTools.exe.config",
    "LockWin10_22H2.ps1",
    "NativeBinaries\README.txt",
    "NativeBinaries\ffmpeg\README.txt"
)

foreach ($relativePath in $requiredFiles) {
    $source = Join-Path $mainOutput $relativePath
    if (-not (Test-Path $source)) {
        throw "Missing payload file: $source"
    }

    $destination = Join-Path $payloadRoot $relativePath
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination -Force
}

if (Test-Path $payloadZip) {
    Remove-Item -LiteralPath $payloadZip -Force
}

Compress-Archive -Path (Join-Path $payloadRoot "*") -DestinationPath $payloadZip -CompressionLevel Optimal

Invoke-CheckedDotnetBuild @("build", $installerProject, "-c", $Configuration, "/p:Version=$Version", "/p:PayloadZip=$payloadZip", "/p:UninstallerExe=$uninstallerOutput")

Copy-Item -LiteralPath $setupOutput -Destination $setupArtifact -Force

$artifactInfo = Get-Item $setupArtifact
Write-Host "Installer created: $($artifactInfo.FullName)"
Write-Host "Size: $([Math]::Round($artifactInfo.Length / 1MB, 2)) MB"
