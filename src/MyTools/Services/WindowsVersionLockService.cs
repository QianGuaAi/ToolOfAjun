using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace MyTools.Services
{
    public static class WindowsVersionLockService
    {
        private const string CurrentVersionKey = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";

        public static WindowsVersionLockTarget GetCurrentTarget()
        {
            if (!OsVersionService.IsWindows10OrGreater)
            {
                throw new PlatformNotSupportedException("系统版本锁定仅支持 Windows 10 和 Windows 11。");
            }

            var productVersion = OsVersionService.IsWindows11OrGreater ? "Windows 11" : "Windows 10";
            var releaseVersion = ReadCurrentVersionValue("DisplayVersion");
            if (string.IsNullOrWhiteSpace(releaseVersion))
            {
                releaseVersion = ReadCurrentVersionValue("ReleaseId");
            }

            if (string.IsNullOrWhiteSpace(releaseVersion))
            {
                throw new InvalidOperationException("无法读取当前 Windows 功能版本（DisplayVersion/ReleaseId）。");
            }

            return new WindowsVersionLockTarget(productVersion, releaseVersion.Trim(), OsVersionService.DisplayName);
        }

        public static Task ApplyCurrentVersionLockAsync(CancellationToken cancellationToken)
        {
            var target = GetCurrentTarget();
            var script = $@"
$ErrorActionPreference = 'Stop'
$path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
if (-not (Test-Path $path)) {{
    New-Item -Path $path -Force | Out-Null
}}
Set-ItemProperty -Path $path -Name 'TargetReleaseVersion' -Value 1 -Type DWord
Set-ItemProperty -Path $path -Name 'TargetReleaseVersionInfo' -Value '{EscapePowerShellSingleQuotedString(target.ReleaseVersion)}' -Type String
Set-ItemProperty -Path $path -Name 'ProductVersion' -Value '{EscapePowerShellSingleQuotedString(target.ProductVersion)}' -Type String
";

            AppLogService.Information(
                "Applying Windows version lock: product = {ProductVersion}, release = {ReleaseVersion}, detected = {DetectedVersion}",
                target.ProductVersion,
                target.ReleaseVersion,
                target.DetectedVersion);

            return ElevatedScriptRunner.RunElevatedScriptAsync(script, true, cancellationToken);
        }

        private static string ReadCurrentVersionValue(string name)
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(CurrentVersionKey))
                {
                    return key?.GetValue(name) as string;
                }
            }
            catch
            {
                return null;
            }
        }

        private static string EscapePowerShellSingleQuotedString(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }
    }

    public sealed class WindowsVersionLockTarget
    {
        public WindowsVersionLockTarget(string productVersion, string releaseVersion, string detectedVersion)
        {
            ProductVersion = productVersion;
            ReleaseVersion = releaseVersion;
            DetectedVersion = detectedVersion;
        }

        public string ProductVersion { get; }
        public string ReleaseVersion { get; }
        public string DetectedVersion { get; }

        public string DisplayName => $"{ProductVersion} {ReleaseVersion}";
    }
}
