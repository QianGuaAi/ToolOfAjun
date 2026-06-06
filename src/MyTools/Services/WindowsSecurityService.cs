using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace MyTools.Services
{
    public static class WindowsSecurityService
    {
        private const string DefenderPolicyKey =
            @"SOFTWARE\Policies\Microsoft\Windows Defender";

        private const string DefenderRtpKey =
            @"SOFTWARE\Microsoft\Windows Defender\Real-Time Protection";

        private const string WuPolicyAuKey =
            @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU";

        public static bool GetDefenderRealtimeStatus()
        {
            var mpPreferenceStatus = TryGetDefenderRealtimeStatusFromPowerShell();
            if (mpPreferenceStatus.HasValue)
            {
                return mpPreferenceStatus.Value;
            }

            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(DefenderPolicyKey))
                {
                    if (key?.GetValue("DisableAntiSpyware") is int v1 && v1 == 1)
                        return false;
                }

                using (var key = Registry.LocalMachine.OpenSubKey(DefenderRtpKey))
                {
                    if (key?.GetValue("DisableRealtimeMonitoring") is int v2 && v2 == 1)
                        return false;
                }
            }
            catch { }

            return true;
        }

        public static async Task SetDefenderRealtimeAsync(bool enable)
        {
            string script = enable
                ? @"if (Get-Command Set-MpPreference -ErrorAction SilentlyContinue) {
    Set-MpPreference -DisableRealtimeMonitoring $false -DisableBehaviorMonitoring $false -DisableBlockAtFirstSeen $false -DisableIOAVProtection $false -DisableScriptScanning $false -ErrorAction SilentlyContinue
}"
                : @"if (Get-Command Set-MpPreference -ErrorAction SilentlyContinue) {
    Set-MpPreference -DisableRealtimeMonitoring $true -DisableBehaviorMonitoring $true -DisableBlockAtFirstSeen $true -DisableIOAVProtection $true -DisableScriptScanning $true -ErrorAction SilentlyContinue
}";

            await RunElevatedScriptAsync(script, waitForExit: true);
        }

        public static bool GetAutoUpdateStatus()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(WuPolicyAuKey))
                {
                    if (key?.GetValue("NoAutoUpdate") is int v && v == 1)
                        return false;
                }
            }
            catch { }

            return true;
        }

        public static async Task SetAutoUpdateAsync(bool enable)
        {
            string script = enable
                ? @"$p='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
if (Test-Path $p) {
    Remove-ItemProperty -Path $p -Name 'NoAutoUpdate' -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $p -Name 'AUOptions' -ErrorAction SilentlyContinue
}
foreach ($svc in 'wuauserv','UsoSvc','BITS','DoSvc') {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Set-Service -Name $svc -StartupType Manual -ErrorAction SilentlyContinue
        Start-Service -Name $svc -ErrorAction SilentlyContinue
    }
}"
                : @"$p='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
if (!(Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
Set-ItemProperty -Path $p -Name 'NoAutoUpdate' -Value 1 -Type DWord
Set-ItemProperty -Path $p -Name 'AUOptions'    -Value 1 -Type DWord
foreach ($svc in 'UsoSvc','wuauserv','BITS','DoSvc') {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
    }
}";

            await RunElevatedScriptAsync(script, waitForExit: true);
        }

        public static async Task TriggerImmediateUpdateAsync()
        {
            string script = @"# Ensure auto-update policy not blocking
$p='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
if (Test-Path $p) {
    Remove-ItemProperty -Path $p -Name 'NoAutoUpdate' -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $p -Name 'AUOptions'    -ErrorAction SilentlyContinue
}

# Start required services
foreach ($svc in 'wuauserv','UsoSvc','BITS','DoSvc') {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Set-Service -Name $svc -StartupType Manual -ErrorAction SilentlyContinue
        Start-Service -Name $svc -ErrorAction SilentlyContinue
    }
}
Start-Sleep   -Seconds 2

# Trigger Windows 10/11 foreground scan first, then keep older command fallback.
$uso = ""$env:SystemRoot\System32\UsoClient.exe""
if (Test-Path $uso) {
    & $uso RefreshSettings
    Start-Sleep -Seconds 2
    & $uso StartInteractiveScan
    Start-Sleep -Seconds 5
    & $uso StartDownload
    Start-Sleep -Seconds 3
    & $uso StartInstall
}

if (Get-Command wuauclt.exe -ErrorAction SilentlyContinue) {
    wuauclt.exe /detectnow
    wuauclt.exe /updatenow
}

# Open Windows Update settings so user can monitor progress
Start-Process 'ms-settings:windowsupdate' -ErrorAction SilentlyContinue";

            await RunElevatedScriptAsync(script, waitForExit: true);
        }

        private static bool? TryGetDefenderRealtimeStatusFromPowerShell()
        {
            const string script = "$p=Get-MpPreference -ErrorAction Stop; if ([bool]$p.DisableRealtimeMonitoring) { 'disabled' } else { 'enabled' }";
            try
            {
                var output = RunPowerShellForOutput(script, 5000);
                if (output.IndexOf("enabled", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (output.IndexOf("disabled", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return false;
                }
            }
            catch
            {
                // Registry fallback below.
            }

            return null;
        }

        private static string RunPowerShellForOutput(string script, int timeoutMs)
        {
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"& { " + EscapeForPowerShellCommand(script) + " }\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                if (!process.Start())
                {
                    return string.Empty;
                }

                var outputTask = process.StandardOutput.ReadToEndAsync();
                if (!process.WaitForExit(timeoutMs))
                {
                    try { process.Kill(); } catch { }
                    return string.Empty;
                }

                return outputTask.GetAwaiter().GetResult() ?? string.Empty;
            }
        }

        private static string EscapeForPowerShellCommand(string script)
        {
            return (script ?? string.Empty)
                .Replace("`", "``")
                .Replace("\"", "`\"");
        }

        private static async Task RunElevatedScriptAsync(string script, bool waitForExit)
        {
            string tempPath = Path.Combine(
                Path.GetTempPath(),
                $"mytools_{Guid.NewGuid():N}.ps1");

            using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                await writer.WriteAsync(script).ConfigureAwait(false);
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{tempPath}\"",
                    UseShellExecute = true,
                    Verb = "runas"
                };

                var proc = Process.Start(psi);
                if (proc != null && waitForExit)
                {
                    await Task.Run(() => proc.WaitForExit(60000));
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                throw new OperationCanceledException("用户取消了 UAC 授权。");
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }
    }
}
