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
                ? "Set-MpPreference -DisableRealtimeMonitoring $false" +
                  " -DisableBehaviorMonitoring $false" +
                  " -DisableBlockAtFirstSeen $false" +
                  " -DisableIOAVProtection $false" +
                  " -DisableScriptScanning $false" +
                  " -ErrorAction SilentlyContinue"
                : "Set-MpPreference -DisableRealtimeMonitoring $true" +
                  " -DisableBehaviorMonitoring $true" +
                  " -DisableBlockAtFirstSeen $true" +
                  " -DisableIOAVProtection $true" +
                  " -DisableScriptScanning $true" +
                  " -ErrorAction SilentlyContinue";

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
Start-Service -Name wuauserv -ErrorAction SilentlyContinue
Start-Service -Name UsoSvc   -ErrorAction SilentlyContinue"
                : @"$p='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
if (!(Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
Set-ItemProperty -Path $p -Name 'NoAutoUpdate' -Value 1 -Type DWord
Set-ItemProperty -Path $p -Name 'AUOptions'    -Value 1 -Type DWord
Stop-Service -Name UsoSvc   -Force -ErrorAction SilentlyContinue
Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue";

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
Start-Service -Name wuauserv -ErrorAction SilentlyContinue
Start-Service -Name UsoSvc   -ErrorAction SilentlyContinue
Start-Sleep   -Seconds 2

# Trigger interactive foreground scan (highest priority, fastest)
$uso = ""$env:SystemRoot\System32\UsoClient.exe""
if (Test-Path $uso) {
    & $uso StartInteractiveScan
    Start-Sleep -Seconds 5
    & $uso StartDownload
    Start-Sleep -Seconds 3
    & $uso StartInstall
}

# Open Windows Update settings so user can monitor progress
Start-Process 'ms-settings:windowsupdate' -ErrorAction SilentlyContinue";

            await RunElevatedScriptAsync(script, waitForExit: false);
        }

        private static async Task RunElevatedScriptAsync(string script, bool waitForExit)
        {
            string tempPath = Path.Combine(
                Path.GetTempPath(),
                $"mytools_{Guid.NewGuid():N}.ps1");

            File.WriteAllText(tempPath, script, new UTF8Encoding(false));

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
