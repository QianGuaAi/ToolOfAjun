using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
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

        private static readonly string[] AutoUpdateBlockingServices =
        {
            "wuauserv",
            "UsoSvc",
            "BITS",
            "DoSvc",
            "WaaSMedicSvc"
        };

        private static readonly string[] AutoUpdateScheduledTasks =
        {
            @"\Microsoft\Windows\WindowsUpdate\Scheduled Start",
            @"\Microsoft\Windows\WindowsUpdate\sih",
            @"\Microsoft\Windows\WindowsUpdate\sihboot",
            @"\Microsoft\Windows\UpdateOrchestrator\Schedule Scan",
            @"\Microsoft\Windows\UpdateOrchestrator\Schedule Scan Static Task",
            @"\Microsoft\Windows\UpdateOrchestrator\USO_UxBroker",
            @"\Microsoft\Windows\UpdateOrchestrator\Maintenance Install",
            @"\Microsoft\Windows\UpdateOrchestrator\Reboot",
            @"\Microsoft\Windows\UpdateOrchestrator\Reboot_AC",
            @"\Microsoft\Windows\UpdateOrchestrator\Reboot_Battery",
            @"\Microsoft\Windows\UpdateOrchestrator\Refresh Settings"
        };

        private const int ServiceStartDisabled = 4;

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
            return IsAutoUpdateRestoreReady();
        }

        public static bool IsAutoUpdateStopEnforced()
        {
            return IsAutoUpdatePolicyDisabled()
                && AreExistingServicesDisabled(AutoUpdateBlockingServices)
                && AreExistingScheduledTasksDisabled(AutoUpdateScheduledTasks);
        }

        public static bool IsAutoUpdateRestoreReady()
        {
            return !IsAutoUpdatePolicyDisabled()
                && !AnyExistingServiceDisabled(AutoUpdateBlockingServices)
                && AreExistingScheduledTasksEnabled(AutoUpdateScheduledTasks);
        }

        private static bool IsAutoUpdatePolicyDisabled()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(WuPolicyAuKey))
                {
                    if (key?.GetValue("NoAutoUpdate") is int v && v == 1)
                        return true;
                }
            }
            catch { }

            return false;
        }

        private static bool AreExistingServicesDisabled(string[] serviceNames)
        {
            var foundAny = false;
            foreach (var serviceName in serviceNames)
            {
                var startValue = GetServiceStartValue(serviceName);
                if (!startValue.HasValue)
                {
                    continue;
                }

                foundAny = true;
                if (startValue.Value != ServiceStartDisabled)
                {
                    return false;
                }
            }

            return foundAny;
        }

        private static bool AnyExistingServiceDisabled(string[] serviceNames)
        {
            foreach (var serviceName in serviceNames)
            {
                var startValue = GetServiceStartValue(serviceName);
                if (startValue.HasValue && startValue.Value == ServiceStartDisabled)
                {
                    return true;
                }
            }

            return false;
        }

        private static int? GetServiceStartValue(string serviceName)
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(
                    @"SYSTEM\CurrentControlSet\Services\" + serviceName))
                {
                    if (key?.GetValue("Start") is int value)
                    {
                        return value;
                    }
                }
            }
            catch { }

            return null;
        }

        private static bool AreExistingScheduledTasksDisabled(string[] taskPaths)
        {
            return CheckExistingScheduledTasks(taskPaths, expectedEnabled: false);
        }

        private static bool AreExistingScheduledTasksEnabled(string[] taskPaths)
        {
            return CheckExistingScheduledTasks(taskPaths, expectedEnabled: true);
        }

        private static bool CheckExistingScheduledTasks(string[] taskPaths, bool expectedEnabled)
        {
            object scheduler = null;
            try
            {
                var schedulerType = Type.GetTypeFromProgID("Schedule.Service");
                if (schedulerType == null)
                {
                    return false;
                }

                scheduler = Activator.CreateInstance(schedulerType);
                ((dynamic)scheduler).Connect();

                foreach (var taskPath in taskPaths)
                {
                    var state = GetScheduledTaskState(scheduler, taskPath);
                    if (state == ScheduledTaskState.Missing)
                    {
                        continue;
                    }

                    if (state == ScheduledTaskState.Unknown)
                    {
                        return false;
                    }

                    if (expectedEnabled && state != ScheduledTaskState.Enabled)
                    {
                        return false;
                    }

                    if (!expectedEnabled && state != ScheduledTaskState.Disabled)
                    {
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComObject(scheduler);
            }
        }

        private static ScheduledTaskState GetScheduledTaskState(object scheduler, string taskPath)
        {
            object folder = null;
            object task = null;
            try
            {
                var lastSlash = taskPath.LastIndexOf('\\');
                if (lastSlash <= 0 || lastSlash >= taskPath.Length - 1)
                {
                    return ScheduledTaskState.Unknown;
                }

                var folderPath = taskPath.Substring(0, lastSlash);
                var taskName = taskPath.Substring(lastSlash + 1);
                folder = ((dynamic)scheduler).GetFolder(folderPath);
                task = ((dynamic)folder).GetTask(taskName);

                return ((dynamic)task).Enabled
                    ? ScheduledTaskState.Enabled
                    : ScheduledTaskState.Disabled;
            }
            catch (COMException ex)
            {
                return IsTaskSchedulerNotFound(ex)
                    ? ScheduledTaskState.Missing
                    : ScheduledTaskState.Unknown;
            }
            catch
            {
                return ScheduledTaskState.Unknown;
            }
            finally
            {
                ReleaseComObject(task);
                ReleaseComObject(folder);
            }
        }

        private static bool IsTaskSchedulerNotFound(COMException ex)
        {
            return ex.ErrorCode == unchecked((int)0x80070002)
                || ex.ErrorCode == unchecked((int)0x80070003);
        }

        private static void ReleaseComObject(object instance)
        {
            try
            {
                if (instance != null && Marshal.IsComObject(instance))
                {
                    Marshal.FinalReleaseComObject(instance);
                }
            }
            catch { }
        }

        private enum ScheduledTaskState
        {
            Missing,
            Enabled,
            Disabled,
            Unknown
        }

        public static async Task SetAutoUpdateAsync(bool enable)
        {
            string script = enable
                ? BuildRestoreAutoUpdateScript()
                : BuildStopAutoUpdateScript();

            await RunElevatedScriptAsync(script, waitForExit: true);
        }

        public static async Task TriggerImmediateUpdateAsync()
        {
            string script = BuildRestoreAutoUpdateScript() + @"
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

        private static string BuildStopAutoUpdateScript()
        {
            return @"$wu='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
if (!(Test-Path $wu)) { New-Item -Path $wu -Force | Out-Null }
$p=Join-Path $wu 'AU'
if (!(Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
New-ItemProperty -Path $p -Name 'NoAutoUpdate' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $p -Name 'AUOptions' -Value 1 -PropertyType DWord -Force | Out-Null

$tasks = @(
    '\Microsoft\Windows\WindowsUpdate\Scheduled Start',
    '\Microsoft\Windows\WindowsUpdate\sih',
    '\Microsoft\Windows\WindowsUpdate\sihboot',
    '\Microsoft\Windows\UpdateOrchestrator\Schedule Scan',
    '\Microsoft\Windows\UpdateOrchestrator\Schedule Scan Static Task',
    '\Microsoft\Windows\UpdateOrchestrator\USO_UxBroker',
    '\Microsoft\Windows\UpdateOrchestrator\Maintenance Install',
    '\Microsoft\Windows\UpdateOrchestrator\Reboot',
    '\Microsoft\Windows\UpdateOrchestrator\Reboot_AC',
    '\Microsoft\Windows\UpdateOrchestrator\Reboot_Battery',
    '\Microsoft\Windows\UpdateOrchestrator\Refresh Settings'
)
foreach ($task in $tasks) {
    & schtasks.exe /Change /TN $task /Disable 2>$null | Out-Null
}

foreach ($svc in 'WaaSMedicSvc','UsoSvc','wuauserv','BITS','DoSvc') {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
        Set-Service -Name $svc -StartupType Disabled -ErrorAction SilentlyContinue
    }

    $svcPath = ""HKLM:\SYSTEM\CurrentControlSet\Services\$svc""
    if (Test-Path $svcPath) {
        Set-ItemProperty -Path $svcPath -Name 'Start' -Value 4 -ErrorAction SilentlyContinue
    }
}";
        }

        private static string BuildRestoreAutoUpdateScript()
        {
            return @"$p='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
if (Test-Path $p) {
    Remove-ItemProperty -Path $p -Name 'NoAutoUpdate' -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $p -Name 'AUOptions' -ErrorAction SilentlyContinue
}

$tasks = @(
    '\Microsoft\Windows\WindowsUpdate\Scheduled Start',
    '\Microsoft\Windows\WindowsUpdate\sih',
    '\Microsoft\Windows\WindowsUpdate\sihboot',
    '\Microsoft\Windows\UpdateOrchestrator\Schedule Scan',
    '\Microsoft\Windows\UpdateOrchestrator\Schedule Scan Static Task',
    '\Microsoft\Windows\UpdateOrchestrator\USO_UxBroker',
    '\Microsoft\Windows\UpdateOrchestrator\Maintenance Install',
    '\Microsoft\Windows\UpdateOrchestrator\Reboot',
    '\Microsoft\Windows\UpdateOrchestrator\Reboot_AC',
    '\Microsoft\Windows\UpdateOrchestrator\Reboot_Battery',
    '\Microsoft\Windows\UpdateOrchestrator\Refresh Settings'
)
foreach ($task in $tasks) {
    & schtasks.exe /Change /TN $task /Enable 2>$null | Out-Null
}

$serviceStartupTypes = @{
    wuauserv = 'Manual'
    UsoSvc = 'Automatic'
    BITS = 'Manual'
    DoSvc = 'Automatic'
    WaaSMedicSvc = 'Manual'
}
$serviceStartValues = @{
    wuauserv = 3
    UsoSvc = 2
    BITS = 3
    DoSvc = 2
    WaaSMedicSvc = 3
}
foreach ($entry in $serviceStartupTypes.GetEnumerator()) {
    $svc = $entry.Key
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Set-Service -Name $svc -StartupType $entry.Value -ErrorAction SilentlyContinue
    }

    $svcPath = ""HKLM:\SYSTEM\CurrentControlSet\Services\$svc""
    if (Test-Path $svcPath) {
        Set-ItemProperty -Path $svcPath -Name 'Start' -Value $serviceStartValues[$svc] -ErrorAction SilentlyContinue
    }
}

foreach ($svc in 'BITS','DoSvc','wuauserv','UsoSvc') {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Start-Service -Name $svc -ErrorAction SilentlyContinue
    }
}";
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
