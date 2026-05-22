using System;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public static class PowerPolicyService
    {
        public static Task ApplyAlwaysOnPolicyAsync(CancellationToken cancellationToken)
        {
            var script = @"
$ErrorActionPreference = 'Stop'

powercfg.exe /change monitor-timeout-ac 0
powercfg.exe /change monitor-timeout-dc 0
powercfg.exe /change disk-timeout-ac 0
powercfg.exe /change disk-timeout-dc 0
powercfg.exe /change standby-timeout-ac 0
powercfg.exe /change standby-timeout-dc 0
powercfg.exe /change hibernate-timeout-ac 0
powercfg.exe /change hibernate-timeout-dc 0
powercfg.exe /hibernate off
";

            AppLogService.Information("Applying always-on power policy: monitor/disk/sleep/hibernate timeouts disabled.");
            return ElevatedScriptRunner.RunElevatedScriptAsync(script, true, cancellationToken);
        }
    }
}
