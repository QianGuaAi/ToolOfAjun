# MyTools UI Test - Alternative approach
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Write-Host "MyTools UI Test - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$myTools = Get-Process | Where-Object { $_.Path -like "*MyTools*" } | Select-Object -First 1

if ($myTools) {
    Write-Host "PID: $($myTools.Id)"
    Write-Host "MainWindowHandle: $($myTools.MainWindowHandle)"

    # Try to get window by handle
    $hwnd = $myTools.MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        Write-Host "Using MainWindowHandle for automation"

        $root = [System.Windows.Automation.AutomationElement]::RootElement

        # Try to find the window by its handle
        $cond = New-Object System.Windows.Automation.OrCondition(@(
            (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $myTools.Id)),
            (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NativeWindowHandleProperty, $hwnd))
        ))

        $window = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)

        if (-not $window) {
            # Try all children
            $allChildren = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
            Write-Host "Total root children: $($allChildren.Count)"

            foreach ($child in $allChildren) {
                try {
                    if ($child.Current.ProcessId -eq $myTools.Id) {
                        $window = $child
                        Write-Host "Found window by ProcessId match"
                        break
                    }
                } catch {}
            }
        }

        if ($window) {
            Write-Host "Window found - Name: '$($window.Current.Name)'"
            Write-Host "Window ClassName: $($window.Current.ClassName)"

            # Get ALL elements regardless of visibility
            $descendants = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
            Write-Host "`nTotal descendants: $($descendants.Count)"

            # Group by control type
            $controlTypes = @{}
            foreach ($el in $descendants) {
                try {
                    $type = $el.Current.ControlType.ProgrammaticName
                    if (-not $controlTypes.ContainsKey($type)) {
                        $controlTypes[$type] = 0
                    }
                    $controlTypes[$type]++
                } catch {}
            }

            Write-Host "`nControl type breakdown:"
            foreach ($ct in $controlTypes.GetEnumerator() | Sort-Object Value -Descending) {
                Write-Host "  $($ct.Key): $($ct.Value)"
            }

            # Try to find elements with names containing keywords
            $keywords = @("SQL", "Database", "Codex", "Hash", "File", "Verify", "Export", "Import", "Connect", "Query", "Profile")
            Write-Host "`nKeyword search:"
            foreach ($kw in $keywords) {
                $pattern = "*$kw*"
                $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $pattern)
                $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)
                if ($found.Count -gt 0) {
                    Write-Host "  [$kw] Found $($found.Count)"
                    for ($i = 0; $i -lt [Math]::Min(3, $found.Count); $i++) {
                        Write-Host "      - '$($found[$i].Current.Name)' ( $($found[$i].Current.ControlType.ProgrammaticName) )"
                    }
                }
            }
        }
    }
} else {
    Write-Host "MyTools process not found"
}
