# MyTools UI Test - Simplified
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Write-Host "MyTools UI Test - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$myTools = Get-Process | Where-Object { $_.Path -like "*MyTools*" } | Select-Object -First 1

if ($myTools) {
    Write-Host "PID: $($myTools.Id)"
    Write-Host "Title: $($myTools.MainWindowTitle)"

    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $myTools.Id)
    $window = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)

    if ($window) {
        Write-Host "Automation window obtained"

        # Get all buttons
        $btnCond = New-Object System.Windows.Automation.AndCondition(@(
            (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)),
            (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::IsOffscreenProperty, $false))
        ))
        $buttons = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)

        Write-Host "`n=== Found $($buttons.Count) visible buttons ==="
        $names = @()
        for ($i = 0; $i -lt [Math]::Min($buttons.Count, 50); $i++) {
            $name = $buttons[$i].Current.Name
            if ($name -and $name.Trim() -ne "") {
                $names += $name
            }
        }
        $names | Sort-Object -Unique | ForEach-Object { Write-Host "  - $_" }

        # Check for SQL related elements
        Write-Host "`n=== SQL Related Elements ==="
        $sqlPatterns = @("SQL", "Database", "Server", "Query", "Export", "Connect")
        foreach ($p in $sqlPatterns) {
            $cond = New-Object System.Windows.Automation.OrCondition(@(
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$p*")),
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$($p.ToLower())*"))
            ))
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)
            if ($found.Count -gt 0) {
                Write-Host "  [$p] Found $($found.Count) elements"
            }
        }

        # Check for Codex related elements
        Write-Host "`n=== Codex Related Elements ==="
        $codexPatterns = @("Codex", "Account", "Profile", "Token", "Import", "Export")
        foreach ($p in $codexPatterns) {
            $cond = New-Object System.Windows.Automation.OrCondition(@(
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$p*")),
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$($p.ToLower())*"))
            ))
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)
            if ($found.Count -gt 0) {
                Write-Host "  [$p] Found $($found.Count) elements"
            }
        }

        # Check for Hash related elements
        Write-Host "`n=== Hash Related Elements ==="
        $hashPatterns = @("Hash", "File", "Verify", "MD5", "SHA", "CRC", "Compute")
        foreach ($p in $hashPatterns) {
            $cond = New-Object System.Windows.Automation.OrCondition(@(
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$p*")),
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$($p.ToLower())*"))
            ))
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)
            if ($found.Count -gt 0) {
                Write-Host "  [$p] Found $($found.Count) elements"
            }
        }

        # Get text boxes
        Write-Host "`n=== Text Input Fields ==="
        $tbCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
        $textboxes = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tbCond)
        Write-Host "Found $($textboxes.Count) text input fields"

        # Get tab controls
        Write-Host "`n=== Tab Controls ==="
        $tabCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Tab)
        $tabs = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tabCond)
        Write-Host "Found $($tabs.Count) tab controls"

        Write-Host "`nTest completed successfully"
    } else {
        Write-Host "Could not get automation window"
    }
} else {
    Write-Host "MyTools process not found"
}
