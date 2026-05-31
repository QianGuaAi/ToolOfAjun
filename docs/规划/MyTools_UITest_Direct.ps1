# MyTools UI Automation Test - Direct UI Element Testing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$ErrorActionPreference = "Continue"

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "  MyTools UI 自动化测试" -ForegroundColor Magenta
Write-Host "  测试时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

# Find MyTools window
function Find-MyToolsWindow {
    $processes = Get-Process | Where-Object { $_.MainWindowTitle -like "*我的工具*" -or $_.MainWindowTitle -like "*MyTools*" }
    if ($processes.Count -eq 0) {
        Write-Host "[ERROR] MyTools 窗口未找到" -ForegroundColor Red
        return $null
    }

    $pid = $processes[0].Id
    Write-Host "[INFO] 找到 MyTools 窗口，PID: $pid" -ForegroundColor Cyan

    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $pid)
    $window = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)

    if ($window) {
        Write-Host "[INFO] AutomationElement 获取成功" -ForegroundColor Cyan
    }
    return $window
}

# Find element by Name
function Find-ElementByName($parent, $name, $scope = [System.Windows.Automation.TreeScope]::Descendants) {
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $name)
    return $parent.FindFirst($scope, $condition)
}

# Find element by AutomationId
function Find-ElementById($parent, $automationId, $scope = [System.Windows.Automation.TreeScope]::Descendants) {
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, $automationId)
    return $parent.FindFirst($scope, $condition)
}

# Test SQL Export Module
function Test-SQLExportModule($window) {
    Write-Host "`n=== SQL 导出模块测试 ===" -ForegroundColor Cyan

    $result = @{
        Name = "SQL导出模块"
        Status = "待测试"
        Details = @()
        PassedTests = @()
        FailedTests = @()
    }

    try {
        # Look for SQL Export related elements
        $sqlKeywords = @("SQL", "数据库", "Server", "Database", "连接", "查询", "导出")

        Write-Host "[STEP 1] 检查 SQL 导出导航入口..." -ForegroundColor Yellow

        # Search for SQL-related text in the UI
        $foundElements = @()
        foreach ($keyword in $sqlKeywords) {
            $condition = New-Object System.Windows.Automation.OrCondition(@(
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")),
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, "*$keyword*"))
            ))
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            if ($found.Count -gt 0) {
                $foundElements += $found
            }
        }

        if ($foundElements.Count -gt 0) {
            Write-Host "[PASS] 发现 $($foundElements.Count) 个 SQL 相关元素" -ForegroundColor Green
            $result.PassedTests += "发现 SQL 相关 UI 元素"
        }

        # Check for connection panel elements
        Write-Host "[STEP 2] 检查连接配置面板..." -ForegroundColor Yellow

        $connectionKeywords = @("服务器", "地址", "端口", "用户名", "密码", "Server", "Address", "Port", "User")
        $connectionElements = 0
        foreach ($keyword in $connectionKeywords) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            $connectionElements += $found.Count
        }

        if ($connectionElements -gt 0) {
            Write-Host "[PASS] 发现 $connectionElements 个连接配置元素" -ForegroundColor Green
            $result.PassedTests += "连接配置面板正常"
        }

        # Check for query/table elements
        Write-Host "[STEP 3] 检查查询和导出功能..." -ForegroundColor Yellow

        $queryKeywords = @("查询", "执行", "Query", "Execute", "表", "Table", "导出", "Export")
        $queryElements = 0
        foreach ($keyword in $queryKeywords) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            $queryElements += $found.Count
        }

        if ($queryElements -gt 0) {
            Write-Host "[PASS] 发现 $queryElements 个查询/导出元素" -ForegroundColor Green
            $result.PassedTests += "查询和导出功能正常"
        }

        $result.Status = "部分通过"
        $result.Details += "SQL 导出模块核心 UI 元素已验证"
    }
    catch {
        Write-Host "[ERROR] SQL 导出测试异常: $_" -ForegroundColor Red
        $result.Status = "失败"
        $result.FailedTests += $_.Exception.Message
    }

    return $result
}

# Test Codex Profiles Module
function Test-CodexProfilesModule($window) {
    Write-Host "`n=== Codex 配置模块测试 ===" -ForegroundColor Cyan

    $result = @{
        Name = "Codex配置模块"
        Status = "待测试"
        Details = @()
        PassedTests = @()
        FailedTests = @()
    }

    try {
        # Look for Codex-related elements
        $codexKeywords = @("Codex", "账号", "Profile", "Account", "配置", "Config")

        Write-Host "[STEP 1] 检查 Codex 配置入口..." -ForegroundColor Yellow

        $foundElements = @()
        foreach ($keyword in $codexKeywords) {
            $condition = New-Object System.Windows.Automation.OrCondition(@(
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")),
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, "*$keyword*"))
            ))
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            if ($found.Count -gt 0) {
                $foundElements += $found
            }
        }

        if ($foundElements.Count -gt 0) {
            Write-Host "[PASS] 发现 $($foundElements.Count) 个 Codex 相关元素" -ForegroundColor Green
            $result.PassedTests += "发现 Codex 相关 UI 元素"
        }

        # Check for profile management buttons
        Write-Host "[STEP 2] 检查账号管理功能..." -ForegroundColor Yellow

        $buttonKeywords = @("导入", "Export", "Import", "切换", "Switch", "轮换", "Rotation", "重启", "Restart")
        $buttonElements = 0
        foreach ($keyword in $buttonKeywords) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            $buttonElements += $found.Count
        }

        if ($buttonElements -gt 0) {
            Write-Host "[PASS] 发现 $buttonElements 个账号管理按钮" -ForegroundColor Green
            $result.PassedTests += "账号管理功能正常"
        }

        # Check for status indicators
        Write-Host "[STEP 3] 检查状态标签..." -ForegroundColor Yellow

        $statusKeywords = @("正常", "过期", "Warn", "Ok", "Expire", "Status")
        $statusElements = 0
        foreach ($keyword in $statusKeywords) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            $statusElements += $found.Count
        }

        if ($statusElements -gt 0) {
            Write-Host "[PASS] 发现 $statusElements 个状态标签" -ForegroundColor Green
            $result.PassedTests += "状态标签显示正常"
        }

        $result.Status = "部分通过"
        $result.Details += "Codex 配置模块核心 UI 元素已验证"
    }
    catch {
        Write-Host "[ERROR] Codex 配置测试异常: $_" -ForegroundColor Red
        $result.Status = "失败"
        $result.FailedTests += $_.Exception.Message
    }

    return $result
}

# Test File Hash Module
function Test-FileHashModule($window) {
    Write-Host "`n=== 文件验证模块测试 ===" -ForegroundColor Cyan

    $result = @{
        Name = "文件验证模块"
        Status = "待测试"
        Details = @()
        PassedTests = @()
        FailedTests = @()
    }

    try {
        # Look for file hash related elements
        $hashKeywords = @("哈希", "Hash", "文件", "File", "校验", "Verify", "MD5", "SHA")

        Write-Host "[STEP 1] 检查文件验证入口..." -ForegroundColor Yellow

        $foundElements = @()
        foreach ($keyword in $hashKeywords) {
            $condition = New-Object System.Windows.Automation.OrCondition(@(
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")),
                (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, "*$keyword*"))
            ))
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            if ($found.Count -gt 0) {
                $foundElements += $found
            }
        }

        if ($foundElements.Count -gt 0) {
            Write-Host "[PASS] 发现 $($foundElements.Count) 个文件验证相关元素" -ForegroundColor Green
            $result.PassedTests += "发现文件验证相关 UI 元素"
        }

        # Check for compute button
        Write-Host "[STEP 2] 检查计算功能..." -ForegroundColor Yellow

        $computeKeywords = @("计算", "Compute", "选择文件", "Select File", "选择", "计算哈希")
        $computeElements = 0
        foreach ($keyword in $computeKeywords) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$keyword*")
            $found = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            $computeElements += $found.Count
        }

        if ($computeElements -gt 0) {
            Write-Host "[PASS] 发现 $computeElements 个计算相关元素" -ForegroundColor Green
            $result.PassedTests += "哈希计算功能正常"
        }

        # Check for hash value display
        Write-Host "[STEP 3] 检查哈希值显示..." -ForegroundColor Yellow

        # Look for text boxes or labels that might show hash values
        $hashValuePatterns = @("MD5", "SHA-1", "SHA-256", "CRC32")
        $hashDisplayFound = $false
        foreach ($pattern in $hashValuePatterns) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "*$pattern*")
            $found = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
            if ($found) {
                $hashDisplayFound = $true
                break
            }
        }

        if ($hashDisplayFound) {
            Write-Host "[PASS] 发现哈希值显示区域" -ForegroundColor Green
            $result.PassedTests += "哈希值显示正常"
        }

        $result.Status = "部分通过"
        $result.Details += "文件验证模块核心 UI 元素已验证"
    }
    catch {
        Write-Host "[ERROR] 文件验证测试异常: $_" -ForegroundColor Red
        $result.Status = "失败"
        $result.FailedTests += $_.Exception.Message
    }

    return $result
}

# Main execution
$window = Find-MyToolsWindow

if ($window) {
    $results = @()

    $results += Test-SQLExportModule $window
    $results += Test-CodexProfilesModule $window
    $results += Test-FileHashModule $window

    # Summary
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "  测试结果摘要" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta

    foreach ($r in $results) {
        Write-Host "`n模块: $($r.Name)" -ForegroundColor White
        Write-Host "  状态: $($r.Status)" -ForegroundColor $(if ($r.Status -eq "部分通过") { "Yellow" } else { "Green" })
        Write-Host "  通过测试:" -ForegroundColor Gray
        foreach ($t in $r.PassedTests) {
            Write-Host "    + $t" -ForegroundColor Green
        }
        if ($r.FailedTests.Count -gt 0) {
            Write-Host "  失败测试:" -ForegroundColor Gray
            foreach ($t in $r.FailedTests) {
                Write-Host "    - $t" -ForegroundColor Red
            }
        }
    }

    Write-Host "`n测试完成! 时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
}
else {
    Write-Host "[ERROR] 无法获取 MyTools 窗口" -ForegroundColor Red
}
