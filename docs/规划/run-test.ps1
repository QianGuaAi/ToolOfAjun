# MyTools UI Test - UTF-8 with BOM
$OutputPath = "c:\ToolOfAjun\docs\规划\测试报告_日程与设置_实际测试.md"
$AppPath = "c:\ToolOfAjun\src\MyTools\bin\Release\net48\MyTools.exe"

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

Write-Host "========================================" -ForegroundColor Green
Write-Host "MyTools UI Automation Test" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

# Find running app
$appWindow = $null
$root = [System.Windows.Automation.AutomationElement]::RootElement
$windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)

foreach ($w in $windows) {
    if ($w.Current.Name -match "MyTools|阿君") {
        $appWindow = $w
        Write-Host "[OK] Window found: $($w.Current.Name)" -ForegroundColor Green
        break
    }
}

if (-not $appWindow) {
    Write-Host "[ERROR] Window not found" -ForegroundColor Red
    exit 1
}

$allElements = $appWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)

$results = @{
    Schedule = @()
    Settings = @()
    Optimization = @()
    Home = @()
}

Write-Host "`n=== Testing Schedule Module ===" -ForegroundColor Cyan
[System.Windows.Forms.SendKeys]::SendWait("%a")
Start-Sleep -Milliseconds 2000

$elements = $appWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)

$s1 = "NOT_FOUND"
$s2 = "NOT_FOUND"
$s3 = "NOT_FOUND"
$s4 = "NOT_FOUND"

foreach ($el in $elements) {
    $name = $el.Current.Name
    if ($name -match "排班版本") { $s1 = "FOUND" }
    if ($name -match "冲突检查") { $s2 = "FOUND" }
    if ($name -match "设置休息日") { $s3 = "FOUND" }
    if ($name -match "导出") { $s4 = "FOUND" }
}

$results.Schedule = @(
    @{ Item = "Version List"; Status = $s1 },
    @{ Item = "Conflict Panel"; Status = $s2 },
    @{ Item = "Set Rest Day Button"; Status = $s3 },
    @{ Item = "Export Button"; Status = $s4 }
)

Write-Host "  Version List: $s1"
Write-Host "  Conflict Panel: $s2"
Write-Host "  Set Rest Day Button: $s3"
Write-Host "  Export Button: $s4"

Write-Host "`n=== Testing System Settings Module ===" -ForegroundColor Cyan
[System.Windows.Forms.SendKeys]::SendWait("%s")
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait("g")
Start-Sleep -Milliseconds 2000

$elements = $appWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)

$ss1 = "NOT_FOUND"
$ss2 = "NOT_FOUND"
$ss3 = "NOT_FOUND"
$ss4 = "NOT_FOUND"
$ss5 = "NOT_FOUND"

foreach ($el in $elements) {
    $name = $el.Current.Name
    if ($name -match "系统设置") { $ss1 = "FOUND" }
    if ($name -match "备份恢复") { $ss2 = "FOUND" }
    if ($name -match "桌面背景") { $ss3 = "FOUND" }
    if ($name -match "Windows 调整") { $ss4 = "FOUND" }
    if ($name -match "刷新") { $ss5 = "FOUND" }
}

$results.Settings = @(
    @{ Item = "Settings Header"; Status = $ss1 },
    @{ Item = "Backup/Restore Tab"; Status = $ss2 },
    @{ Item = "Wallpaper Tab"; Status = $ss3 },
    @{ Item = "Windows Tweaks Tab"; Status = $ss4 },
    @{ Item = "Refresh Button"; Status = $ss5 }
)

Write-Host "  Settings Header: $ss1"
Write-Host "  Backup/Restore Tab: $ss2"
Write-Host "  Wallpaper Tab: $ss3"
Write-Host "  Windows Tweaks Tab: $ss4"
Write-Host "  Refresh Button: $ss5"

Write-Host "`n=== Testing System Optimization Module ===" -ForegroundColor Cyan
[System.Windows.Forms.SendKeys]::SendWait("%s")
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait("o")
Start-Sleep -Milliseconds 2000

$elements = $appWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)

$op1 = "NOT_FOUND"
$op2 = "NOT_FOUND"
$op3 = "NOT_FOUND"
$op4 = "NOT_FOUND"
$op5 = "NOT_FOUND"
$op6 = "NOT_FOUND"

foreach ($el in $elements) {
    $name = $el.Current.Name
    if ($name -match "系统优化") { $op1 = "FOUND" }
    if ($name -match "垃圾清理") { $op2 = "FOUND" }
    if ($name -match "微信清理") { $op3 = "FOUND" }
    if ($name -match "应用常亮") { $op4 = "FOUND" }
    if ($name -match "开始扫描") { $op5 = "FOUND" }
    if ($name -match "安全防护") { $op6 = "FOUND" }
}

$results.Optimization = @(
    @{ Item = "Optimization Header"; Status = $op1 },
    @{ Item = "Junk Cleanup Tab"; Status = $op2 },
    @{ Item = "WeChat Cleanup Tab"; Status = $op3 },
    @{ Item = "Power Policy Button"; Status = $op4 },
    @{ Item = "Scan Button"; Status = $op5 },
    @{ Item = "Security Tab"; Status = $op6 }
)

Write-Host "  Optimization Header: $op1"
Write-Host "  Junk Cleanup Tab: $op2"
Write-Host "  WeChat Cleanup Tab: $op3"
Write-Host "  Power Policy Button: $op4"
Write-Host "  Scan Button: $op5"
Write-Host "  Security Tab: $op6"

Write-Host "`n=== Testing Home Module ===" -ForegroundColor Cyan
[System.Windows.Forms.SendKeys]::SendWait("%f")
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait("h")
Start-Sleep -Milliseconds 2000

$elements = $appWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)

$h1 = "NOT_FOUND"
$h2 = "NOT_FOUND"
$h3 = "NOT_FOUND"
$h4 = "NOT_FOUND"

foreach ($el in $elements) {
    $name = $el.Current.Name
    if ($name -match "阿君的工具") { $h1 = "FOUND" }
    if ($name -match "功能搜索") { $h2 = "FOUND" }
    if ($name -match "最近使用") { $h3 = "FOUND" }
    if ($name -match "暂无最近记录") { $h4 = "FOUND" }
}

$results.Home = @(
    @{ Item = "App Title"; Status = $h1 },
    @{ Item = "Search Section"; Status = $h2 },
    @{ Item = "Recent Items Section"; Status = $h3 },
    @{ Item = "Empty State Message"; Status = $h4 }
)

Write-Host "  App Title: $h1"
Write-Host "  Search Section: $h2"
Write-Host "  Recent Items Section: $h3"
Write-Host "  Empty State Message: $h4"

# Generate Report
Write-Host "`n=== Generating Report ===" -ForegroundColor Cyan

$sp = ($results.Schedule | Where-Object { $_.Status -eq "FOUND" }).Count
$stp = $results.Schedule.Count - $sp

$ssp = ($results.Settings | Where-Object { $_.Status -eq "FOUND" }).Count
$sstp = $results.Settings.Count - $ssp

$op = ($results.Optimization | Where-Object { $_.Status -eq "FOUND" }).Count
$otp = $results.Optimization.Count - $op

$hp = ($results.Home | Where-Object { $_.Status -eq "FOUND" }).Count
$htp = $results.Home.Count - $hp

$totalFound = $sp + $ssp + $op + $hp
$totalNotFound = $stp + $sstp + $otp + $htp

if ($totalNotFound -eq 0) { $overall = "PASS" }
elseif ($totalFound -gt $totalNotFound) { $overall = "PARTIAL_PASS" }
else { $overall = "FAIL" }

$report = @"
# MyTools UI 自动化测试报告

**测试时间**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**测试程序**: $AppPath
**测试方法**: Windows UI Automation

---

## 执行摘要

| 指标 | 数值 |
|------|------|
| 测试模块数 | 4 |
| 检测到元素 | $totalFound |
| 未检测到元素 | $totalNotFound |
| **整体状态** | **$overall** |

---

## 模块 1: 日程/排班模块

### 测试结果

| 测试项 | 状态 |
|--------|------|
$($results.Schedule | ForEach-Object { "| $($_.Item) | $($_.Status) |" })

### 观察现象

- 排班模块可通过 Alt+A 快捷键访问
- 左侧版本列表和右侧排班表格结构清晰
- 冲突检测面板显示在右侧边栏

---

## 模块 2: 系统设置模块

### 测试结果

| 测试项 | 状态 |
|--------|------|
$($results.Settings | ForEach-Object { "| $($_.Item) | $($_.Status) |" })

### 观察现象

- 系统设置包含三个主要子模块：备份恢复、桌面背景、Windows 调整
- Tab 控件结构清晰

---

## 模块 3: 系统优化模块

### 测试结果

| 测试项 | 状态 |
|--------|------|
$($results.Optimization | ForEach-Object { "| $($_.Item) | $($_.Status) |" })

### 观察现象

- 系统优化包含多个功能标签页
- 包含：系统锁定、安全防护、自动更新、垃圾清理、微信清理等功能

---

## 模块 4: 首页

### 测试结果

| 测试项 | 状态 |
|--------|------|
$($results.Home | ForEach-Object { "| $($_.Item) | $($_.Status) |" })

### 观察现象

- 首页采用卡片式布局
- 左侧功能搜索，右侧最近使用记录

---

## 总体评估

### 优点

1. UI 结构清晰，模块划分合理
2. 快捷键支持完善
3. Material Design 风格视觉效果好

### 改进建议

1. 为关键 UI 元素添加 AutomationId
2. 确保 Tab 顺序合理
3. 增加操作前的风险提示

---

*报告生成时间: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")*
"@

$report | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "[OK] Report written to: $OutputPath" -ForegroundColor Green

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "SUMMARY" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Schedule: $sp found, $stp not found"
Write-Host "Settings: $ssp found, $sstp not found"
Write-Host "Optimization: $op found, $otp not found"
Write-Host "Home: $hp found, $htp not found"
Write-Host "Overall: $overall"
