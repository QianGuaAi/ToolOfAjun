$docs = [Environment]::GetFolderPath('MyDocuments')
Write-Host "MyDocuments: $docs"
Write-Host "GetFullPath: $([System.IO.Path]::GetFullPath($docs))"
Write-Host "Exists: $([System.IO.Directory]::Exists($docs))"

Write-Host ''
Write-Host 'Contents through junction:'
try {
    foreach ($e in [System.IO.Directory]::EnumerateFileSystemEntries($docs)) {
        Write-Host "  $([System.IO.Path]::GetFileName($e))"
    }
} catch {
    Write-Host "  ERROR: $_"
}

Write-Host ''
$base = [System.IO.Path]::Combine($docs, 'xwechat_files')
Write-Host "xwechat_files path: $base"
Write-Host "Exists: $([System.IO.Directory]::Exists($base))"
if ([System.IO.Directory]::Exists($base)) {
    try {
        $children = [System.IO.Directory]::GetDirectories($base, '*', 'TopDirectoryOnly')
        Write-Host "Children ($($children.Length)):"
        foreach ($c in $children) {
            $name = [System.IO.Path]::GetFileName($c)
            Write-Host "  [$name] -> $c"
        }
    } catch {
        Write-Host "  ERROR: $_"
    }
}

Write-Host ''
$base2 = [System.IO.Path]::Combine($docs, 'Tencent Files')
Write-Host "Tencent Files path: $base2"
Write-Host "Exists: $([System.IO.Directory]::Exists($base2))"
if ([System.IO.Directory]::Exists($base2)) {
    try {
        $children = [System.IO.Directory]::GetDirectories($base2, '*', 'TopDirectoryOnly')
        Write-Host "Children ($($children.Length)):"
        foreach ($c in $children) {
            $name = [System.IO.Path]::GetFileName($c)
            Write-Host "  [$name] -> $c"
        }
    } catch {
        Write-Host "  ERROR: $_"
    }
}

Write-Host ''
$pattern = [regex]::new('^(wxid_[A-Za-z0-9]+|\d+_.+)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)
$testNames = @('wxid_6evo1mkpqh1c22_d613', '262679118', '3523174748', 'nt_qq', 'Tencent Files', 'xwechat_files')
Write-Host 'WxId Pattern:'
foreach ($n in $testNames) {
    Write-Host "  '$n': $($pattern.IsMatch($n))"
}

Write-Host ''
Write-Host 'User Shell Folders:'
try {
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders')
    $val = $key.GetValue('Personal')
    Write-Host "  Personal = $val"
    $expanded = [Environment]::ExpandEnvironmentVariables($val)
    Write-Host "  Expanded = $expanded"
    $resolved = [System.IO.Path]::GetFullPath($expanded)
    Write-Host "  GetFullPath = $resolved"
    $key.Close()
} catch {
    Write-Host "  ERROR: $_"
}

Write-Host ''
Write-Host '=== Done ==='
