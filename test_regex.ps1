$pattern = [regex]::new('^(wxid_[A-Za-z0-9_]+|\d+)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)
$tests = @('wxid_6evo1mkpqh1c22_d613', '262679118', '3523174748', 'nt_qq', 'Tencent Files', 'xwechat_files', 'wxid_abc', '123456')
foreach ($t in $tests) {
    Write-Host "$t : $($pattern.IsMatch($t))"
}
