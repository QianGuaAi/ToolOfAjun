# Lock Windows 10 to version 22H2
$path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
if (-not (Test-Path $path)) {
    New-Item -Path $path -Force | Out-Null
}
Set-ItemProperty -Path $path -Name "TargetReleaseVersion" -Value 1 -Type DWord
Set-ItemProperty -Path $path -Name "TargetReleaseVersionInfo" -Value "22H2" -Type String
Set-ItemProperty -Path $path -Name "ProductVersion" -Value "Windows 10" -Type String
Write-Host "Windows 10 version has been locked to 22H2."
