if (-not ($env:Path -like "*C:\Program Files\WindowsPowerShell\Scripts*")) {
    $env:Path += ";C:\Program Files\WindowsPowerShell\Scripts"
}

Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue > $null


if (-not (Get-PSRepository | Where-Object {$_.Name -eq "PSGallery" -and $_.InstallationPolicy -eq "Trusted"})) {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
}

Install-Script -Name Get-WindowsAutoPilotInfo -force -ErrorAction SilentlyContinue

Get-WindowsAutoPilotInfo | Select-Object "Device Serial Number","Windows Product ID","Hardware Hash" | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1

Uninstall-Script -Name Get-WindowsAutoPilotInfo

if ($env:Path -like "*C:\Program Files\WindowsPowerShell\Scripts*") {
$Delete = "C:\Program Files\WindowsPowerShell\Scripts"
$paths = $env:Path -split ';' | Where-Object { $_ -ne $Delete }
$newPath = $paths -join ';'
$env:Path = $newPath
}

$nugetSource = Get-PackageSource -Name "NuGet" -ErrorAction SilentlyContinue
if ($nugetSource) {
    Unregister-PackageSource -Name "NuGet"
}

if (-not (Get-PSRepository | Where-Object {$_.Name -eq "PSGallery" -and $_.InstallationPolicy -eq "Trusted"})) {
    Unregister-PSRepository -Name PSGallery
}
