<#
Name......: Uninstaller.ps1
Version...: 19.11.1
Author....: Dario CORRADA

This script uninstall software looking at register keys
#>

$WarningPreference = 'SilentlyContinue'

$searchkey = Read-Host "Software to uninstall"

$record = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
          Get-ItemProperty | Where-Object {$_.DisplayName -match $searchkey } | Select-Object -Property DisplayName, UninstallString

foreach ($elem in $record) {
    Write-Host -ForegroundColor Cyan -NoNewline $elem.DisplayName
    $elem.UninstallString -match "MsiExec\.exe .+\{([A-Z0-9\-]+)\}" > $null
    Write-Host -ForegroundColor Yellow "`t`t", $Matches[1]

}

$string = Read-Host "Insert uninstall string"

msiexec "/X{$string}"
