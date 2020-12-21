<#
Name......: Uninstaller.ps1
Version...: 19.11.1
Author....: Dario CORRADA

Questo script serve per disinstallare software usando le chiavi di registro
(serve nei casi non sia presente nella lista di uninstall)
#>

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'

$searchkey = Read-Host "Chiave di ricerca"

$record = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
          Get-ItemProperty | Where-Object {$_.DisplayName -match $searchkey } | Select-Object -Property DisplayName, UninstallString

foreach ($elem in $record) {
    Write-Host -ForegroundColor Cyan -NoNewline $elem.DisplayName
    $elem.UninstallString -match "MsiExec\.exe .+\{([A-Z0-9\-]+)\}" > $null
    Write-Host -ForegroundColor Yellow "`t`t", $Matches[1]

}

$string = Read-Host "Inserisci la stringa di uninstall"

msiexec "/X{$string}"
