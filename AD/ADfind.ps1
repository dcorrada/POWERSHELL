<#
Name......: ADfind.ps1
Version...: 22.06.1
Author....: Dario CORRADA

Loooking for a PC at Active Directory
#>


# import Active Directory module
$ErrorActionPreference= 'Stop'
try {
    Import-Module ActiveDirectory
} catch {
    Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
    Import-Module ActiveDirectory
}
$ErrorActionPreference= 'Inquire'

$computer_name = Read-Host "Hostname"

$ErrorActionPreference= 'Stop'
try {
    Get-ADComputer -Identity $computer_name -Properties *
} catch {
    Write-Host -ForegroundColor Red "[$computer_name] not found"
}
$ErrorActionPreference= 'Inquire'
Pause
