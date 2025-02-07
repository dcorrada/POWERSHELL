<#
Name......: ADfind.ps1
Version...: 22.06.1
Author....: Dario CORRADA

Loooking for a PC at Active Directory
#>

# check Active Directory module
if ((Get-Module -Name ActiveDirectory -ListAvailable) -eq $null) {
    $ErrorActionPreference= 'Stop'
    try {
        Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability –Online
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to install RSAT"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'
}

$computer_name = Read-Host "Hostname"

$ErrorActionPreference= 'Stop'
try {
    Get-ADComputer -Identity $computer_name -Properties *
} catch {
    Write-Host -ForegroundColor Red "[$computer_name] not found"
}
$ErrorActionPreference= 'Inquire'
Pause
