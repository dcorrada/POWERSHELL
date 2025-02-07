 <#
Name......: Belong2OU.ps1
Version...: 19.04.1
Author....: Dario CORRADA

This script get computer list belonging to a specific OU
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

$ou2find = Read-Host "OU to search"

# retrieving available OUs
$ou_available = Get-ADOrganizationalUnit -Filter *

$obj = Get-ADDomain
$suffix = ',' + $obj.DistinguishedName
if ($ou_available.Name -contains $ou2find) {
    Write-Host "Searching computers belonging to [$ou2find]"
    $string = "OU=" + $ou2find + $suffix
    $ErrorActionPreference= 'SilentlyContinue'
    $computer_list = Get-ADComputer -Filter * -SearchBase $string
    $ErrorActionPreference= 'Inquire'
    if ($computer_list -eq $null) {
        $string = "CN=" + $ou2find + $suffix
        $computer_list = Get-ADComputer -Filter * -SearchBase $string
    }
    $computer_list.Name >> "$env:USERPROFILE\Downloads\OU_$ou2find.log"
    Write-Host -Nonewline "Computer list saved in "
    Write-Host -ForegroundColor Green "$env:USERPROFILE\Downloads\OU_$ou2find.log"
} else {
    Write-Host -ForegroundColor Red "OU unknown"
}
pause
