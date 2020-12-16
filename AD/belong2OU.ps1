 <#
Name......: Belong2OU.ps1
Version...: 19.04.1
Author....: Dario CORRADA

Questo script accede ad Active Directory e fornisce la lista dei computer che appartengono ad una OU
#>

# header
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\Forms.psm1'

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# Controllo accesso
$login = LoginWindow

# Importo il modulo di Active Directory
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory } 

$ou2find = Read-Host "OU in cui cercare"

Write-host "Recupero la lista delle OU disponibili..."
$ou_available = Get-ADOrganizationalUnit -Filter *

$suffix = ",DC=agm,DC=local"
if ($ou_available.Name -contains $ou2find) {
    Write-Host "Ricerca computer appartenenti a [$ou2find]"
    $string = "OU=" + $ou2find + $suffix
    $ErrorActionPreference= 'SilentlyContinue'
    $computer_list = Get-ADComputer -Filter * -SearchBase $string
    $ErrorActionPreference= 'Inquire'
    if ($computer_list -eq $null) {
        $string = "CN=" + $ou2find + $suffix
        $computer_list = Get-ADComputer -Filter * -SearchBase $string
    }
    $computer_list.Name >> "C:\Users\$env:USERNAME\Desktop\OU_$ou2find.log"
    Write-Host -Nonewline "Lista copiata in "
    Write-Host -ForegroundColor Green "C:\Users\$env:USERNAME\Desktop\OU_$ou2find.log"
} else {
    Write-Host -ForegroundColor Red "OU sconosciuta"
}
pause
