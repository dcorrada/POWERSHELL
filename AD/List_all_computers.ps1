<#
Name......: List_all_computers.ps1
Version...: 19.08.1
Author....: Dario CORRADA

Questo script accede ad Active Directory ed estrae in un file CSV l'elenco di tutti computer presenti
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

# recupero la lista di tutti i computer
Write-Host "Recupero la lista di tutti i computer..."
$computer_list = Get-ADComputer -Filter * -Property *
Write-Host -ForegroundColor Green "Trovati" $computer_list.Count "computer"

$rawdata = @{}
$i = 1
foreach ($computer_name in $computer_list.Name) {
    
    Clear-Host
    Write-Host "Registrazione" $i "di" $computer_list.Count

    $infopc = Get-ADComputer -Identity $computer_name -Properties *
    $infopc.CanonicalName -match "/(.+)/$computer_name$" > $null
    $ou = $matches[1]

    $rawdata.$computer_name = @{
        OperatingSystem = $infopc.OperatingSystem
        OperatingSystemVersion = $infopc.OperatingSystemVersion
        Created = $infopc.Created
        Description = $infopc.Description
        LastLogonDate = $infopc.LastLogonDate
        OrganizationalUnit = $ou
    }

    $i++
}

$outfile = "C:\Users\$env:USERNAME\Desktop\AD_computers.csv"
"Name;OrganizationalUnit;Created;LastLogonDate;OperatingSystem;OperatingSystemVersion;Description" | Out-File $outfile -Encoding ASCII -Append
foreach ($pc in $rawdata.Keys) {
    $new_record = @(
        $pc,
        $rawdata.$pc.OrganizationalUnit,
        $rawdata.$pc.Created,
        $rawdata.$pc.LastLogonDate,
        $rawdata.$pc.OperatingSystem,
        $rawdata.$pc.OperatingSystemVersion,
        $rawdata.$pc.Description
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}