<#
Name......: SpostamentoOU.ps1
Version...: 19.04.1
Author....: Dario CORRADA

Questo script accede ad Active Directory e sposta una lista di computer da una OU specificata ad un'altra


+++ UPDATES +++

[2019-04-15 Dario CORRADA] 
Primo rilascio

[2019-04-17 Dario CORRADA] 
Si puo' evitare di specificare la OU sorgente

#>

$ErrorActionPreference= 'Inquire'

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# Controllo accesso
Import-Module -Name '\\itmilitgroup\SD_Utilities\SCRIPT\Moduli_PowerShell\Patrol.psm1'
$AD_login = Patrol -scriptname SpostamentoOU -ITuser

# Importo il modulo di Active Directory
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory } 

# inputs
$file_path = Read-Host "File lista dei PC..."
$source_ou = Read-Host "OU sorgente........."
$dest_ou = Read-Host "OU destinazione....."

if ($source_ou -eq '') {
    $source_ou = $dest_ou
}


if (Test-Path $file_path -PathType Leaf) {    
    Write-host "Recupero la lista delle OU disponibili..."
    $ou_available = Get-ADOrganizationalUnit -Filter *
    $computer_list = Get-Content $file_path # recupero la lista dei PC
    if ($ou_available.Name -contains $source_ou) {
        if ($ou_available.Name -contains $dest_ou) {
            foreach ($computer_name in $computer_list) {
                Write-Host -Nonewline $computer_name
                $computer_ADobj = Get-ADComputer $computer_name -Credential $AD_login
                # Write-Host $computer_ADobj.DistinguishedName
                if ($computer_ADobj.DistinguishedName -match $dest_ou) {
                    Write-Host -ForegroundColor Cyan " skipped"
                } elseif ($computer_ADobj.DistinguishedName -match $source_ou) {
                    $target_path = "OU=" + $dest_ou + ",DC=it,DC=kworld,DC=kpmg,DC=com"
                    $computer_ADobj | Move-ADObject -Credential $AD_login -TargetPath $target_path
                    Write-Host -ForegroundColor Green " remapped"
                } elseif ($source_ou -eq $dest_ou) {
                    $target_path = "OU=" + $dest_ou + ",DC=it,DC=kworld,DC=kpmg,DC=com"
                    $computer_ADobj | Move-ADObject -Credential $AD_login -TargetPath $target_path
                    Write-Host -ForegroundColor Green " remapped"
                } else {
                    Write-Host -ForegroundColor Cyan " skipped"
                }
            }
            pause
        } else {
            Write-Host -ForegroundColor Red "E- OU destinazione non trovata"
            pause
            Write-Host "`nLista OU disponibili:"
            $ou_available.Name | Sort-Object
            pause
            exit
        }
    } else {
        Write-Host -ForegroundColor Red "E- Ou sorgente non trovata"
        pause
        Write-Host -ForegroundColor Cyan "`nLista OU disponibili:"
        $ou_available.Name | Sort-Object
        pause
        exit
    }
} else {
    Write-Host -ForegroundColor Red "E- File non trovato"
    pause
    exit
}