<#
Name......: Aggiorna_Win10.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script serve per lanciare in automatico gli aggiornamenti di Windows

+++ UPDATES +++
#>

# faccio in modo di elevare l'esecuzione dello script con privilegi di admin
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# setto le policy di esecuzione dello script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

$ErrorActionPreference= 'Stop'
try {
    Import-Module PSWindowsUpdate
} catch {
    Install-Module PSWindowsUpdate -Confirm:$False -Force
    Import-Module PSWindowsUpdate
}
$ErrorActionPreference= 'Inquire'

# elenco degli aggiornamenti disponibili
# Get-Windowsupdate

# installo gli aggiornamenti
Install-WindowsUpdate -AcceptAll -Install -Confirm:$False #| Out-File "C:\Users\$env:USERNAME\Desktop\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log" -force

# reboot
$answ = [System.Windows.MessageBox]::Show("Riavvio computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}