<#
Name......: Metti_a_dominio.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script mette a dominio un PC
#>

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
Import-Module -Name "$workdir\Moduli_PowerShell\Forms.psm1"

$hostname = $env:computername
$dominio = 'agm.local'

# Recupero le credenziali AD
$ad_login = LoginWindow

# form di scelta OU
$form_modalita = FormBase -w 300 -h 230 -text "OU DESTINAZIONE"
$noou = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "null"
$consulenti  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Client Consulenti"
$milano = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "Client Milano"
$torino = RadioButton -form $form_modalita -checked $false -x 30 -y 110 -text "Client Torino"
OKButton -form $form_modalita -x 90 -y 150 -text "Ok"
$result = $form_modalita.ShowDialog()

if ($result -eq "OK") {
    if ($noou.Checked) {
        $outarget = "null"
    } elseif ($consulenti.Checked) {
        $outarget = 'OU=Client Consulenti,OU=Delegate,DC=agm,DC=local'
    } elseif ($milano.Checked) {
        $outarget = 'OU=Client Milano,OU=Delegate,DC=agm,DC=local'
    } elseif ($torino.Checked) {
        $outarget = 'OU=Client Torino,OU=Delegate,DC=agm,DC=local'
    }    
}

Write-Host "Domain...: " -NoNewline
Write-Host $dominio -ForegroundColor Cyan
Write-Host "OU.......: " -NoNewline
Write-Host $outarget -ForegroundColor Cyan

$ErrorActionPreference= 'Stop'
Try {
    if ($outarget -eq "null") {
        Add-Computer -ComputerName $hostname -Credential $ad_login -DomainName $dominio -Force
    } else {
        Add-Computer -ComputerName $hostname -Credential $ad_login -DomainName $dominio -OUPath $outarget -Force
    }
    Write-Host "PC messo a dominio" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
} 

# reboot
$answ = [System.Windows.MessageBox]::Show("Riavvio computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}