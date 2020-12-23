<#
Name......: Metti_a_dominio.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script mette a dominio un PC
#>

# faccio in modo di elevare l'esecuzione dello script con privilegi di admin
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# recupero il percorso di installazione
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Metti_a_dominio\.ps1$" > $null
$workdir = $matches[1]

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

# recupero il nome del dominio
$output = nslookup ls
$output[0] -match "Server:\s+[a-zA-Z_\-0-9]+\.([a-zA-Z\-0-9\.]+)$" > $null
$dominio = $matches[1]
$answ = [System.Windows.MessageBox]::Show("Regitrarsi sul dominio [$dominio]?",'DOMAIN','YesNo','Info')
if ($answ -eq "No") {    
    $form = FormBase -w 520 -h 220 -text "DOMAIN"
    $font = New-Object System.Drawing.Font("Arial", 12)
    $form.Font = $font
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(500,30)
    $label.Text = "Nome dominio:"
    $form.Controls.Add($label)
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,60)
    $textBox.Size = New-Object System.Drawing.Size(450,30)
    $form.Controls.Add($textBox)
    $OKButton = New-Object System.Windows.Forms.Button
    OKButton -form $form -x 200 -y 110 -text "Ok"
    $form.Topmost = $true
    $result = $form.ShowDialog()
    $dominio = $textBox.Text
}


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