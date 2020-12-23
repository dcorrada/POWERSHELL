﻿<#
Name......: OSversion.ps1
Version...: 19.05.1
Author....: Dario CORRADA

Questo script accede ad Active Directory crea una lista di computer con le versioni dell'OS
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

# Importo il modulo di Active Directory
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory } 

$form = New-Object System.Windows.Forms.Form
$form.Text = "ACTIVE DIRECTORY"
$form.Size = New-Object System.Drawing.Size(550,200)
$form.StartPosition = 'CenterScreen'
$font = New-Object System.Drawing.Font("Arial", 12)
$form.Font = $font
    
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(500,30)
$label.Text = "Inserire la OU su cui cercare i computer:"
$form.Controls.Add($label)
    
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,60)
$textBox.Size = New-Object System.Drawing.Size(450,30)
$form.Controls.Add($textBox)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(100,100)
$OKButton.Size = New-Object System.Drawing.Size(75,30)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

$ou2find = $textBox.Text

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
} else {
    Write-Host -ForegroundColor Red "OU sconosciuta"
}

[System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”)
$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = “Comma separated value (*.csv)| *.csv”
$OpenFileDialog.filename = $filename
$OpenFileDialog.ShowDialog() | Out-Null
$csvfile = $OpenFileDialog.filename

"COMPUTERNAME;OPERATINGSYSTEM;VERSION" >> $csvfile
foreach ($computer_name in $computer_list.Name) {
    $infopc = Get-ADComputer -Identity $computer_name -Properties *

    $osname = $infopc.OperatingSystem
    $osversion = $infopc.OperatingSystemVersion
    "$computer_name;$osname;$osversion" >> $csvfile
}

