<#
Name......: Trusted.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script serve per trustare domini e ip locali
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

# scelta rete
$form_modalita = FormBase -w 300 -h 200 -text "RETE"
$internet = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Dominio internet"
$intranet  = RadioButton -form $form_modalita -checked $false -x 30 -y 60 -text "IP locale"
OKButton -form $form_modalita -x 90 -y 120 -text "Ok"
$result = $form_modalita.ShowDialog()

if ($result -eq "OK") {
    if ($internet.Checked) {
        # nome dominio
        $form_EXP = FormBase -w 300 -h 200 -text "TRUSTED"
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(250,30)
        $label.Text = "Inserire il dominio del sito da aggiungere:"
        $form_EXP.Controls.Add($label)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10,50)
        $textBox.Size = New-Object System.Drawing.Size(250,30)
        $form_EXP.Controls.Add($textBox)
        OKButton -form $form_EXP -x 75 -y 90 -text "Ok"
        $form_EXP.Add_Shown({$textBox.Select()})
        $result = $form_EXP.ShowDialog()
        $DomainName = $textBox.Text

        $prefix = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\'
        $regPath = $prefix + $DomainName
        New-Item $regPath -Force
        New-ItemProperty $regPath -Name http -Value 2 -Force
        New-ItemProperty $regPath -Name https -Value 2 -Force
    } elseif ($intranet.Checked) {
        # nome dominio
        $form_EXP = FormBase -w 300 -h 200 -text "TRUSTED"
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(250,30)
        $label.Text = "Inserire l'indirizzo IP locale da aggiungere:"
        $form_EXP.Controls.Add($label)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10,50)
        $textBox.Size = New-Object System.Drawing.Size(250,30)
        $form_EXP.Controls.Add($textBox)
        OKButton -form $form_EXP -x 75 -y 90 -text "Ok"
        $form_EXP.Add_Shown({$textBox.Select()})
        $result = $form_EXP.ShowDialog()
        $ipaddress = $textBox.Text

        $prefix = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\'
        $tag = -join ((65..90) | Get-Random -Count 8 | % {[char]$_})
        $regPath = $prefix + $tag
        New-Item $regPath -Force
        New-ItemProperty $regPath -Name '*' -PropertyType DWord -Value 1 -Force
        New-ItemProperty $regPath -Name ':Range' -PropertyType String -Value $ipaddress -Force
    }
}