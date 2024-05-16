<#
Name......: Trusted.ps1
Version...: 20.12.1
Author....: Dario CORRADA

This script trusts internet domains or local IPs
#>

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Safety$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"


# dialog box
$form_modalita = FormBase -w 300 -h 200 -text "NETWORK"
$internet = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Internet domain"
$intranet = RadioButton -form $form_modalita -checked $false -x 30 -y 60 -text "Local IP"
OKButton -form $form_modalita -x 90 -y 120 -text "Ok" | Out-Null
$result = $form_modalita.ShowDialog()

if ($result -eq "OK") {
    if ($internet.Checked) {
        # internet domain
        $form_EXP = FormBase -w 300 -h 175 -text "TRUSTED"
        Label -form $form_EXP -x 10 -y 20 -text 'Insert trusted domain name:' | Out-Null
        $adomain = TxtBox -form $form_EXP -x 10 -y 50 -w 250
        OKButton -form $form_EXP -x 100 -y 90 -text "Ok" | Out-Null
        $result = $form_EXP.ShowDialog()
        $DomainName = $adomain.Text       
        $prefix = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\'
        $regPath = $prefix + $DomainName
        New-Item $regPath -Force
        New-ItemProperty $regPath -Name http -Value 2 -Force
        New-ItemProperty $regPath -Name https -Value 2 -Force
    } elseif ($intranet.Checked) {
        # local IP
        $form_EXP = FormBase -w 300 -h 175 -text "TRUSTED"
        Label -form $form_EXP -x 10 -y 20 -text 'Insert trusted local IP:' | Out-Null
        $adomain = TxtBox -form $form_EXP -x 10 -y 50 -w 250
        OKButton -form $form_EXP -x 100 -y 90 -text "Ok" | Out-Null
        $result = $form_EXP.ShowDialog()
        $ipaddress = $adomain.Text
        $prefix = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\'
        $tag = -join ((65..90) | Get-Random -Count 8 | % {[char]$_})
        $regPath = $prefix + $tag
        New-Item $regPath -Force
        New-ItemProperty $regPath -Name '*' -PropertyType DWord -Value 1 -Force
        New-ItemProperty $regPath -Name ':Range' -PropertyType String -Value $ipaddress -Force
    }
}