<#
Name......: OutlookBackupProfile.ps1
Version...: 21.11.1
Author....: Dario CORRADA

This script will backup Outlook Profile into a .reg file. It works with Outlook 2016/2019/Office 365.

See also https://spinbackup.com/blog/backup-outlook-account-settings/
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\O365\\OutlookBackupProfile\.ps1$" > $null
$workdir = $matches[1]

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms')
$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Registry (*.reg)| *.reg'
$OpenFileDialog.filename = $filename
$OpenFileDialog.ShowDialog() | Out-Null
$regfile = $OpenFileDialog.filename
Reg export HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles $regfile