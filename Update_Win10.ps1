<#
Name......: Update_Win10.ps1
Version...: 20.12.1
Author....: Dario CORRADA

This script automatically fetch and install Windows updates
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# looking for NuGet Package Provider
try {
    $pp = Get-PackageProvider -Name NuGet
}
catch {
    Install-PackageProvider -Name NuGet -Force
}

$ErrorActionPreference= 'Stop'
try {
    Import-Module PSWindowsUpdate
} catch {
    Install-Module PSWindowsUpdate -Confirm:$False -Force
    Import-Module PSWindowsUpdate
}
$ErrorActionPreference= 'Inquire'

# list of available updates
# Get-Windowsupdate

# install the updates
Install-WindowsUpdate -AcceptAll -Install -IgnoreReboot -Confirm:$False | Out-File "C:\Users\$env:USERNAME\Desktop\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log" -force

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
