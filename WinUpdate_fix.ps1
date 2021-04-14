<#
Name......: WinUpdate_fix.ps1
Version...: 21.04.1
Author....: Dario CORRADA

This script tries to fix corrupted Windows updates
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\WinUpdate_fix\.ps1$" > $null
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
Import-Module -Name "$workdir\Modules\Forms.psm1"

# OU dialog box
$form_modalita = FormBase -w 300 -h 200 -text "STRATEGIES"
$cache = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Delete cache files"
$restore  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Restore last update image"
OKButton -form $form_modalita -x 90 -y 90 -text "Ok"
$result = $form_modalita.ShowDialog()

if ($result -eq 'Ok') {
    if ($cache.Checked) {
        # stopping services
        Start-Process -wait net "stop wuauserv /y"
        Start-Process -wait net "stop bits /y"
        Start-Process -wait net "stop appidsvc /y"
        Start-Process -wait net "stop cryptsvc /y"

        # backupping cache
        $dest_path = 'C:\BACKUP_WindowsUpdate_' + (Get-Date -Format "yyyy.MM.dd-HH.mm")
        New-Item -ItemType directory -Path $dest_path
        Move-Item -Path "$env:SystemRoot\SoftwareDistribution" -Destination $dest_path -Force
        Move-Item -Path "$env:SystemRoot\system32\catroot2" -Destination $dest_path -Force
        New-Item -ItemType directory -Path "$env:SystemRoot\SoftwareDistribution"
        New-Item -ItemType directory -Path "$env:SystemRoot\system32\catroot2"

        # restarting services
        Start-Process -wait net "start wuauserv"
        Start-Process -wait net "start bits"
        Start-Process -wait net "start appidsvc"
        Start-Process -wait net "start cryptsvc"  
    } elseif ($restore.Checked) {
        Start-Process -wait DISM.exe "/Online /Cleanup-image /Restorehealth"
        Start-Process -wait sfc "/scannow"
    }
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot now?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {
    Restart-Computer
}