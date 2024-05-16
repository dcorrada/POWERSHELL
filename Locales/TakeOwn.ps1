<#
Name......: TakeOwn.ps1
Version...: 21.05.1
Author....: Dario CORRADA

This script takes ownership of a specific path, recursively
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# header 
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# get target path
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.rootfolder = "MyComputer"
$foldername.ShowDialog()
$target = $foldername.SelectedPath

$username = $env:USERNAME

# change ownership
$opts = ("/f", $target, "/r")
takeown @opts

# grant privileges
$opts = ($target, "/grant", "$username\:f", "/c", "/t")
icacls @opts

Pause