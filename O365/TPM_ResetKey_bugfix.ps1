<#
Name......: TPM_ResetKey_bugfix.ps1
Version...: 22.12.1
Author....: Dario CORRADA

This script will try to patch the error code 80090016 "Keyset does not exist".
BE AWARE that "Disabling ADAL or WAM not recommended for fixing Office sign-in 
or activation issues"*

[*] https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/administration/disabling-adal-wam-not-recommended

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
NOTE PER ME
*   Rimuovere l'header da Teams_detach.ps1 una volta portato questo scfript sul 
    branch master
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\O365\\TPM_ResetKey_bugfix\.ps1$" > $null
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
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\O365$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"

