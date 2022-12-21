<#
Name......: detach_test.ps1
Version...: 22.12.1
Author....: Dario CORRADA

This testing script will check if winget would run in a detached window. 
If so, remember to patch Init_PC.ps1 in order to pause the pipe whenever an app is installing.
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\detach_test\.ps1$" > $null
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


$tmppath = "$env:USERPROFILE\Downloads"
$download = New-Object net.webclient
Write-Host -NoNewline "Installing Desktop Package Manager client (winget)..."
# see also https://phoenixnap.com/kb/install-winget
$download.Downloadfile("https://github.com/microsoft/winget-cli/releases/download/v1.3.2091/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle", "$tmppath\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle")
Start-Process -FilePath "$tmppath\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
[System.Windows.MessageBox]::Show("Click Ok once winget will be installed...",'WAIT','Ok','Warning') > $null
Write-Host -ForegroundColor Green " DONE"

Write-Host -NoNewline "Installing Acrobat Reader DC..."
winget install  "Adobe Acrobat Reader DC" --source msstore --accept-package-agreements --accept-source-agreements
Write-Host -ForegroundColor Green " DONE"

Write-Host -ForegroundColor Blue "Go beyond..."
Pause