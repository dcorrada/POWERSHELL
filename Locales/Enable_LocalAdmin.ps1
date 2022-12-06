<#
Name......: Enable_LocalAdmin.ps1
Version...: 20.1.1
Author....: Dario CORRADA

This script grants local admin privileges to an account

see http://stackoverflow.com/questions/16617307/check-if-an-account-is-a-member-of-a-local-group-and-perform-an-if-else-in-power
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Locales\\Enable_LocalAdmin\.ps1$" > $null
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

$group = [ADSI] "WinNT://./Administrators,group"
$members = @($group.psbase.Invoke("Members"))
$AdminList = ($members | ForEach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)})
if ($AdminList -contains $env:USERNAME) {
    [System.Windows.MessageBox]::Show("Current user is already admin member",'ADMIN','Ok','Info')
} else {
    Add-LocalGroupMember -Group "Administrators" -Member $env:USERNAME
    [System.Windows.MessageBox]::Show("Current user is now admin member",'ADMIN','Ok','Info')
}
