<#
Name......: Check_NuGet.ps1
Version...: 22.08.1
Author....: Dario CORRADA

This script check if NuGet package provider is installed, otherwise install it
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'

# looking for it
$foundit
foreach ($pp in (Get-PackageProvider)) {
    if ($pp.Name -eq 'NuGet') {
        $foundit = $pp.Name
    }
}

# take a chance
if ($foundit -ne 'NuGet') {
    $ErrorActionPreference= 'Stop'
    Try {
        Install-PackageProvider -Name "NuGet" -MinimumVersion "2.8.5.208" -Force
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }
    $ErrorActionPreference= 'Inquire'
}
