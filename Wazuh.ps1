<#
Name......: Init_PC.ps1
Version...: 20.12.3
Author....: Dario CORRADA

This script install Wazuh agent
see https://wazuh.com/
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Wazuh\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'

# fetch and install additional softwares
# modify download paths according to updated software versions (updated on 2021/01/18)
$tmppath = "C:\TEMPSOFTWARE"
New-Item -ItemType directory -Path $tmppath > $null
Write-Host -NoNewline "Download software..."
$download = New-Object net.webclient
$download.Downloadfile("https://packages.wazuh.com/4.x/windows/wazuh-agent-4.1.5-1.msi", "$tmppath\wazuh-agent.msi")
Write-Host -ForegroundColor Green " DONE"

Write-Host -NoNewline "Install software..."
Start-Process -FilePath "$tmppath\wazuh-agent.msi" -ArgumentList '/q','WAZUH_MANAGER="109.168.85.241"','WAZUH_REGISTRATION_SERVER="109.168.85.241"' -Wait
Write-Host -ForegroundColor Green " DONE"

Remove-Item $tmppath -Recurse -Force
