<#
Name......: Init_PC.ps1
Version...: 23.10.2
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\3rd_Parties\\Wazuh\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'

# fetch and install additional softwares
$tmppath = "C:\TEMPSOFTWARE"
New-Item -ItemType directory -Path $tmppath > $null
Write-Host -NoNewline "Download software..."
$wazuh_uri = 'https://documentation.wazuh.com/current/installation-guide/wazuh-agent/wazuh-agent-package-windows.html'
$wazuh_page = Invoke-WebRequest -Uri $wazuh_uri -UseBasicParsing
foreach ($item in $wazuh_page.Links) {
    if ($item -match 'Windows installer') {
        $DownloadPage = $item.href
    }
}
$download = New-Object net.webclient
$download.Downloadfile("$DownloadPage", "$tmppath\wazuh-agent.msi")
Write-Host -ForegroundColor Green " DONE"

Write-Host -NoNewline "Install software..."
Start-Process -FilePath "$tmppath\wazuh-agent.msi" -ArgumentList '/q','WAZUH_MANAGER="109.168.85.241"','WAZUH_REGISTRATION_SERVER="109.168.85.241"' -Wait
Write-Host -ForegroundColor Green " DONE"

Remove-Item $tmppath -Recurse -Force
