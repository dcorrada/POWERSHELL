<#
Name......: Init_PC.ps1
Version...: 23.10.2
Author....: Dario CORRADA

This script install Wazuh agent
see https://wazuh.com/
#>

# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script. "
        Pause
        Exit
    }
}

# elevated script execution with admin privileges
$ErrorActionPreference= 'Stop'
try {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($testadmin -eq $false) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
        exit $LASTEXITCODE
    }
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
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
