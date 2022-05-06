<#
Name......: o365_update.ps1
Version...: 22.05.1
Author....: Dario CORRADA

Script for Update or Rollback Office 365 Client build

Based on 
https://www.powershellgallery.com/packages/Update-Office365/1.1.4

Release notes about o365 update are available at 
https://docs.microsoft.com/en-us/officeupdates/update-history-office365-proplus-by-date
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\o365_update\.ps1$" > $null
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

if (Test-Path "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe") {

    # getting update list
    Write-Host -NoNewline "Collecting builds... "
    $HTML = Invoke-WebRequest -Uri 'https://docs.microsoft.com/en-us/officeupdates/update-history-office365-proplus-by-date'
    $result = $HTML.Content 
    $current = [regex]::matches( $result, '<a href=\"(monthly-channel|current-channel)(.*?)</a>')
    $builds = @('LATEST UPDATE')
    for($i=0;$i -lt $current.count;$i++){
        $date_build = ([regex]::matches($current.value[$i],'Version \d{4} \(Build \d{4,5}\.\d{4,5}\)' )).value
        $builds += $date_build
    }
    Write-Host -ForegroundColor Green 'DONE'

    # update Content Delivery Network (CDN)
    $CDNBaseUrlCurrent = 'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlCurrent)        {
        $ChannelChanged = $true
        Start-Process powershell.exe -Verb runAs{
        Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value $CDNBaseUrlCurrent
        Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
        }
    }

    # dialog box
    # *** TODO ***
    # fare un menu a tendina in cui listare @builds ed assegnare la scelta a $selected_build

    # updating
    Write-Host -NoNewline "Updating o365... "
    if ($selected_build -ne 'LATEST UPDATE') {
        $build = "16.0."+(($date_build -split "Build ")[1] -split "\)")[0]
        & "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user updatetoversion=$build
    } else {
        & "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user
    }
    Write-Host -ForegroundColor Green 'DONE'
    
} else {
    [System.Windows.MessageBox]::Show("Can't find 'OfficeC2RClient.exe'`nPlease verify Office 365 is installed correctly.",'ERROR','Ok','Error') > $null
}














