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

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
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

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"


<# *******************************************************************************
                                    BODY
******************************************************************************* #>

if (Test-Path "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe") {

    # getting update list
    Write-Host -NoNewline "Collecting builds... "
    $HTML = Invoke-WebRequest -Uri 'https://docs.microsoft.com/en-us/officeupdates/update-history-office365-proplus-by-date' -UseBasicParsing
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
    $formlist = FormBase -w 300 -h 200 -text 'RELEASES'
    Label -form $formlist -x 10 -y 20 -text 'Select build' | Out-Null
    $buildlist = DropDown -form $formlist -x 10 -y 50 -w 250 -opts $builds
    OKButton -form $formlist -x 100 -y 100 -text "Ok" | Out-Null
    $result = $formlist.ShowDialog()
    $selected_build = $buildlist.Text

    # updating
    Write-Host -NoNewline "Updating o365... "
    if ($selected_build -ne 'LATEST UPDATE') { # rollback
        $tobuild = "16.0."+(($selected_build -split "Build ")[1] -split "\)")[0]
        $answ = [System.Windows.MessageBox]::Show("You are getting a rollback to build [$tobuild].`nAutomatic update will be disabled!`n`nDo you want to proceed?",'ROLLBACK','YesNo','Warning')
        if ($answ -eq "Yes") {    
            Start-Process powershell.exe -Verb runAs{
                Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdatesEnabled -Value "False"
            }
            & "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user updatetoversion=$tobuild
        }
    } else {
        & "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user
    }
    Write-Host -ForegroundColor Green 'DONE'
    
} else {
    [System.Windows.MessageBox]::Show("Can't find 'OfficeC2RClient.exe'`nPlease verify Office 365 is installed correctly.",'ERROR','Ok','Error') > $null
}














