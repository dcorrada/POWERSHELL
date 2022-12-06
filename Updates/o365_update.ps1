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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Updates\\o365_update\.ps1$" > $null
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
    $formlist = FormBase -w 300 -h 200 -text $item
    $DropDown = new-object System.Windows.Forms.ComboBox
    $DropDown.Location = new-object System.Drawing.Size(10,50)
    $DropDown.Size = new-object System.Drawing.Size(250,30)
    $DropDown.Text = $builds[0]
    foreach ($elem in $builds) {
        if ($elem) {
            $DropDown.Items.Add($elem)  > $null
        }
    }
    $formlist.Controls.Add($DropDown)
    $DropDownLabel = new-object System.Windows.Forms.Label
    $DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
    $DropDownLabel.size = new-object System.Drawing.Size(250,30) 
    $DropDownLabel.Text = "Select build"
    $formlist.Controls.Add($DropDownLabel)
    OKButton -form $formlist -x 100 -y 100 -text "Ok"
    $formlist.Add_Shown({$DropDown.Select()})
    $result = $formlist.ShowDialog()
    $selected_build = $DropDown.Text

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














