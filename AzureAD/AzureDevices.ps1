<#
Name......: AzureDevices.ps1
Version...: 24.06.1
Author....: Dario CORRADA

This script will connect to Azure AD and query a detailed list of device properties

For more info see:
https://learn.microsoft.com/en-us/powershell/module/azuread/get-azureaddevice

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

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent
# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module AzureAD
} catch {
    Install-Module AzureAD -Confirm:$False -Force
    Import-Module AzureAD
}
$ErrorActionPreference= 'Inquire'

# connect to Tenant
<#
In that cases in which MFA has been enabled on Microsot 365 accounts the option 
"-Credential" of cmdlet "Connect-MsolService" doesn't work.
Rather such cmdlet should be used without prior specification of any credential 
(a dialog of registered account will appear, instead).
#>
# $credits = LoginWindow
$ErrorActionPreference= 'Stop'
Try {
    Connect-AzureAD #-Credential $credits
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Host -ForegroundColor Red "*** ERROR ACCESSING TENANT ***"
    # Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

# Get a list of all devices and initialize dataframe for collecting data
$rawdata = Get-AzureADDevice -All $true
$parseddata = @{}

$tot = $rawdata.Count
$i = 0
$parsebar = ProgressBar
foreach ($item in $rawdata) {
    $i++
    $ownedby = Get-AzureADDeviceRegisteredOwner -ObjectId $item.ObjectId
    $akey = $item.DisplayName
    if ($parseddata.ContainsKey($akey)) {
        $uptime = $item.ApproximateLastLogonTimeStamp | Get-Date -f yyy-MM-dd
        if ($uptime -gt $parseddata[$akey].LastLogon) {
            $parseddata[$akey].LastLogon = $uptime
            $parseddata[$akey].OSType = $item.DeviceOSType
            $parseddata[$akey].OSVersion = $item.DeviceOSVersion
            $parseddata[$akey].Owner = $ownedby.DisplayName
            $parseddata[$akey].Email = $ownedby.UserPrincipalName
        }
    } else {
        $parseddata[$akey] = @{
            'LastLogon' = $item.ApproximateLastLogonTimeStamp | Get-Date -f yyy-MM-dd
            'OSType' = $item.DeviceOSType
            'OSVersion' = $item.DeviceOSVersion
            'HostName' = $item.DisplayName
            'Owner' = $ownedby.DisplayName
            'Email' = $ownedby.UserPrincipalName        
        }
    }

    # progress
    $percent = ($i / $tot)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($i, $tot, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
$parsebar[0].Close()


# writing output file
Write-Host -NoNewline "Writing output file... "
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.Title = "Save File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'CSV file (*.csv)| *.csv'
$OpenFileDialog.filename = 'AzureDevices'
$OpenFileDialog.ShowDialog() | Out-Null
$outfile = $OpenFileDialog.filename

'UPTIME;OSTYPE;OSVERSION;HOSTNAME;OWNER;EMAIL' | Out-File $outfile -Encoding utf8

$i = 1
$totrec = $parseddata.Count
$parsebar = ProgressBar
foreach ($item in $parseddata.Keys) {
    $string = ("{0};{1};{2};{3};{4};{5}" -f ($parseddata[$item].LastLogon,$parseddata[$item].OSType,$parseddata[$item].OSVersion,$parseddata[$item].HostName,$parseddata[$item].Owner,$parseddata[$item].Email))
    $string = $string -replace ';\s*;', ';NULL;'
    $string = $string -replace ';+\s*$', ';NULL'
    $string = $string -replace ';', '";"'
    $string = '"' + $string + '"'
    $string = $string -replace '"NULL"', 'NULL'
    $string | Out-File $outfile -Encoding utf8 -Append
    $i++

    # progress
    $percent = (($i-1) / $totrec)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Writing {0} out of {1} records [{2}%]" -f (($i-1), $totrec, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()    
}
$parsebar[0].Close()
Write-Host -ForegroundColor Green "DONE"

