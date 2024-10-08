<#
Name......: AutoreplyAPI.ps1
Version...: 24.06.1
Author....: Dario CORRADA

This script sets an autoreply message from Outlook 365 through RESTful Graph API. See more details on:
https://docs.microsoft.com/en-us/graph/api/user-update-mailboxsettings?view=graph-rest-1.0&tabs=http

 *** Please Note ***
It requires the Application/Delegated Permission (MailboxSettings.ReadWrite)
Check it out on Graph Explorer "Modify permissions" tab.


*** WONTFIX(?) ***
I still stuck in troubleshooting - getting 400 or 403 errors - probably due to:
https://learn.microsoft.com/en-us/graph/resolve-auth-errors#400-bad-request-or-403-forbidden-does-the-user-comply-with-their-organizations-conditional-access-ca-policies

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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent | Split-Path -Parent

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
Import-Module -Name "$workdir\Modules\Forms.psm1"
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                            CREDENTIALS MANAGEMENT
******************************************************************************* #>
# starting release 24.05.1 credentials are managed from PSWallet
Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Graph\AppKeyring.ps1"
if ($pswout.Count -eq 4) {
    $UPN = $pswout[0]
    $clientID = $pswout[1]
    $tenantID = $pswout[2]
    $Clientsecret = $pswout[3]
    Write-Host -ForegroundColor Green 'Ok'
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "Ko"
    Pause
    exit
}

# get mail message
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select Mail Message"
$OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
$OpenFileDialog.filter = 'html file (*.html)| *.html'
$OpenFileDialog.ShowDialog() | Out-Null
$InFile = $OpenFileDialog.filename
$MessageInaBottle = Get-Content -Path $InFile | Out-String

# TGIF
$whatta = Get-Date
$StartTime = (Get-Date -Hour 18 -Minute 0 -Second 0 -Format "yyyy-MM-dd HH:mm:ss").ToString()
if (($whatta.DayOfWeek -eq 'Saturday') -or ($whatta.DayOfWeek -eq 'Sunday')) {
    Write-Host -ForegroundColor Cyan "HAVE A NICE DAY!!!"
    Start-Sleep 2
    exit
} elseif ($whatta.DayOfWeek -eq 'Friday') {
    $EndTime = ((Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(3) | Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()
} else {
    $EndTime = ((Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(1) | Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()
}

#Connect to GRAPH API
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
$headers = @{
    "Authorization" = "Bearer $($tokenResponse.access_token)"
    "Content-type"  = "application/json"
}

#Set MailboxSettings
$URLSETOOF = "https://graph.microsoft.com/v1.0/users/$UPN/mailboxSettings" 
#plan
$BodyJsonSETOOF = @"
            {
                "automaticRepliesSetting": {
                    "status": "Scheduled",
                    "scheduledStartDateTime": {
                      "dateTime": "$StartTime",
                      "timeZone": "UTC"
                    },
                    "scheduledEndDateTime": {
                      "dateTime": "$EndTime",
                      "timeZone": "UTC"
                    },
                    "internalReplyMessage": "$MessageInaBottle",
                    "externalReplyMessage": "$MessageInaBottle"
                }
            }
"@

$Result = Invoke-RestMethod -Headers $headers -Body $BodyJsonSETOOF -Uri $URLSETOOF -Method PATCH

