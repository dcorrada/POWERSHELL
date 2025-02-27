<#
Name......: AutoReplySDK.ps1
Version...: 25.02.alfa
Author....: Dario CORRADA

This script sets an autoreply message from Outlook 365.

*** Please Note ***
It requires the Application/Delegated Permission (MailboxSettings.ReadWrite)
Check it out on Graph Explorer "Modify permissions" tab.


+++ BUGFIX NEEDED +++

    After parsing a value an unexpected character was encountered: f. Path 
    'automaticRepliesSetting.extealReplyMessage', line 3, position 128.

This bug arised once upgraded Microsoft.Graph module from 2.25 to 2.26 version.

From Microsoft Learn I haven't found any clue about putative syntax changes, 
regarding the synopsis of the cmdlet Update-MgUserMailboxSetting.

Reverse engineering approach for debugging: I could try to get and check key 
and values of the related hash table after manually autoreply setting from the 
OWA web interface. This below is the hash table obtained (edited with ****) and 
it doesn't work when setted for the script:

{
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('****')/mailboxSettings",
    "archiveFolder": "****",
    "timeZone": "W. Europe Standard Time",
    "delegateMeetingMessageDeliveryOptions": "sendToDelegateOnly",
    "dateFormat": "dd/MM/yyyy",
    "timeFormat": "HH:mm",
    "userPurpose": "user",
    "automaticRepliesSetting": {
        "status": "disabled",
        "externalAudience": "none",
        "internalReplyMessage": "<html>\n<body>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">Hi there,</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">currently I am out of office.</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">I will be available from monday to friday, 09:00-16:00. Preferably, I will reply to your email in such period.</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">Please note: my MS Teams is in sleep mode. I may not read your messages.</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">For any support request you should send an email to\n<a href=\"mailto:dario.corrada@gmail.com\" style=\"margin-top:0px; margin-bottom:0px\">\ndario.corrada@gmail.com</a></span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">Kind regards</span></p>\n</body>\n</html>\n",
        "externalReplyMessage": "<html>\n<body>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">Hi there,</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">currently I am out of office.</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">I will be available from monday to friday, 09:00-16:00. Preferably, I will reply to your email in such period.</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">Please note: my MS Teams is in sleep mode. I may not read your messages.</span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">For any support request you should send an email to\n<a href=\"mailto:dario.corrada@gmail.com\" style=\"margin-top:0px; margin-bottom:0px\">\ndario.corrada@gmail.com</a></span></p>\n<p><span style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt\">Kind regards</span></p>\n</body>\n</html>\n",
        "scheduledStartDateTime": {
            "dateTime": "2025-02-27T15:00:00.0000000",
            "timeZone": "UTC"
        },
        "scheduledEndDateTime": {
            "dateTime": "2025-02-28T15:00:00.0000000",
            "timeZone": "UTC"
        }
    },
    "language": {
        "locale": "it-IT",
        "displayName": "Italian (Italy)"
    },
    "workingHours": {
        "daysOfWeek": [
            "monday",
            "tuesday",
            "wednesday",
            "thursday",
            "friday"
        ],
        "startTime": "08:00:00.0000000",
        "endTime": "17:00:00.0000000",
        "timeZone": {
            "name": "W. Europe Standard Time"
        }
    }
}

[opened an issue at: https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3194]
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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent | Split-Path -Parent


# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
try {
    Import-Module -Name "$workdir\Modules\Forms.psm1"
    Import-Module Microsoft.Graph.Users    
} catch {
    if (!(((Get-InstalledModule).Name) -contains 'Microsoft.Graph')) {
        Install-Module Microsoft.Graph -Scope AllUsers
        [System.Windows.MessageBox]::Show("Installed [MIcrosoft.Graph] module: please restart the script",'RESTART','Ok','warning')
        exit
    } else {
        [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error')
        Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
        Pause
        exit
    }
}
$ErrorActionPreference= 'Inquire'

$splash = Connect-MgGraph -Scopes 'MailboxSettings.ReadWrite'
$UPN = (Get-MgContext).Account

<# 
# get mail message
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select Mail Message"
$OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
$OpenFileDialog.filter = 'html file (*.html)| *.html'
$OpenFileDialog.ShowDialog() | Out-Null
$InFile = $OpenFileDialog.filename
$MessageInaBottle = Get-Content -Path $InFile | Out-String
#>

# TGIF
$whatta = Get-Date
$StartTime = (Get-Date -Hour 18 -Minute 0 -Second 0 -Format "yyyy-MM-ddTHH:mm:ss.0000000").ToString()
if (($whatta.DayOfWeek -eq 'Saturday') -or ($whatta.DayOfWeek -eq 'Sunday')) {
    Write-Host -ForegroundColor Cyan "HAVE A NICE DAY!!!"
    Start-Sleep 2
    exit
} elseif ($whatta.DayOfWeek -eq 'Friday') {
    $EndTime = ((Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(3) | Get-Date -Format "yyyy-MM-ddTHH:mm:ss.0000000").ToString()
    $MessageInaBottle = @"
<html> <body> <div  style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rg b(0,0,0)">
<p>Ciao,</p>
<p>attualmente non sono in ufficio.</p>
<p>Saro' disponibile lunedi mattina prossimo.
<p>Per qualsiasi richiesta di supporto tecnico inviare una mail a <a href="mailto:helpdesk@agmsolutions.net">helpdesk@agmsolutions.net</a></p>
<p>Buon weekend</p>
<p>  </p>
<p> --- </p>
<p>  </p>
<p>Hi there,</p>
<p>currently I am out of office.</p>
<p>I will be available next monday morning.
<p>For any support request you should send an email to <a href="mailto:helpdesk@agmsolutions.net">helpdesk@agmsolutions.net</a></p>
<p>Kind regards</p>
</div> </body> </html>
"@
} else {
    $EndTime = ((Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(1) | Get-Date -Format "yyyy-MM-ddTHH:mm:ss.0000000").ToString()
    $MessageInaBottle = @"
<html> <body> <div  style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rg b(0,0,0)">
<p>Ciao,</p>
<p>attualmente non sono in ufficio.</p>
<p>Saro' disponibile da lunedi a venerdi, dalle ore 09:00 alle 18:00. Preferibilmente rispondero' alla tua mail in tale orario.</p>
<p>Nota bene: il mio account MS Teams e' attualmente configurato in sleep mode. Potrei non ricevere notifiche push.</p>
<p>Per qualsiasi richiesta di supporto tecnico inviare una mail a <a href="mailto:helpdesk@agmsolutions.net">helpdesk@agmsolutions.net</a></p>
<p>Buona serata</p>
<p>  </p>
<p> --- </p>
<p>  </p>
<p>Hi there,</p>
<p>currently I am out of office.</p>
<p>I will be available from monday to friday, 09:00-18:00. Preferably, I will reply to your email in such period.</p>
<p>Please note: my MS Teams is in sleep mode. I may not receive your push notifications.</p>
<p>For any support request you should send an email to <a href="mailto:helpdesk@agmsolutions.net">helpdesk@agmsolutions.net</a></p>
<p>Kind regards</p>
</div> </body> </html>
"@
}


# send request

<#
To set specific time zone you can list those locally stored in the registry as follows:

    $TimeZone = Get-ChildItem "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Time zones" | foreach {Get-ItemProperty $_.PSPath} 
    $TimeZone | sort Display | Format-Table -Auto PSChildname,Display

To check out the values for the setted parameters and/or add new ones, launch such an instance on Graph Explorer:

    GET https://graph.microsoft.com/v1.0/me/mailboxSettings
#>

$params = @{
	"@odata.context" = "https://graph.microsoft.com/v1.0/$UPN/mailboxSettings"
	automaticRepliesSetting = @{
		status = "Scheduled"
        externalAudience = "all"
		scheduledStartDateTime = @{
			dateTime = $StartTime
			timeZone = "W. Europe Standard Time"
		}
		scheduledEndDateTime = @{
			dateTime = $EndTime
			timeZone = "W. Europe Standard Time"
		}
        internalReplyMessage = $MessageInaBottle
        externalReplyMessage = $MessageInaBottle
	}
}
$ErrorActionPreference= 'Stop'
Write-Host -NoNewline "Setting autoreply..."
try {
    $stdout = Update-MgUserMailboxSetting -UserId $UPN -BodyParameter $params
    Write-Host -ForegroundColor Green ' DONE'
    Start-Sleep -Milliseconds 1000
}
catch {
    Write-Host -ForegroundColor Red ' FAILED'
    Write-Host "$($error[0].ToString())"
    Write-Host -ForegroundColor Blue "TIP: get <`$stdout> variable for debugging purposes"
    Pause
}
$ErrorActionPreference= 'Inquire'
