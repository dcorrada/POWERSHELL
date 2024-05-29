<#
Name......: AutoReply.ps1
Version...: 24.05.1
Author....: Dario CORRADA

This script sets an autoreply message from Outlook 365. See more details on:
https://docs.microsoft.com/en-us/graph/api/user-update-mailboxsettings?view=graph-rest-1.0&tabs=http

 *** Please Note ***
 It requires the Application/Delegated Permission (MailboxSettings.ReadWrite) 
 from Azure App Registration.
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Graph\\API\\AutoReply\.ps1$" > $null
$workdir = $matches[1]
<# for testing purposes
$workdir = Get-Location
$workdir = $workdir.Path
#>

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
Import-Module -Name "$workdir\Modules\Forms.psm1"
$ErrorActionPreference= 'Inquire'

<#
Here there is a template of script retrieved from:
https://www.techguy.at/set-out-of-office-reply-with-powershell-and-ms-graph-api/



$clientID = "your Client ID"
$Clientsecret = "Your Secret"
$tenantID = "your Tenant"


$UPN = "first.last@techguy.at"

$HTMLintern=@"
<html>\n<body>\n<p>I'm at our company's worldwide reunion and will respond to your message as soon as I return.<br>\n</p></body>\n</html>\n
"@

$HTMLextern=@"
<html>\n<body>\n<p>I'm at the Contoso worldwide reunion and will respond to your message as soon as I return.<br>\n</p></body>\n</html>\n
"@


#Function



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
                      "dateTime": " 2020-08-25 12:00:00",
                      "timeZone": "UTC"
                    },
                    "scheduledEndDateTime": {
                      "dateTime": " 2021-08-25 12:00:00",
                      "timeZone": "UTC"
                    },
                    "internalReplyMessage": "$HTMLintern",
                    "externalReplyMessage": "$HTMLextern"
                }
            }
"@
#immediately
$BodyJsonSETOOF = @"
            {
                
                "automaticRepliesSetting": {
                    "status": "alwaysEnabled",
                    "internalReplyMessage": "$HTMLintern",
                    "externalReplyMessage": "$HTMLextern"
                }
            }
"@
    
$Result = Invoke-RestMethod -Headers $headers -Body $BodyJsonSETOOF -Uri $URLSETOOF -Method PATCH
#>