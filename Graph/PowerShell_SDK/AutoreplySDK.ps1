<#
Name......: AutoReplySDK.ps1
Version...: 24.06.3
Author....: Dario CORRADA

This script sets an autoreply message from Outlook 365.

*** Please Note ***
It requires the Application/Delegated Permission (MailboxSettings.ReadWrite)
Check it out on Graph Explorer "Modify permissions" tab.
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

# +++ BUGFIX NEEDED +++
$TheDialog = FormBase -w 580 -h 120 -text "ABORTING"
$Disclaimer = Label -form $TheDialog -x 25 -y 15 -w 150 -h 45 -text 'BUGFIX NEEDED'
$Disclaimer.Font = [System.Drawing.Font]::new("Arial", 12, [System.Drawing.FontStyle]::Bold)
$Disclaimer.TextAlign = 'MiddleCenter'
$Disclaimer.BackColor = 'Red'
$Disclaimer.ForeColor = 'Yellow'
$ExLinkLabel = New-Object System.Windows.Forms.LinkLabel
$ExLinkLabel.Location = New-Object System.Drawing.Size(185,25)
$ExLinkLabel.Size = New-Object System.Drawing.Size(450,65)
$ExLinkLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$ExLinkLabel.Text = @"
2025/02/27 - Unexpected exception arised working with 
Microsoft Graph module. See more detail by click here.
"@
$ExLinkLabel.add_Click({[system.Diagnostics.Process]::start("https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3194")})
$TheDialog.Controls.Add($ExLinkLabel)
$TheDialog.ShowDialog() | Out-Null
Exit

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
