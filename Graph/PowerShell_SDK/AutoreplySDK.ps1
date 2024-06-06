<#
Name......: AutoReplySDK.ps1
Version...: 24.06.2
Author....: Dario CORRADA

This script sets an autoreply message from Outlook 365.

*** Please Note ***
It requires the Application/Delegated Permission (MailboxSettings.ReadWrite)
Check it out on Graph Explorer "Modify permissions" tab.
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Graph\\PowerShell_SDK\\AutoreplySDK\.ps1$" > $null
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

# get UPN
$AccessForm = FormBase -w 300 -h 200 -text 'SENDER'
Label -form $AccessForm -x 10 -y 20 -w 150 -text 'User Principal Name (UPN):' | Out-Null
$usrname = TxtBox -form $AccessForm -text 'foo@bar.baz' -x 160 -y 20 -w 120
Label -form $AccessForm -x 10 -y 60 -w 280 -text "Please Note: the UPN should be the same one`nyou will connect to Graph with" | Out-Null
OKButton -form $AccessForm -x 80 -y 110 -w 120 -text "Ok"
$resultButton = $AccessForm.ShowDialog()
$UPN = $usrname.Text

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
$StartTime = (Get-Date -Hour 18 -Minute 0 -Second 0 -Format "yyyy-MM-ddTHH:mm:ss.0000000").ToString()
if (($whatta.DayOfWeek -eq 'Saturday') -or ($whatta.DayOfWeek -eq 'Sunday')) {
    Write-Host -ForegroundColor Cyan "HAVE A NICE DAY!!!"
    Start-Sleep 2
    exit
} elseif ($whatta.DayOfWeek -eq 'Friday') {
    $EndTime = ((Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(3) | Get-Date -Format "yyyy-MM-ddTHH:mm:ss.0000000").ToString()
} else {
    $EndTime = ((Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(1) | Get-Date -Format "yyyy-MM-ddTHH:mm:ss.0000000").ToString()
}


# send request
$ConnectInfo =Connect-MgGraph -Scopes 'MailboxSettings.ReadWrite'

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
}
catch {
    Write-Host -ForegroundColor Red ' FAILED'
    Write-Host "$($error[0].ToString())"
    Write-Host -ForegroundColor Blue "TIP: get <`$stdout> variable for debugging purposes"
    Pause
}
$ErrorActionPreference= 'Inquire'
