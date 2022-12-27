<#
Name......: Autoreply.ps1
Version...: 22.12.1
Author....: Dario CORRADA

This script sets an autoreply message in Outlook. In the following example I will set an autoreply from 04:00pm to 09:00am of the day after. 
see also https://superuser.com/questions/1683334/scheduled-autoreply/1683591#1683591 
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\O365\\Autoreply\.ps1$" > $null
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

# import the EXO V2 module
$ErrorActionPreference= 'Stop'
try {
    Import-Module ExchangeOnlineManagement
} catch {
    Install-Module ExchangeOnlineManagement -Confirm:$False -Force
    Import-Module ExchangeOnlineManagement
}
$ErrorActionPreference= 'Inquire'

# closing Outlook
$answ = [System.Windows.MessageBox]::Show("Click Ok to close Outlook client...",'WARNING','Ok','Warning')
$ErrorActionPreference= 'SilentlyContinue'
$outproc = Get-Process outlook
if ($outproc -ne $null) {
    $ErrorActionPreference= 'Stop'
    Try {
        Stop-Process -ID $outproc.Id -Force
        Start-Sleep 2
    }
    Catch { 
        [System.Windows.MessageBox]::Show("Check out that all Oulook processes have been closed before go ahead",'TASK MANAGER','Ok','Warning') > $null
    }
}
$ErrorActionPreference= 'Inquire'

# get credentials
$form = FormBase -w 520 -h 300 -text "ACCOUNT"
$labelusr = Label -form $form -x 10 -y 20 -w 500 -h 30 -text 'Username:'
$usrname = TxtBox -form $form -x 10 -y 60 -w 450 -h 30 -text ''
$labelpwd = Label -form $form -x 10 -y 100 -w 500 -h 30 -text 'Password:'
$passwd = TxtBox -form $form -x 10 -y 140 -w 450 -h 30 -text ''
$passwd.PasswordChar = '*'
OKButton -form $form -x 200 -y 190 -text "Ok" | Out-Null
$form.Topmost = $true
$result = $form.ShowDialog()

<#
# If you need to set specific time zone you can list those locally stored in the registry as follows
$TimeZone = Get-ChildItem "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Time zones" | foreach {Get-ItemProperty $_.PSPath}
$TimeZone | sort Display | Format-Table -Auto PSChildname,Display
#>
Set-MailboxRegionalConfiguration -TimeZone 'UTC' -Identity $unique

# setting autoreply
$username = $usrname.Text
$username -match "^([a-zA-Z_\-\.\\\s0-9:]+)@.+$" > $null
$unique = $matches[1]
$password = ConvertTo-SecureString $passwd.Text -AsPlainText -Force
$UserCredential = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)
Connect-ExchangeOnline -Credential $UserCredential
$message = @'
<html> <body> <div  style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rg b(0,0,0)">
<p>Hi there,</p>
<p>currently I am out of office.</p>
<p>I will be available from monday to friday, 09:00-16:00. Preferably, I will reply to your email in such period.</p>
<p>Please note: my MS Teams is in sleep mode. I may not read your messages.</p>
<p>For any support request you should send an email to <a href="mailto:dario.corrada@gmail.com">dario.corrada@gmail.com</a></p>
<p>Kind regards</p>
</div> </body> </html>
'@
Set-MailboxAutoReplyConfiguration `
    -Identity $unique `
    -AutoReplyState "Scheduled" `
    -ExternalMessage $message `
    -InternalMessage $message `
    -StartTime (Get-Date -Hour 16 -Minute 0 -Second 0) `
    -EndTime (((Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(1))) `
    -ExternalAudience All

# restart Outlook
$answ = [System.Windows.MessageBox]::Show("Restart Outlook client?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Start-Process outlook
}
