<#
Name......: Remove_PC.ps1
Version...: 22.12.1
Author....: Dario CORRADA

Remove a computer from AD
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AD\\Remove_PC\.ps1$" > $null
$repopath = $matches[1]

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$repopath\Modules\Forms.psm1"

# import Active Directory module
$ErrorActionPreference= 'Stop'
try {
    Import-Module ActiveDirectory
} catch {
    Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
    Import-Module ActiveDirectory
}
$ErrorActionPreference= 'Inquire'

# login to AD
$AD_login = LoginWindow

# dialog
$aform = FormBase -w 300 -h 150 -text 'COMPUTER'
Label -form $aform -x 10 -y 20 -w 80 -h 30 -text 'hostname:' | Out-Null
$abox = TxtBox -form $aform -x 90 -y 20 -w 150 -h 30 -text ''
OKButton -form $aform -x 100 -y 60 -text "Ok" | Out-Null
$result = $aform.ShowDialog()
$tosearch = $abox.Text

$ErrorActionPreference= 'Stop'
try {
    Get-ADComputer -Identity $tosearch -Properties *
} catch {
    [System.Windows.MessageBox]::Show("[$tosearch] not found!",'FAIL','Ok','Warning') | Out-Null
    exit
}
$ErrorActionPreference= 'Inquire'

$ErrorActionPreference= 'Stop'
Try {
    Remove-ADComputer -Identity $tosearch -Credential $AD_login -confirm:$false
    [System.Windows.MessageBox]::Show("Computer removed!",'DONE','Ok','Info') | Out-Null
}
Catch {
    [System.Windows.MessageBox]::Show("An error occurs!",'FAIL','Ok','Error') | Out-Null
    Write-Output "Error: $($error[0].ToString())`n"
    pause
}
$ErrorActionPreference= 'Inquire'
