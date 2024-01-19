<#
Name......: AssignedLicenses.ps1
Version...: 24.01.1
Author....: Dario CORRADA

This script will connect to Azure AD and query a list of which license(s) are assigned to each user

For more details about AzureAD cmdlets see:
https://docs.microsoft.com/en-us/powershell/module/azuread
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AzureAD\\AssignedLicenses\.ps1$" > $null
$workdir = $matches[1]
<# for testing purposes
$workdir = Get-Location
$workdir = $workdir.Path
#>

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# loading modules
Import-Module -Name "$workdir\Modules\Forms.psm1"
Import-Module -Name "$workdir\Modules\Gordian.psm1"
$ErrorActionPreference= 'Stop'
try {
    Import-Module MSOnline
} catch {
    Install-Module MSOnline -Confirm:$False -Force
    Import-Module MSOnline
}
$ErrorActionPreference= 'Inquire'

# looking for existing DB files
$dbfile = $env:LOCALAPPDATA + '\AssignedLicenses.encrypted'
$keyfile = $env:LOCALAPPDATA + '\AssignedLicenses.key'

if (Test-Path $dbfile -PathType Leaf) {
    $adialog = FormBase -w 350 -h 170 -text "DATABASE"
    if (Test-Path $keyfile -PathType Leaf) {
        $enterDB = RadioButton -form $adialog -checked $true -x 20 -y 20 -w 500 -h 30 -text "Enter the DB file"
        $cleanDB = RadioButton -form $adialog -checked $false -x 20 -y 50 -w 500 -h 30 -text "Delete the DB file"
    } else {
        $enterDB = RadioButton -form $adialog -enabled $false -checked $false -x 20 -y 20 -w 500 -h 30 -text "Enter the DB file (NO key to decrypt!)"
        $cleanDB = RadioButton -form $adialog -checked $true -x 20 -y 50 -w 500 -h 30 -text "Delete the DB file"
    }
    OKButton -form $adialog -x 100 -y 90 -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
    if ($cleanDB.Checked -eq $true) {
        $answ = [System.Windows.MessageBox]::Show("Really delete DB file?",'DELETE','YesNo','Warning')
        if ($answ -eq "Yes") {    
            Remove-Item -Path $dbfile
        }
    }
}

if (!(Test-Path $dbfile -PathType Leaf)) {
    # creating key file if not available
    if (!(Test-Path $keyfile -PathType Leaf)) {
        CreateKeyFile -keyfile "$keyfile" | Out-Null
    }

    # creating DB file
    $adialog = FormBase -w 400 -h 300 -text "DB INIT"
    Label -form $adialog -x 20 -y 20 -w 500 -h 30 -text "Initialize your DB as follows (NO space allowed)" | Out-Null
    $dbcontent = TxtBox -form $adialog -x 20 -y 50 -w 300 -h 150 -text ''
    $dbcontent.Multiline = $true;
    $dbcontent.Text = @'
USR;PWD
user1@foobar.baz;password1
user2@foobar.baz;password2
'@
    $dbcontent.AcceptsReturn = $true
    OKButton -form $adialog -x 100 -y 220 -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
    $tempusfile = $env:LOCALAPPDATA + '\AssignedLicenses.csv'
    $dbcontent.Text | Out-File $tempusfile
    EncryptFile -keyfile "$keyfile" -infile "$tempusfile" -outfile "$dbfile" | Out-Null
}

# reading DB file
$filecontent = (DecryptFile -keyfile "$keyfile" -infile "$dbfile").Split(" ")
$allowed = @{}
foreach ($newline in $filecontent) {
    if ($newline -ne 'USR;PWD') {
        ($username, $passwd) = $newline.Split(';')
        $allowed[$username] = $passwd
    }
}

# select the account to access
$adialog = FormBase -w 350 -h (($allowed.Count * 30) + 120) -text "SELECT AN ACCOUNT"
$they = 20
$choices = @()
foreach ($username in $allowed.Keys) {
    if ($they -eq 20) {
        $isfirst = $true
    } else {
        $isfirst = $false
    }
    $choices += RadioButton -form $adialog -x 20 -y $they -w 300 -checked $isfirst -text $username
    $they += 30
}
OKButton -form $adialog -x 100 -y ($they + 10) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

# get credentials for accessing
foreach ($item in $choices) {
    if ($item.Checked) {
        $usr = $item.Text
        $plain_pwd = $allowed[$usr]
    }
}
$pwd = ConvertTo-SecureString $plain_pwd -AsPlainText -Force
$credits = New-Object System.Management.Automation.PSCredential($usr, $pwd)

# Only a subset of licenses/plans of interest has been considered in this hash table.
# A complete list is available on:
# https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$managed_licenses = @{
    "ENTERPRISEPACKPLUS_FACULTY"    =   "Office 365 A3 for Faculty"
    "EXCHANGESTANDARD"              =   "Exchange Online P1"
    "EXCHANGEENTERPRISE"            =   "Exchange Online P2"
    "INTUNE_A"                      =   "Intune"
    "M365EDU_A3_FACULTY"            =   "Office 365 A3 for Students"
    "O365_BUSINESS"                 =   "Microsoft 365 Apps for Business"
    "O365_BUSINESS_ESSENTIALS"      =   "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"         =   "Microsoft 365 Business Standard"
    "PROJECTCLIENT"                 =   "Project for Office 365"
    "PROJECTESSENTIALS"             =   "Project Online Essentials"
    "PROJECTPREMIUM"                =   "Project Online Premium"
    "PROJECT_P1"                    =   "Project Plan 1"
    "PROJECTPROFESSIONAL"           =   "Project Plan 3"
    "SHAREPOINTSTORAGE"             =   "Office 365 Extra File Storage"
    "SMB_BUSINESS"                  =   "Microsoft 365 Apps for Business"
    "SMB_BUSINESS_ESSENTIALS"       =   "Microsoft 365 Business Basic"
    "SPB"                           =   "Microsoft 365 Business Premium"
    "STANDARDWOFFPACK_FACULTY"      =   "Office 365 A1 for Faculty"
    "STANDARDWOFFPACK_STUDENT"      =   "Office 365 A1 for Students"
    "Teams_Ess"                     =   "Microsoft Teams Essentials"
    "TEAMS_ESSENTIALS_AAD"          =   "Microsoft Teams Essentials"
    "TEAMS_EXPLORATORY"             =   "Microsoft Teams Exploratory"
    "VISIOCLIENT"                   =   "Visio Online Plan 2"
    "VISIO_PLAN1_DEPT"              =   "Visio Plan 1"
    "VISIO_PLAN2_DEPT"              =   "Visio Plan 2"
}

# Looking for currently distributed licenses and their availability
Clear-Host
Write-Host -ForegroundColor Yellow "STEP 00 - Available licenses"

$ErrorActionPreference= 'Stop'
Try {
    Write-Host "* Connecting to the tenant"
    Connect-MsolService -Credential $credits
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("Error accessng to the tenant",'ERROR','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

$curdist_licenses = @{}
foreach ($license in (Get-MsolAccountSku)) {
    $string = $license.SkuPartNumber
    if ($managed_licenses.ContainsKey($string)) {
        $curdist_licenses[$string] = $license.ActiveUnits - $license.ConsumedUnits
    } elseif ($license.ActiveUnits -lt 10000) {
        # alert in case you need to consider other licenses
        $answ = [System.Windows.MessageBox]::Show("Unexpected license <$string>`nUpdate `$managed_licenses before proceed?",'ABORTING','YesNo','Warning')
        if ($answ -eq "Yes") {    
            exit
        }
    }
}
$adialog = FormBase -w 400 -h ((($curdist_licenses.Count) * 30) + 120) -text "AVAILABLE LICENSES"
$they = 20
foreach ($item in ($curdist_licenses.GetEnumerator() | Sort Value)) {
    $string = $item.Name + " = " + $item.Value
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,$they)
    $label.Size = New-Object System.Drawing.Size(350,20)
    $label.Text = $string
    $adialog.Controls.Add($label)
    $they += 30
}
OKButton -form $adialog -x 75 -y ($they + 10) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()
