<#
Name......: MFA_Authenticator.ps1
Version...: 25.5.1
Author....: Dario CORRADA

This script is intended to manage, initialize and restore MFA methods - usually 
mediated by MS Authenticator app - related to Microsoft 365 accounts.

Alternatively to MSOnline, the implementation adopting Graph modules should be 
available (for SDK v2.0, currently in beta):
https://learn.microsoft.com/en-us/entra/identity/authentication/howto-mfa-userdevicesettings
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script."
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
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing third party modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module MSOnline
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'MSOnline')) {
            Install-Module MSOnline -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [MSOnline] module: click Ok restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } elseif (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
            Install-Module ImportExcel -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: click Ok restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') > $null
            Write-Output "`nError: $($error[0].ToString())"
            Pause
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                DIALOG
******************************************************************************* #>
$ErrorActionPreference= 'Stop'
Write-Host -NoNewline "Connecting to MSOnLine... "
try {
    Connect-MsolService
    Write-Host -ForegroundColor Green "Ok"
}
catch {
    Write-Host -ForegroundColor Red "Ko"
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

$adialog = FormBase -w 300 -h 200 -text "SELECT AN ITEM"
$ResetOpt = RadioButton -form $adialog -x 55 -y 20 -checked $true -text "Per user re-register MFA"
$OrphanOpt = RadioButton -form $adialog -x 55 -y 50 -checked $false -text "List users without MFA"
$NextBut = OKButton -form $adialog -x 40 -y 100 -text "GO"
$AbortBut = RETRYButton -form $adialog -x 140 -y 100 -text "ABORT"
$AbortBut.DialogResult = [System.Windows.Forms.DialogResult]::ABORT
$goahead = $adialog.ShowDialog()

if ($goahead -eq 'ABORT') {
    exit
}

<# *******************************************************************************
                                ORPHANED
******************************************************************************* #>
if ($OrphanOpt.Checked) {
    Write-Host "Checking users..."
    $BimbiSperduti = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true } | ForEach-Object {
        if ($_.StrongAuthenticationMethods.Count -le 0) {        
            New-Object -TypeName PSObject -Property @{
                UPN         = "$($_.UserPrincipalName)"
                DISPLAY     = "$($_.DisplayName)"
            } | Select UPN, DISPLAY
            Write-Host -ForegroundColor Yellow "$($_.DisplayName)"
        } else {
            Write-Host -ForegroundColor DarkGray "$($_.DisplayName)"
        }
        Start-Sleep -Milliseconds 20
    }
    

    $xlsx_file = "C:$env:HOMEPATH\Downloads\MFA_Authenticator-" + (Get-Date -format "yyMMddHHmm") + '.xlsx'
    $XlsPkg = Open-ExcelPackage -Path $xlsx_file -Create

    $label = 'UPNnoMFA'
    Write-Host -NoNewline "`nWriting worksheet [$label]..."
    $XlsPkg = $BimbiSperduti | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium1' -AutoSize -PassThru
    Write-Host -ForegroundColor Green ' DONE'

    Close-ExcelPackage -ExcelPackage $XlsPkg

    [System.Windows.MessageBox]::Show("File [$xlsx_file] has been created",'OUTPUT','Ok','Info') | Out-Null
}

<# *******************************************************************************
                                RESET
******************************************************************************* #>
if ($ResetOpt.Checked) {
    $aUPNdialog = FormBase -w 320 -h 200 -text 'SELECT USER'
    $Disclaimer = Label -form $aUPNDialog -x 10 -y 10 -w 300 -h 45 -text @"
This method force user ONLY to 
re-register without clearing their 
phonenumber or App shared secret.
"@
    $Disclaimer.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Italic)
    $Disclaimer.TextAlign = 'MiddleCenter'
    $Disclaimer.ForeColor = 'Blue'
    Label -form $aUPNdialog -x 35 -y 73 -w 40 -text 'UPN:' | Out-Null
    $aUPNtxt = TxtBox -form $aUPNdialog -x 75 -y 70 -w 200
    OKButton -form $aUPNdialog -x 100 -y 120 -text "Ok" | Out-Null
    $aUPNdialog.ShowDialog() | Out-Null


    $ErrorActionPreference= 'Stop'
    try {
        $info = Get-MsolUser -UserPrincipalName "$($aUPNtxt.Text)"
        $answ = [System.Windows.MessageBox]::Show("Do you really want to proceed with `n[$($info.DisplayName)]?",'PROCEED', 'YesNo','Info')
        if ($answ -eq 'Yes') {
            Set-MsolUser -UserPrincipalName "$($aUPNtxt.Text)" -StrongAuthenticationMethods @()
        }
    }
    catch {
        Write-Host -ForegroundColor Red "`nError: $($error[0].ToString())"
        [System.Windows.MessageBox]::Show("There is something nasty with `n[$($aUPNtxt.Text)]",'Ooops!', 'Ok','Error') | Out-Null
    }
    $ErrorActionPreference= 'Inquire'
}

<# *******************************************************************************
                            PROOF OF CONCEPT
******************************************************************************* #>
<# 
The following chunk may be a template for managing MFA in more details, see also:
https://techcommunity.microsoft.com/discussions/microsoft-entra/powershell-cmdlets-for-mfa-settings/157678


#Selected user in cloud
$Userpricipalname = "abc@org.com"

#Get settings for a user with exsisting auth data
$User = Get-MSolUser -UserPrincipalName $Userpricipalname
# Viewing default method
$User.StrongAuthenticationMethods
# Getting the detail of the object related for the first method (as a template for the following custom object)
$($User.StrongAuthenticationMethods)[0] | Get-Member

# Creating custom object for default method (here you just put in $true insted of $false, on the prefeered method you like)
$m1=New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationMethod
$m1.IsDefault = $false
$m1.MethodType="OneWaySMS"

$m2=New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationMethod
$m2.IsDefault = $false
$m2.MethodType="TwoWayVoiceMobile"

$m3=New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationMethod
$m3.IsDefault = $false
$m3.MethodType="PhoneAppOTP"

$m4=New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationMethod
$m4.IsDefault = $True
$m4.MethodType="PhoneAppNotification"

# To set the users default method for doing second factor
#$m=@($m1,$m2,$m3,$m4)

# To force user ONLY to re-register without clearing their phonenumber or App shared secret.
$m=@()

# Set command to define new settings
set-msoluser -Userprincipalname $user.UserPrincipalName -StrongAuthenticationMethods $m

#Settings should be empty, and user is required to register new phone number or whatever they like, i case they lost their phone.
$User = Get-MSolUser -UserPrincipalName $Userpricipalname
$User.StrongAuthenticationMethods
#>
