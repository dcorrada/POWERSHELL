<#
Questo script vuole essere solamente un proof of concept di come gestire e 
reinizailizzare la MFA di account Microsoft365 (via app MSAuthenticator, ecc. )
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
        Import-Module MSOnline
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'MSOnline')) {
            Install-Module MSOnline -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [MSOnline] module: click Ok restart the script",'RESTART','Ok','warning') > $null
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
                                QUERYING
******************************************************************************* #>
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

# qui di seguito una lista di account con licenza ma senza alcuna MFA attiva
foreach ($licensedusr in (Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true } | Sort-Object DisplayName)) {
    if ($licensedusr.StrongAuthenticationMethods.Count -le 0) {
        Write-Host "$($licensedusr.UserPrincipalName)"
    }
} 

<#
qui di seguito un chunk code di come si potrebbe implementare con MSOnline
https://techcommunity.microsoft.com/discussions/microsoft-entra/powershell-cmdlets-for-mfa-settings/157678

#Selected user in cloud
$Userpricipalname = "abc@org.com"

#Get settings for a user with exsisting auth data
$User = Get-MSolUser -UserPrincipalName $Userpricipalname
# Viewing default method
$User.StrongAuthenticationMethods

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

<# 
Qui vengono menzionati alcuni cmdlets x usare Graph alternativamente a MSOnline
https://learn.microsoft.com/en-us/entra/identity/authentication/howto-mfa-userdevicesettings
#>