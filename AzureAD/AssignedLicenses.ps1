<#
Name......: AssignedLicenses.ps1
Version...: 22.09.1
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

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module MSOnline
} catch {
    Install-Module MSOnline -Confirm:$False -Force
    Import-Module MSOnline
}
$ErrorActionPreference= 'Inquire'

# connect to Tenant
Connect-MsolService

# get all accounts available 
#Get-MsolAccountSku

# retrieve all users that are licensed
$Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" } | Sort-Object DisplayName

# initialize dataframe for collecting data
$parseddata = @{}

$tot = $Users.Count
$usrcount = 0
foreach ($User in $Users) {
    $usrcount ++
    Clear-Host
    Write-Host "Processing $usrcount users out of $tot..."

    $username = $User.UserPrincipalName
    $fullname = $User.DisplayName

    $licenses = (Get-MsolUser -UserPrincipalName $username).Licenses.AccountSku | Sort-Object SkuPartNumber
    if ($licenses.Count -ge 1) { # at least one license
        foreach ($license in $licenses) {
            $license = $license.SkuPartNumber
            $splitted = $fullname.Split(' ')
            $parseddata[$username] = @{
                'nome' = $splitted[0]
                'cognome' = $splitted[1]
                'email' = $username
                'licenza' = ''
                'pluslicenza' = ''
                'start' = ''
            }

            if ($license -match "O365_BUSINESS_PREMIUM") {
                $parseddata[$username].licenza += "*Standard"
            } elseif ($license -match "O365_BUSINESS_ESSENTIALS") {
                $parseddata[$username].licenza += "*Basic"
            } elseif ($license -match "EXCHANGESTANDARD") {
                $parseddata[$username].licenza += "*Exchange"   
            } else {
                $parseddata[$username].pluslicenza += "*$license"
            }
        }
    }
}

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module AzureAD
} catch {
    Install-Module AzureAD -Confirm:$False -Force
    Import-Module AzureAD
}
$ErrorActionPreference= 'Inquire'

# connect to AzureAD
Connect-AzureAD

Clear-Host
Write-Host -ForegroundColor Yellow '*** RICERCA CREAZIONE ACCOUNT ***'
Start-Sleep 2

foreach ($User in $Users) {
    $username = $User.UserPrincipalName
    $plans = (Get-AzureADUser -SearchString $username).AssignedPlans

    Write-Host -NoNewline "Looking account creation for $username... "

    foreach ($record in $plans) {
        if (($record.Service -eq 'MicrosoftOffice') -and ($record.CapabilityStatus -eq 'Enabled')){
            $started = $record.AssignedTimestamp | Get-Date -format "yyyy/MM/dd"
            if ($parseddata[$username].start -eq '') {
                $parseddata[$username].start = $started
            } elseif ($started -lt $parseddata[$username].start) {
                $parseddata[$username].start = $started
            }
        }
    }
    
    Write-Host -ForegroundColor Green 'DONE'
}

$outfile = "C:\Users\$env:USERNAME\Desktop\Licenses.csv"
Write-Host -NoNewline "`n`nWriting to $outfile... "

'NOME;COGNOME;EMAIL;DATA;LICENZA;PLUS' | Out-File $outfile -Encoding ASCII -Append

foreach ($item in $parseddata.Keys) {
    $new_record = @(
        $parseddata[$item].nome,
        $parseddata[$item].cognome,
        $parseddata[$item].email,
        $parseddata[$item].start,
        $parseddata[$item].licenza,
        $parseddata[$item].pluslicenza
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}

Write-Host -ForegroundColor Green "DONE"
Pause
