<#
Name......: AssignedLicenses.ps1
Version...: 21.08.1
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
$local_array = @()

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
        $licenselist = $licenses[0].SkuPartNumber
        if ($licenses.Count -gt 1) {
            for ($i = 1; $i -lt $licenses.Count; $i++) {
                $licenselist = $licenselist + ':' + $licenses[$i].SkuPartNumber
            }
        }

        # initialize record for collecting data
        $local_hash = [ordered]@{ 
            Fullname = $fullname;
            Username = $username;
            Licenses = $licenselist
        }

        # update dataframe
        $local_array += $local_hash
    }
}

# output dataframe to a CSV file
$outfile = "C:\Users\$env:USERNAME\Desktop\Licenses.csv"
Write-Host -NoNewline "Writing to $outfile... "

$header = @($local_array[0].Keys)
$new_string = [system.String]::Join(";", $header)
$new_string | Out-File $outfile -Encoding ASCII -Append

foreach ($item in $local_array) {
    $record = @($item.Values)
    $new_string = [system.String]::Join(";", $record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}

Write-Host "DONE"
