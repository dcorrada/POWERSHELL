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

# retrieve all users that are licensed
$Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" } | Sort-Object DisplayName

Get-MsolAccountSku







Connect-AzureAD

# initialize dataframe for collecting data
$local_array = @()

# retrieve data from AzureAD
$userlist = Get-AzureADUser -All $true
$tot = $userlist.Count
$usrcount = 0
$ErrorActionPreference= 'SilentlyContinue'
foreach ($user in $userlist) {
    $usrcount ++
    Clear-Host
    Write-Host "Processing $usrcount users out of $tot..."

    # user data
    $fullname = $user.DisplayName
    $user.Mail -match "(.+)@agmsolutions\.net$" > $null
    $username = $matches[1]

	



    Get-AzureAdUser | ForEach { $licensed=$True ; For ($i=0; $i -le ($_.AssignedLicenses | Measure).Count ; $i++) { If( [string]::IsNullOrEmpty(  $_.AssignedLicenses[$i].SkuId ) -ne $True) { $licensed=$true } } ; If( $licensed -eq $true) { Write-Host $_.UserPrincipalName} }



      
            # initialize record for collecting data
            $local_hash = [ordered]@{ 
                Fullname = $fullname;
                Username = $username;
            }

            # update dataframe
            $local_array += $local_hash



            # da valutare
            Import-Module MSOnline
            Connect-MsolService
            Get-MsolUser -All | where {$_.isLicensed -eq $true}
            Get-MsolAccountSku # fornisce gli account disponibili

            # v. https://www.thelazyadministrator.com/2018/03/19/get-friendly-license-name-for-all-users-in-office-365-using-powershell/




}
$ErrorActionPreference= 'Inquire'

# disconnect from AzureAD
Disconnect-AzureAD

# output dataframe to a CSV file
$outfile = "C:\Users\$env:USERNAME\Desktop\AzureAD.csv"
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
