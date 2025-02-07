<#
Name......: List_all_users.ps1
Version...: 23.07.1
Author....: Dario CORRADA

This script retrieve a list of all users belonging to a domain and save it in a CSV file
#>

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# check Active Directory module
if ((Get-Module -Name ActiveDirectory -ListAvailable) -eq $null) {
    $ErrorActionPreference= 'Stop'
    try {
        Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability –Online
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to install RSAT"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'
}

# retrieve user list
Write-Host "Retrieve user list..."
$user_list = Get-ADUser -Filter * -Property *
Write-Host -ForegroundColor Green "Found" $user_list.Count "user"

$rawdata = @{}
$i = 1
Write-Host -NoNewline "Retrieving..."
foreach ($auser in $user_list) {
    $user_name = $auser.Name
    
    Write-Host -NoNewline '.'

    $auser.CanonicalName -match "(.+)/[a-zA-Z_\-\.\\\s0-9:]+$" > $null
    $ou = $matches[1]

    $rawdata.$user_name = @{  
        OrganizationalUnit = $ou         
        Company = $auser.Company
        Created = $auser.Created
        Department = $auser.Department
        Description = $auser.Description
        DisplayName = $fullname
        EmailAddress = $auser.EmailAddress
        LastLogonDate = $auser.LastLogonDate
        Office = $auser.Office
        Title = $auser.Title
    }
    $i++
}
Write-Host "DONE"
$ErrorActionPreference= 'Inquire'

$outfile = "$env:USERPROFILE\Downloads\AD_Users.csv"
"UserName;OrganizationalUnit;Company;Created;Department;Description;DisplayName;EmailAddress;LastLogonDate;Office;Title" | Out-File $outfile -Encoding ASCII -Append
foreach ($usr in $rawdata.Keys) {
    $uname = $usr.ToLower()
    $new_record = @(
        $uname,
        $rawdata.$usr.OrganizationalUnit,
        $rawdata.$usr.Company,
        $rawdata.$usr.Created,
        $rawdata.$usr.Department,
        $rawdata.$usr.Description,
        $rawdata.$usr.DisplayName,
        $rawdata.$usr.EmailAddress,
        $rawdata.$usr.LastLogonDate,
        $rawdata.$usr.Office,
        $rawdata.$usr.Title
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}