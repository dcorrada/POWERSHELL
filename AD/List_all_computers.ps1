<#
Name......: List_all_computers.ps1
Version...: 19.08.1
Author....: Dario CORRADA

This script retrieve a list of all computers belonging to a domain and save it in a CSV file
#>

# header 
$WarningPreference = 'SilentlyContinue'

# check Active Directory module
if ((Get-Module -Name ActiveDirectory -ListAvailable) -eq $null) {
    $ErrorActionPreference= 'Stop'
    try {
        Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to install RSAT"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'
}

# retrieve computer list
Write-Host "Retrieve computer list..."
$computer_list = Get-ADComputer -Filter * -Property *
Write-Host -ForegroundColor Green "Found" $computer_list.Count "computer"

$rawdata = @{}
$i = 1
foreach ($computer_name in $computer_list.Name) {
    
    Clear-Host
    Write-Host "Retrieving" $i "of" $computer_list.Count

    $infopc = Get-ADComputer -Identity $computer_name -Properties *
    $infopc.CanonicalName -match "/(.+)/$computer_name$" > $null
    $ou = $matches[1]

    $rawdata.$computer_name = @{
        OperatingSystem = $infopc.OperatingSystem
        OperatingSystemVersion = $infopc.OperatingSystemVersion
        Created = $infopc.Created
        Description = $infopc.Description
        LastLogonDate = $infopc.LastLogonDate
        OrganizationalUnit = $ou
    }

    $i++
}

$outfile = "$env:USERPROFILE\Downloads\AD_computers.csv"
"Name;OrganizationalUnit;Created;LastLogonDate;OperatingSystem;OperatingSystemVersion;Description" | Out-File $outfile -Encoding ASCII -Append
foreach ($pc in $rawdata.Keys) {
    $new_record = @(
        $pc,
        $rawdata.$pc.OrganizationalUnit,
        $rawdata.$pc.Created,
        $rawdata.$pc.LastLogonDate,
        $rawdata.$pc.OperatingSystem,
        $rawdata.$pc.OperatingSystemVersion,
        $rawdata.$pc.Description
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}