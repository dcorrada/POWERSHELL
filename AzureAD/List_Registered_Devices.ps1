<#
Name......: List_Registered_Devices.ps1
Version...: 21.03.1
Author....: Dario CORRADA

This script will connect to Azure AD and query a list of the registered devices for each user

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

    # user data (edit your domain name)
    $fullname = $user.DisplayName
    $user.Mail -match "(.+)@([a-zA-Z0-9:]+)\.([a-zA-Z]+)$" > $null
    $username = $matches[1]

    # check device(s) assigned to user
    $User_ObjectID = $user.ObjectID
    if ($User_ObjectID -ne $null) {
        $User_Devices = (Get-AzureADUserRegisteredDevice -ObjectId $User_ObjectID)
        if ($User_Devices -ne $null) {
            $Count_User_Devices = $User_Devices.Count
        }
    }
	
    # device data
    if ($Count_User_Devices -ne $null) {
        foreach ($device in $User_Devices) {
            
            # initialize record for collecting data
            $local_hash = [ordered]@{ 
                Fullname = $fullname;
                Username = $username;
                Devicename = $device.DisplayName;
                Lastlogon = $device.ApproximateLastLogonTimeStamp.ToString("yyyy-MM-dd HH:mm:ss");
                Osname = $device.DeviceOSType;
                Osversion = $device.DeviceOSVersion
            }

            # update dataframe
            $local_array += $local_hash
        }
    }
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
