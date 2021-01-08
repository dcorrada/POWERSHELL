<#
Name......: LocalAdmin-Cshared.ps1
Version...: 20.1.2
Author....: Dario CORRADA

This script prepare source PC for data migration

NOTE: in older versions of robocopy the stdout slughtly differs, causing malfunction of the script. 
In case replace all occurrences of "Byte:\s+(\d+)\s+\d+" to "Bytes :\s+(\d+)\s+\d+"
#>

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# getting the current directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\LocalAdmin-Cshared\.ps1$" > $null
$workdir = $matches[1]

# elevate script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# temporary working directory
$tmppath = 'C:\TEMPSOFTWARE'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
}

Write-Host "Looking for paths to be migrated..."
$backup_list = @{} # variable in which the paths will be listed

# retrieving paths to be migrated    
[string[]]$allow_list = Get-Content -Path "$workdir\allow_list.log"
$allow_list = $allow_list -replace ('\$username', $env:USERNAME)             
foreach ($folder in $allow_list) {
    $full_path = 'C:\' + $folder
    if (Test-Path $full_path) {
        $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Byte:\s+(\d+)\s+\d+" > $null
        $size = $Matches[1]
        if ($size -gt 1KB) {
            $backup_list[$folder] = $size
        }
    }
}

# generating list of paths to be escluded from migration
[string[]]$exclude_list = Get-Content -Path "$workdir\exclude_list.log"
$root_path = 'C:\'
$remote_root_list = Get-ChildItem $root_path -Attributes D
$elenco = @();
foreach ($folder in $remote_root_list.Name) {
    if (!($exclude_list -contains $folder)) {
        $elenco += $folder
    }
}  
foreach ($folder in $elenco) {
    if (!($folder -eq "")) {
        if (!($exclude_list -contains $folder)) {
            $full_path = $root_path + $folder
            $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $output = [system.String]::Join(" ", $output)
            $output -match "Byte:\s+(\d+)\s+\d+" > $null
            $size = $Matches[1]
            $backup_list[$folder] = $size
        }
    }
}

# add paths to be migrated downstream of C:\Users\[username]
$exclude_list = ( # exclude list may be updated according to user flavours
    "Links",
    "OneDrive",
    "DropBox",
    "Searches",
    ".config"
)
$root_path = "C:\Users\$env:USERNAME\"
$remote_root_list = Get-ChildItem $root_path -Attributes D
foreach ($folder in $remote_root_list.Name) {
    if (!($exclude_list -contains $folder)) {
        $full_path = $root_path + $folder
        $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Bytes :\s+(\d+)\s+\d+" > $null
        $size = $Matches[1]
        $full_path = "Users\$env:USERNAME\$folder"
        $backup_list[$full_path] = $size
    }
}
foreach($key in $backup_list.keys) {  
    Write-Host -NoNewline "C:\$key "
    $size = $backup_list[$key] / 1GB
    $formattato = '{0:0.0}' -f $size
    if ($size -le 5) {
        Write-Host -ForegroundColor Green "<5,0 GB"
    } elseif ($size -le 40) {
        Write-Host -ForegroundColor Yellow "$formattato GB"
    } else {
        Write-Host -ForegroundColor Red "$formattato GB"
    }
}
pause

# backup of WiFi profiles
Write-Host -NoNewline "Backup wireless network profile..."
New-Item -ItemType directory -Path "C:\TEMPSOFTWARE\wifi_profiles" > $null
netsh wlan export profile key=clear folder="C:\TEMPSOFTWARE\wifi_profiles" > $null
Write-Host -ForegroundColor Green " DONE"

# get data from customized taskbar
Copy-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Destination "C:\TEMPSOFTWARE" -Force -Recurse > $null
Reg export HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband C:\TEMPSOFTWARE\custom_pinned.reg

# grant admin privilege to the account
$ErrorActionPreference= 'SilentlyContinue'
Add-LocalGroupMember -Group "Administrators" -Member $env:USERNAME
$ErrorActionPreference= 'Inquire'
Write-Host "Local admin enabled" -fore Green

# set volume C: in share
# NOTE: the following command enable full access to everyone.
# If you wont to do so replace it with the following line:
# New-SmbShare –Name "C" –Path "C:\" -FullAccess [domain\]username
New-SmbShare –Name "C" –Path "C:\" | Grant-SmbShareAccess -AccountName Everyone -AccessRight Full -Force
Write-Host "Volume C: shared" -fore Green

New-Item -ItemType file "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" > $null

"`n*** PC NAME ***`n" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
$env:computername | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append

"`n*** IP ADDRESS ***`n" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
(Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null -and $_.NetAdapter.Status -ne "Disconnected" }).IPv4Address.IPAddress | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append

$localIP = (
    Get-NetIPConfiguration |
    Where-Object {
        $_.IPv4DefaultGateway -ne $null -and
        $_.NetAdapter.Status -ne "Disconnected"
    }
).IPv4Address.IPAddress

# network shortcut list
Write-Host "Retrieving network shortcuts"
$drives = Get-PSDrive
foreach ($mounted in $drives) {
    if ($mounted.DisplayRoot -match "^\\") {
        New-Item -ItemType file "C:\TEMPSOFTWARE\NetworkDrives.log" > $null
        $string = $mounted.Name + ":;" + $mounted.DisplayRoot
        $string | Out-File "C:\TEMPSOFTWARE\NetworkDrives.log" -Encoding ASCII -Append
    }
}

# visualizzo il file di log
notepad "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log"
