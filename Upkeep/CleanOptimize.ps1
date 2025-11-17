<#
Name......: CleanOptimize.ps1
Version...: 25.10.1
Author....: Dario CORRADA

This script runs a disk cleanup on volume C: and optimize it
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script. "
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

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework


<# *******************************************************************************
                                    BODY
******************************************************************************* #>

# Cleanup
Start-Process -FilePath CleanMgr.exe -ArgumentList '/d c: sageset:1' -Wait
Start-Process -FilePath CleanMgr.exe -ArgumentList '/sagerun:1' -Wait

# Outlook Logging,thx to Rudy Mens very very much
# https://lazyadmin.nl/it/disable-outlook-logging-and-remove-etl-files/
$answ = [System.Windows.MessageBox]::Show("Purge ETL files?",'OUTLOOK','YesNo','Info')
if ($answ -eq "Yes") {
    # Get all Users
    $Users = Get-ChildItem -Path "$($ENV:SystemDrive)\Users"

    # Process all the Users
    $Users | ForEach-Object {
        Write-Host "Processing user: $($_.Name)" -ForegroundColor Cyan
        $path = "$($ENV:SystemDrive)\Users\$($_.Name)\AppData\Local\Temp\Outlook Logging\"

        If (Test-Path $path) {
            Write-host "Removing log files from $path" -ForegroundColor Cyan
            Remove-Item -Path $path -Recurse

            <# keep this trick disabled by now
            Write-host "Creating dummy file to prevent log files" -ForegroundColor Cyan
            New-Item -Path "$($ENV:SystemDrive)\Users\$($_.Name)\AppData\Local\Temp" -Name "Outlook Logging" -ItemType File
            #>
        }
    }
}

# see https://docs.microsoft.com/en-us/powershell/module/storage/optimize-volume?view=windowsserver2019-ps
$issd = Get-PhysicalDisk

# looking for volume C:
$VolMap = Get-PhysicalDisk | ForEach-Object {
    $physicalDisk = $_
    $physicalDisk |
        Get-Disk |
        Get-Partition |
        Where-Object DriveLetter |
        Select-Object DriveLetter, @{n='SerialNumber';e={ $physicalDisk.SerialNumber}}
}
foreach ($item in $VolMap) {
    if ($item.DriveLetter -ceq 'C') {
        $theC = $item.SerialNumber
    }
}

$i = 1
foreach ($aDisk in $issd) {
    Clear-Host
    Write-Host -ForegroundColor Yellow "STORAGE DEVICE FOUND #$i"
    Write-Host -ForegroundColor Cyan @"
    * $($aDisk.Model)
    * $($aDisk.MediaType)
    * $($aDisk.HealthStatus)
    * $($aDisk.BusType)
"@
    if ($aDisk.SerialNumber -ceq $theC) {
        $answ = [System.Windows.MessageBox]::Show("Optimize Volume C:?",'OPTIMIZE','YesNo','Info')
        if ($answ -eq "Yes") {
            if ($aDisk.MediaType -eq 'SSD') {       
                Optimize-Volume -DriveLetter C -ReTrim -Verbose
            } else {
                Optimize-Volume -DriveLetter C -Defrag -Verbose
            }  
        }
    } else {
        Pause
    }
    $i++
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot now?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {
    Restart-Computer
}