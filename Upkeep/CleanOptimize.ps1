<#
Name......: CleanOptimize.ps1
Version...: 21.05.1
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

# Optimize
$answ = [System.Windows.MessageBox]::Show("Optimize Volume C:?",'OPTIMIZE','YesNo','Info')
if ($answ -eq "Yes") {
    # see https://docs.microsoft.com/en-us/powershell/module/storage/optimize-volume?view=windowsserver2019-ps
    $issd = Get-PhysicalDisk
    if ($issd.MediaType -eq 'SSD') {
        Write-Host "Found SSD drive"
        Optimize-Volume -DriveLetter C -ReTrim -Verbose
    } else {
        Write-Host "Found HDD drive"
        Optimize-Volume -DriveLetter C -Defrag -Verbose
    }  
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot now?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {
    Restart-Computer
}