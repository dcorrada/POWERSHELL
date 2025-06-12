<#
Name......: Disable_BitLocker.ps1
Version...: 25.06.2
Author....: Dario CORRADA

This script disables BitLocker: by defaults this script expects to find a unique 
mountpoint related to volume C:
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

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

<# *******************************************************************************
                                    CHECK
******************************************************************************* #>
# searching system volume defaults
Write-Host -NoNewline "Searching volume(s)..."
foreach ($item in Get-BitLockerVolume) {
    if (($item.MountPoint -ceq 'C:') -and ($item.VolumeType -ceq 'OperatingSystem')) {
        $BLnow = $item.VolumeStatus
        $BLprot = ($item.KeyProtector).KeyProtectorType -join ','
        $BLauto = $item.AutoUnlockEnabled
    }
}
Write-Host -ForegroundColor Green " DONE"

# check status
$Proceed = 'No'
Write-Host -NoNewline "Checking status..."
if (($BLnow -ceq 'FullyEncrypted') -or ($BLnow -ceq 'EncryptionInProgress')) {
    $Proceed = [System.Windows.MessageBox]::Show("The system volume is encrypted.`nDo you wanto to proceed to disable Bitlocker?",'BitLocker','YesNo','Info')

} elseif ($BLnow -ceq 'FullyDecrypted') {
    [System.Windows.MessageBox]::Show("The system volume seems already decrypted.`nNothing to do",'BitLocker','Ok','Info') | Out-Null
} else {
    # stampo una lista degli attributi, per gestione eccezioni
    Write-Host -ForegroundColor Yellow "Unexpected output: please check-out the following settings"
    Get-BitLockerVolume | Format-List
    Pause
}
Write-Host -ForegroundColor Green " DONE"

# extras
Write-Host -NoNewline "Looking for additional settings..."
if ($BLprot -cne 'Tpm,RecoveryPassword') {
    $Proceed = [System.Windows.MessageBox]::Show("Unexpected protectors setting: [$BLprotectors]`nProceed anyway?",'BitLocker','YesNo','Warning')
}
if ($BLauto -ne $null) {
    $Proceed = [System.Windows.MessageBox]::Show("Automatic unlocking keys founr`nClear them before proceed?",'BitLocker','YesNo','Warning')
    if ($Proceed -eq 'Yes') {
         Disable-BitLockerAutoUnlock -MountPoint 'C:'
    }
}
Write-Host -ForegroundColor Green " NONE"

<# *******************************************************************************
                                    DECRYPT
******************************************************************************* #>
if ($Proceed -eq 'Yes') {
    $ErrorActionPreference= 'Stop'
    Try {
        $stdout = Disable-BitLocker -MountPoint 'C:'
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        [System.Windows.MessageBox]::Show("Disabling BitLocker Failed",'BitLocker','Ok','Warning') | Out-Null
        exit
    }

    # monitoring until decryption is full
    $parsebar = ProgressBar
    $BLremain = 100
    Write-Host -NoNewline "Decryption in progress..."
    while ($BLremain -gt 0) {
        Write-Host -NoNewline "."
        $BLremain = (Get-BitLockerVolume -MountPoint 'C:').EncryptionPercentage
        
        # progressbar
        $percent = $BLremain
        if ($percent -gt 100) {
            $percent = 100
        }
        $formattato = '{0:0.0}' -f $percent
        [int32]$progress = $percent   
        $parsebar[2].Text = ("Encryption remaining [$percent%]")
        if ($progress -ge 100) {
            $parsebar[1].Value = 100
        } else {
            $parsebar[1].Value = $progress
        }
        [System.Windows.Forms.Application]::DoEvents()

        Start-Sleep -Milliseconds 5000
    }
    Write-Host -ForegroundColor Green " DONE"
    $parsebar[0].Close()
    [System.Windows.MessageBox]::Show("Bitlocker disabled!",'BitLocker','Ok','Info') | Out-Null    
}

