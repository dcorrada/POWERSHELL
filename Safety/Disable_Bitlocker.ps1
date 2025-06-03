<#
Name......: Disable_BitLocker.ps1
Version...: 25.06.1
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

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

<# *******************************************************************************
                                    BODY
******************************************************************************* #>
if ((Get-BitLockerVolume).VolumeStatus -ceq 'FullyEncrypted') {
    $answ = [System.Windows.MessageBox]::Show("The system volume is encrypted.`nDo you wanto to proceed to disable Bitlocker?",'BitLocker','YesNo','Info')
    if ($answ -eq 'Yes') {
        $ErrorActionPreference= 'Stop'
        Try {
            Disable-BitLocker -MountPoint 'C:'
            $ErrorActionPreference= 'Inquire'
        }
        Catch {
            Write-Output "`nError: $($error[0].ToString())"
            [System.Windows.MessageBox]::Show("Disabling BitLocker Failed",'BitLocker','Ok','Warning') | Out-Null
            exit
        }

        # monitoring until decryption is full
        $BLremain = 100
        while ($BLremain -gt 0) {
            #Clear-Host
            Write-Host "Decryption in progress - $BLremain%"
            $BLremain = (Get-BitLockerVolume).EncryptionPercentage
            Start-Sleep -Milliseconds 5000
        }

        [System.Windows.MessageBox]::Show("Bitlocker disabled!",'BitLocker','Ok','Info') | Out-Null    
    }
} elseif ((Get-BitLockerVolume).VolumeStatus -ceq 'FullyDecrypted') {
    [System.Windows.MessageBox]::Show("The system volume seems already decrypted.`nNothing to do",'BitLocker','Ok','Info') | Out-Null
} else {
    # stampo una lista degli attributi, per gestione eccezioni
    Write-Host -ForegroundColor Yellow "Unexpected output: please check-out the following settings"
    Get-BitLockerVolume | Format-List
    Pause
}