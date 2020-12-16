<#
Name......: Attiva_Bitlocker.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script attiva bitlocker su C: e fa un backup delle chiavi in cloud su Azure AD
#>

# header
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# attivo Bitlocker
$ErrorActionPreference= 'Stop'
Try {
    Enable-BitLocker -MountPoint "C:" -EncryptionMethod Aes256 –UsedSpaceOnly -TpmProtector -RecoveryPasswordProtector
    Write-Host "BitLocker attivato" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Impossibile attivare BitLocker",'BitLocker','Ok','Warning')
    exit
}  

# faccio un backup su Azure AD
$ErrorActionPreference= 'Stop'
Try {
    $BLV = Get-BitLockerVolume -MountPoint "C:"
    Backup-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $BLV.KeyProtector[1].KeyProtectorId
    Write-Host "Backup BitLocker effettuato" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Impossibile effettuare il backup su Azure AD",'BitLocker','Ok','Warning')
    exit
} 