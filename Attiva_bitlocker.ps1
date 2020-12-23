<#
Name......: Attiva_Bitlocker.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script attiva bitlocker su C: e fa un backup delle chiavi in cloud su Azure AD
#>

# faccio in modo di elevare l'esecuzione dello script con privilegi di admin
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# recupero il percorso di installazione
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Attiva_bitlocker\.ps1$" > $null
$workdir = $matches[1]

# header
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# attivo Bitlocker
$ErrorActionPreference= 'Stop'
Try {
    Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAes128 -UsedSpaceOnly -TpmProtector
    Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
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
	# Get BitLocker Volume info
	$BitLockerVolumeInfo = (Get-BitLockerVolume -ErrorAction 'Stop' | Where-Object -FilterScript {
			$_.VolumeType -eq 'OperatingSystem'
	})

	# Get the Mount Point
	$BootDrive = $BitLockerVolumeInfo.MountPoint

<#
	# Check if the drive is encrypted
	if ($BitLockerVolumeInfo.ProtectionStatus -ne 'On') {
		[System.Windows.MessageBox]::Show("BitLocker non attivato",'BitLocker','Ok','Warning')
		exit
	}
#>

	# Get the correct ID (The one from the RecoveryPassword)
	$BitLockerKeyProtectorId = ($BitLockerVolumeInfo.KeyProtector | Where-Object -FilterScript {
		$_.KeyProtectorType -eq 'RecoveryPassword'
	} | Select-Object -ExpandProperty KeyProtectorId)

	# Check if we have a recovery password/id
	if ($BitLockerKeyProtectorId) {
		# Do the backup towards AzureAD
		$null = (BackupToAAD-BitLockerKeyProtector -MountPoint $BootDrive -KeyProtectorId $BitLockerKeyProtectorId -Confirm:$false -ErrorAction 'Stop')
		[System.Windows.MessageBox]::Show("BitLocker recovery salvato su Azure AD",'BitLocker','Ok','Info')
	} else {
		[System.Windows.MessageBox]::Show("Non ci sono recovery info da salvare su Azure AD",'BitLocker','Ok','Warning')
		exit
	}
    Write-Host "Backup BitLocker effettuato" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Impossibile effettuare il backup su Azure AD",'BitLocker','Ok','Warning')
    exit
} 
