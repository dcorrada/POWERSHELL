<#
Name......: Enable_BitLocker.ps1
Version...: 20.12.1
Author....: Dario CORRADA

This script enables BitLocker onto volume C: and backups recovery keys on Azure AD cloud
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get the working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Safety\\Enable_BitLocker\.ps1$" > $null
$workdir = $matches[1]

# header
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# enable Bitlocker
$ErrorActionPreference= 'Stop'
Try {
    Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAes128 -UsedSpaceOnly -TpmProtector
    Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
    Write-Host "BitLocker enabled" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Enabling BitLocker Failed",'BitLocker','Ok','Warning')
    exit
}  

# backup on Azure AD
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
		[System.Windows.MessageBox]::Show("BitLocker disabled",'BitLocker','Ok','Warning')
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
		[System.Windows.MessageBox]::Show("Recovery key saved on Azure AD",'BitLocker','Ok','Info')
	} else {
		[System.Windows.MessageBox]::Show("No recovery key to save",'BitLocker','Ok','Warning')
		exit
	}
    Write-Host "Recovery key saved on Azure AD" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Saving recovery key failed",'BitLocker','Ok','Warning')
    exit
} 
