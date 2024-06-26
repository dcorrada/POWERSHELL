<#
Name......: TPM_ResetKey_bugfix.ps1
Version...: 24.06.1
Author....: Dario CORRADA

This script will try to patch the error code 80090016 "Keyset does not exist".
BE AWARE that "Disabling ADAL or WAM not recommended for fixing Office sign-in 
or activation issues"*

[*] https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/administration/disabling-adal-wam-not-recommended
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Troubleshooting\\TPM_ResetKey_bugfix\.ps1$" > $null
$workdir = $matches[1]

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\O365$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"

# control panel
$adialog = FormBase -w 275 -h 190 -text "OPTIONS"
$clearTPM = RadioButton -form $adialog -x 20 -y 20 -checked $true -text 'Clear TPM Keys'
$ADALbypass = RadioButton -form $adialog -x 20 -y 50 -checked $false -text 'ADAL/WAM override'
OKButton -form $adialog -x 75 -y 90 -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

# operations
if ($clearTPM.Checked) { 
# https://learn.microsoft.com/en-us/powershell/module/trustedplatformmodule/Clear-Tpm?view=windowsserver2022-ps&viewFallbackFrom=win10-ps

    # Check if the drive is encrypted
    Write-Host -NoNewline "Checking BitLocker status... "
    $BitLockerVolumeInfo = (Get-BitLockerVolume -ErrorAction 'Stop' | Where-Object -FilterScript {$_.VolumeType -eq 'OperatingSystem'})
	if ($BitLockerVolumeInfo.ProtectionStatus -eq 'On') {
		Write-Host -ForegroundColor Yellow "ON"
        [System.Windows.MessageBox]::Show("Please, keep the BitLocker recovery key`non hand before proceeding",'WARNING','Ok','Warning') | Out-Null
        Clear-Tpm -UsePPI
	} else {
        Write-Host -ForegroundColor DarkGray "OFF"
    }

} elseif ($ADALbypass.Checked) {
# https://answers.microsoft.com/en-us/msoffice/forum/all/keyset-does-not-exist-tpm/de690cea-bba8-4260-8985-872e136e76c2

    [System.Windows.MessageBox]::Show("The following approaches are not recommended`nfor fixing Office sign-in or activation issues.`n`nClick Ok to close Outlook and Teams clients`nand then proceeding...",'WARNING','Ok','Warning') | Out-Null
    
    # killing Outlook and Teams
    $ErrorActionPreference= 'SilentlyContinue'
    $outproc = Get-Process outlook
    $teamproc = Get-Process teams
    if (($outproc -ne $null) -or ($teamproc -ne $null)) {
        $ErrorActionPreference= 'Stop'
        Try {
            $procs = @($outproc.ID, $teamproc.ID)
            foreach ($pid in $procs) {
                Stop-Process -ID $pid -Force
            }
            Start-Sleep 2
        }
        Catch { 
            [System.Windows.MessageBox]::Show("Check out that all Oulook/Teams processes`nhave been closed before go ahead",'TASK MANAGER','Ok','Warning') > $null
        }
    }
    $ErrorActionPreference= 'Inquire'

    $sdialog = FormBase -w 275 -h 200 -text "STRATEGY"
    $ADALoff = RadioButton -form $sdialog -x 20 -y 10 -checked $false -text 'Disable ADAL'
    $WAMon = RadioButton -form $sdialog -x 20 -y 40 -checked $true -text 'WAM override'
    $Basil = RadioButton -form $sdialog -x 20 -y 70 -checked $false -text 'Enable ProtectionPolicy'
    OKButton -form $sdialog -x 75 -y 120 -text "Ok" | Out-Null
    $result = $sdialog.ShowDialog()

    if ($ADALoff.Checked) {
        # Disabling ADAL, according to the original thread:
        # https://answers.microsoft.com/en-us/outlook_com/forum/all/keyset-does-not-exist-outlook-throw-an-error-if-i/4205d705-10ca-4dbf-acca-a851c45bd212
        $regpath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity'
        New-ItemProperty -Path $regpath -Name 'EnableADAL' -Value 0 -PropertyType DWord
    } elseif ($WAMon.Checked) {
        # WAM override, according to Rick Song
        $regpath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity'
        New-ItemProperty -Path $regpath -Name 'DisableADALatopWAMOverride' -Value 1 -PropertyType DWord
    } elseif ($Basil.Checked) {
        # Basilio's Shuffle
        $regpath = 'HKLM:\SOFTWARE\Microsoft\Cryptography\Protect\Providers\df9d8cd0-1501-11d1-8c7a-00c04fc297eb'
        New-ItemProperty -Path $regpath -Name 'ProtectionPolicy' -Value 1 -PropertyType DWord
    }
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}