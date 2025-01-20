<#
Name......: TeamsForceUninstall_ErrCode-0x80073cfa.ps1
Version...: 24.11.1
Author....: Dario CORRADA

This script force Teams removal, that wont to uninstall (error code 0x80073cfa).
Thx to Sergio Ignazio SURDO for the case study ticket.

"There was an issue in the WinAppSDK 1.6.2 release which affected Windows 10 
version 19045 machines. The 1.6.2 release has been pulled to prevent more 
machines being impacted. This issue was not caused by an update to Windows and 
will not be fixed by uninstalling any Windows Cumulative Updates. There is an 
upcoming Windows Update which will fix this issue (adding extra safety in 
Windows), and a release of WinAppSDK 1.6.3 is coming soon with a fix."
https://github.com/microsoft/WindowsAppSDK/issues/4881#issuecomment-2480939942
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
                                    CLEANING
******************************************************************************* #>
$ErrorActionPreference= 'Stop'
try {
    $Files2Del = @(
        'C:\Program Files\WindowsApps\Microsoft.WindowsAppRuntime.1.6_6000.311.13.0_x64__8wekyb3d8bbwe\WindowsAppRuntime.DeploymentExtensions.OneCore.dll', 
        'C:\Program Files\WindowsApps\Microsoft.WindowsAppRuntime.1.6_6000.311.13.0_x86__8wekyb3d8bbwe\WindowsAppRuntime.DeploymentExtensions.OneCore.dll'
    )
    $renmants = 0
    foreach ($target in $Files2Del) {
        if (Test-Path -Path $target -PathType Leaf) {
            # change ownership, grant privileges and finally delete
            $opts = ('/F', $target)
            takeown @opts
            $opts = ($target, '/grant', 'Administrators:F')
            icacls @opts
            Remove-Item -Path $target -Force
            $renmants++
        }
    }
    if ($renmants -gt 0) {
        $answ = [System.Windows.MessageBox]::Show("Renamnts found and cleaned: a seconde step is needed.`nReboot computer re-run this script...",'REBOOT','YesNo','Info')
        if ($answ -eq "Yes") {    
            Restart-Computer
        }
    } else {
        # remove WinAppRuntime 1.6
        Get-AppxPackage *WinAppRuntime.Main.1.6* -AllUsers | Where { $_.Version -eq '6000.311.13.0' } | Remove-AppxPackage -AllUsers
        Get-AppxPackage *MicrosoftCorporationII.WinAppRuntime.Singleton* -AllUsers | Where { $_.Version -eq '6000.311.13.0' } | Remove-AppxPackage -AllUsers
        Get-AppxPackage *Microsoft.WinAppRuntime.DDLM* -AllUsers | Where { $_.Version -eq '6000.311.13.0' } | Remove-AppxPackage -AllUsers
        Get-Appxpackage *WindowsAppRuntime.1.6* -AllUsers | Where { $_.Version -eq '6000.311.13.0' } | Remove-AppxPackage -AllUsers
    }
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                  HAPPY ENDING
******************************************************************************* #>

$answ = [System.Windows.MessageBox]::Show("Try to manually unistall your Teams app.`nThen click here Yes to proceed",'MANUAL UNINSTALL','YesNo','Warning')
if ($answ -eq "Yes") {    
    Restart-Computer
}

# download and re-install Teams
$DownFile = "$env:USERPROFILE\Downloads\MSTeams-x64.msix"
$download = New-Object net.webclient
$download.Downloadfile("https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix", "$DownFile")

$ErrorActionPreference= 'Stop'
try {
    Add-AppXPackage -Path $DownFile
}
catch {
    
    [System.Windows.MessageBox]::Show("ERROR - $($error[0].ToString())`n`nYou could try from PowerShell the following command:`n`nAdd-AppProvisionedPackage -Online -PackagePath (msixFile) -SkipLicense",'ERROR','Ok','Error') | Out-Null
    exit
}
$ErrorActionPreference= 'Inquire'
