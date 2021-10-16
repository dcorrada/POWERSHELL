<#
Name......: Automigration.ps1
Version...: 21.10.1
Author....: Dario CORRADA

This script performs the migration of a local user profile from a remote machine
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Automigration\\Automigration\.ps1$" > $null
$repopath = $matches[1]

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$repopath\Modules\Forms.psm1"

# temporary folder
$tmppath = 'C:\AUTOMIGRATION'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
    New-Item -ItemType file "$tmppath\Automigration.log" > $null
}

# checking source computer
$form_IP = FormBase -w 400 -h 200 -text "SOURCE PC"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(350,30)
$label.Text = "Insert IP address of source machine:"
$form_IP.Controls.Add($label)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,60)
$textBox.Size = New-Object System.Drawing.Size(350,30)
$form_IP.Controls.Add($textBox)
OKButton -form $form_IP -x 100 -y 100 -text "Ok"
$form_IP.Add_Shown({$textBox.Select()})
$result = $form_IP.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $ipaddress = $textBox.Text
    $prefix = '\\' + $ipaddress + '\C'
}
if (Test-Path $prefix) {
    Write-Host -ForegroundColor Green "Source machine connected"
} else {
    [System.Windows.MessageBox]::Show("$prefix not available",'ERROR','Ok','Error') > $null
    Exit
}

# checking users to be migrated
$usrlist = Get-Content -Path "$prefix\AUTOMIGRATION\usrlist.log"
$usr2migrate = @()
foreach ($usr in $usrlist) {
    Write-Host -ForegroundColor Yellow -NoNewline "Checking [$usr] profile..."
    if (Get-LocalUser | Where-Object {$_.Name -eq $usr}) {
        if (net localgroup administrators | Where {$_ -eq $usr}) {
            $usr2migrate += $usr
            Write-Host -ForegroundColor Green " DONE"
        } else {
            $answ = [System.Windows.MessageBox]::Show("User [$usr] is not local admin. Grant privileges?",'USER PROFILE','YesNo','Warning')
            if ($answ -eq "Yes") {
                $ErrorActionPreference = 'Stop'
                try {
                    Add-LocalGroupMember -Group "Administrators" -Member $usr
                    $usr2migrate += $usr
                    Write-Host -ForegroundColor Red " DONE"
                }
                catch {
                    Write-Host -ForegroundColor Red " IGNORED"
                    Write-Host -ForegroundColor Red "$($error[0].ToString())"
                }
                $ErrorActionPreference = 'Inquire'
            } elseif ($answ -eq "No") {
                Write-Host -ForegroundColor Red " IGNORED"
            }
        }
    } else {
        $answ = [System.Windows.MessageBox]::Show("User [$usr] not found. Proceed?",'USER PROFILE','YesNo','Warning')
        if ($answ -eq "Yes") {
            Write-Host -ForegroundColor Red " IGNORED"
        } elseif ($answ -eq "No") {
            Exit
        }
    }
}
Clear-Host
Write-Host -ForegroundColor Yellow "The following users are able to be migrated (granted local admin privileges)"
foreach ($item in $usr2migrate) {
    Write-Host -ForegroundColor Green "[$item]"
}
Pause

# control panel
$form_panel = FormBase -w 400 -h 300 -text "STEPS"
$config_sw = CheckBox -form $form_panel -checked $false -x 20 -y 20 -text "Install standard software"
$migrazione_dati = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "Data migration"
$network_map = CheckBox -form $form_panel -checked $true -x 20 -y 80 -text "Map network paths"
# feature disabled by default
$network_map.Checked = $false
$network_map.Enabled = $false
$BitLocker_services = CheckBox -form $form_panel -checked $true -x 20 -y 110 -text "Enable BitLocker"
# feature disabled by default
$BitLocker_services.Checked = $false
$BitLocker_services.Enabled = $false
$wifi_profiles = CheckBox -form $form_panel -checked $false -x 20 -y 140 -text "WiFi profiles migration"
OKButton -form $form_panel -x 100 -y 190 -text "Ok"
$result = $form_panel.ShowDialog()

#lancio i singoli moduli selezionati dal pannello di controllo
if ($config_sw.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Installing standard software ***"

    Write-Host -NoNewline "Download software..."
    $download = New-Object net.webclient
    $download.Downloadfile("http://dl.google.com/chrome/install/375.126/chrome_installer.exe", "$tmppath\ChromeSetup.exe")
    $download.Downloadfile("http://ardownload.adobe.com/pub/adobe/reader/win/AcrobatDC/1900820071/AcroRdrDC1900820071_it_IT.exe", "$tmppath\AcroReadDC.exe")
    $download.Downloadfile("https://www.7-zip.org/a/7z1900-x64.exe", "$tmppath\7Zip.exe")
    $download.Downloadfile("https://go.skype.com/windows.desktop.download", "$tmppath\Skype.exe")
    $download.Downloadfile("https://download.ccleaner.com/spsetup132.exe", "$tmppath\Speccy.exe")
    $download.Downloadfile("https://www.revouninstaller.com/download-freeware-version.php", "$tmppath\Revo.exe")
    $download.Downloadfile("https://www.supremocontrol.com/download.aspx?file=Supremo.exe&id_sw=7&ws=supremocontrol.com", "$env:PUBLIC\Desktop\Supremo.exe")
    Write-Host -ForegroundColor Green " DONE"

    Write-Host -NoNewline "Install software..."
    Start-Process -FilePath "$tmppath\ChromeSetup.exe" -Wait
    Start-Process -FilePath "$tmppath\AcroReadDC.exe" -Wait
    Start-Process -FilePath "$tmppath\7Zip.exe" -Wait
    Start-Process -FilePath "$tmppath\Speccy.exe" -Wait
    Start-Process -FilePath "$tmppath\Revo.exe" -Wait
    $answ = [System.Windows.MessageBox]::Show("After installation close all Skype instances to proceed...",'WARNING','Ok','Warning')
    Start-Process -FilePath "$tmppath\Skype.exe" -Wait
    Write-Host -ForegroundColor Green " DONE"
}

if ($migrazione_dati.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Data migration ***"

    # *** TO DO ***
}

if ($network_map.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Mapping network paths ***"

    # *** TO DO ***
}

if ($BitLocker_services.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Enabling BitLocker ***"

    # *** TO TEST ***

    # https://docs.microsoft.com/en-us/powershell/module/bitlocker/enable-bitlocker?view=windowsserver2019-ps

    # https://www.manageengine.com/products/ad-manager/powershell/get-bitlocker-recovery-key.html
    
}

if ($wifi_profiles.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Migrating WiFi profiles ***"

    $wifi_profile_list = Get-ChildItem "$prefix\AUTOMIGRATION\wifi_profiles"
    foreach ($wifi_profile in $wifi_profile_list) {
        netsh wlan add profile filename="$prefix\AUTOMIGRATION\wifi_profiles\$wifi_profile" user=current
    }
}

# cleaning temporary
$answ = [System.Windows.MessageBox]::Show("Data migration accomplished. Delete log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "$tmppath" -Recurse -Force
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
