<#
Name......: Init_PC.ps1
Version...: 24.06.2
Author....: Dario CORRADA

This script finalize fresh OS installations:
* install Chrome, AcrobatDC, 7Zip, Supremo;
* create a local account with admin privileges;
* set hostname according to the serial number.
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Init_PC\.ps1$" > $null
$workdir = $matches[1]

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# fetch and install additional softwares
# modify download paths according to updated software versions (updated on 2021/01/18)
$tmppath = "C:\TEMPSOFTWARE"
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath | Out-Null
}
$swlist = @{}
$form_panel = FormBase -w 350 -h 485 -text "SOFTWARES"
$swlist['Acrobat Reader DC'] = CheckBox -form $form_panel -checked $true -x 20 -y 20 -text "Acrobat Reader DC"
$swlist['Chrome'] = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "Chrome"
$swlist['TempMonitor'] = CheckBox -form $form_panel -checked $true -x 20 -y 80 -text "Open Hardware Monitor"
$swlist['Revo Uninstaller'] = CheckBox -form $form_panel -checked $true -x 20 -y 110 -text "Revo Uninstaller"
$swlist['Skype'] = CheckBox -form $form_panel -checked $false -x 20 -y 140 -text "Skype"
$swlist['Speccy'] = CheckBox -form $form_panel -checked $true -x 20 -y 170 -text "Speccy"
$swlist['Supremo'] = CheckBox -form $form_panel -checked $true -x 20 -y 200 -text "Supremo"
$swlist['Teams'] = CheckBox -form $form_panel -checked $true -x 20 -y 230 -text "Teams"
$swlist['TreeSize'] = CheckBox -form $form_panel -checked $false -x 20 -y 260 -text "TreeSize"
$swlist['WatchGuard'] = CheckBox -form $form_panel -checked $false -x 20 -y 290 -text "WatchGuard VPN"
$swlist['7ZIP'] = CheckBox -form $form_panel -checked $true -x 20 -y 320 -text "7ZIP"
OKButton -form $form_panel -x 100 -y 370 -text "Ok"  | Out-Null
$result = $form_panel.ShowDialog()

# get OS release
$info = systeminfo

$download = New-Object net.webclient
$winget_exe = Get-ChildItem -Path 'C:\Program Files\WindowsApps\' -Filter 'winget.exe' -Recurse -ErrorAction SilentlyContinue -Force
if (($info[2] -match 'Windows 10') -and ($winget_exe -eq $null)) {
    Write-Host -NoNewline "Installing Desktop Package Manager client (winget)..."
    # see also https://phoenixnap.com/kb/install-winget
    $url = 'https://github.com/microsoft/winget-cli/releases/latest'
    $request = [System.Net.WebRequest]::Create($url)
    $response = $request.GetResponse()
    $realTagUrl = $response.ResponseUri.OriginalString
    $version = $realTagUrl.split('/')[-1]
    $fileName = 'https://github.com/microsoft/winget-cli/releases/download/' + $version + '/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
    $download.Downloadfile("$fileName", "$tmppath\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle")
    Start-Process -FilePath "$tmppath\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    [System.Windows.MessageBox]::Show("Click Ok once winget will be installed...",'WAIT','Ok','Warning') > $null  
    $winget_exe = Get-ChildItem -Path 'C:\Program Files\WindowsApps\' -Filter 'winget.exe' -Recurse -ErrorAction SilentlyContinue -Force
}
$msstore_opts = '--source msstore --accept-package-agreements --accept-source-agreements --silent'
$winget_opts = '--source winget --accept-package-agreements --accept-source-agreements --silent'
Write-Host -ForegroundColor Green " DONE"
foreach ($item in ($swlist.Keys | Sort-Object)) {
    if ($swlist[$item].Checked -eq $true) {
        Write-Host -ForegroundColor Blue "[$item]"
        if ($item -eq 'Acrobat Reader DC') {
            Write-Host -NoNewline "Installing Acrobat Reader DC..."
            $StagingArgumentList = 'install  "{0}" {1}' -f 'Adobe Acrobat Reader DC (64-bit)', $winget_opts
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow
            Write-Host -ForegroundColor Green " DONE"     
        } elseif ($item -eq 'Chrome') {
            Write-Host -NoNewline "Installing Google Chrome..."
            $StagingArgumentList = 'install  "{0}" {1}' -f 'Google Chrome (EXE)', $winget_opts
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow
            Write-Host -ForegroundColor Green " DONE" 
        } elseif ($item -eq 'Revo Uninstaller') {
            Write-Host -NoNewline "Installing Revo Uninstaller..."
            $StagingArgumentList = 'install  "{0}" {1}' -f 'Revo Uninstaller', $winget_opts
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow
            Write-Host -ForegroundColor Green " DONE"   
        } elseif ($item -eq 'Skype') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://go.skype.com/windows.desktop.download", "$tmppath\Skype.exe")
            #Invoke-WebRequest -Uri 'https://go.skype.com/windows.desktop.download' -OutFile "$tmppath\Skype.exe"
            Write-Host -ForegroundColor Green " DONE"
            $answ = [System.Windows.MessageBox]::Show("After installation close all Skype instances to proceed...",'WARNING','Ok','Warning')
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\Skype.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"        
        } elseif ($item -eq 'Speccy') {
            Write-Host -NoNewline "Installing Speccy..."
            $StagingArgumentList = 'install  "{0}" {1}' -f 'Speccy', $winget_opts
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow
            Write-Host -ForegroundColor Green " DONE"   
        } elseif ($item -eq 'Supremo') {
            Write-Host -NoNewline "Download software..."
            #$download.Downloadfile("https://www.agmsolutions.net/wp-content/uploads/assistenza/Assistenza_Remota.exe", "$env:PUBLIC\Desktop\Supremo.exe")
            Invoke-WebRequest -Uri 'https://www.nanosystems.it/public/download/Supremo.exe' -OutFile "$env:PUBLIC\Desktop\Supremo.exe"
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Write-Host -ForegroundColor Green " DONE"
        } elseif ($item -eq 'Teams') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://go.microsoft.com/fwlink/p/?LinkID=869426&clcid=0x410&culture=it-it&country=IT&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64", "C:\Users\Public\Desktop\Teams_Installer.exe")
            #Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/p/?LinkID=869426&clcid=0x410&culture=it-it&country=IT&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64' -OutFile "C:\Users\Public\Desktop\Teams_Installer.exe"
            Write-Host -ForegroundColor Green " DONE"
            $answ = [System.Windows.MessageBox]::Show("Please run setup once the target account has been logged in",'INFO','Ok','Info')
        } elseif ($item -eq 'WatchGuard') {
            Write-Host -NoNewline "Download software..."
            #Invoke-WebRequest -Uri 'https://cdn.watchguard.com/SoftwareCenter/Files/MUVPN_SSL/12_7_2/WG-MVPN-SSL_12_7_2.exe' -OutFile "C:\Users\Public\Desktop\WatchGuard.exe"
            $download.Downloadfile("https://cdn.watchguard.com/SoftwareCenter/Files/MUVPN_SSL/12_7_2/WG-MVPN-SSL_12_7_2.exe", "C:\Users\Public\Desktop\WatchGuard.exe")
            Write-Host -ForegroundColor Green " DONE"
            $answ = [System.Windows.MessageBox]::Show("Please run setup once the target account has been logged in",'INFO','Ok','Info')
        } elseif ($item -eq 'TreeSize') {
            Write-Host -NoNewline "Installing TreeSize Free..."
            $StagingArgumentList = 'install  "{0}" {1}' -f 'TreeSize Free', $winget_opts
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow
            Write-Host -ForegroundColor Green " DONE" 
        } elseif ($item -eq '7ZIP') {            
            Write-Host -NoNewline "Installing 7-Zip..."
            $StagingArgumentList = 'install  "{0}" {1}' -f '7-Zip', $winget_opts
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow
            Write-Host -ForegroundColor Green " DONE"
        } elseif ($item -eq 'TempMonitor') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.6.zip", "$tmppath\openhardwaremonitor-v0.9.6.zip")
            #Invoke-WebRequest -Uri 'https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.6.zip' -OutFile "$tmppath\openhardwaremonitor-v0.9.6.zip"
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Expand-Archive "$tmppath\openhardwaremonitor-v0.9.6.zip" -DestinationPath 'C:\'
            Write-Host -ForegroundColor Green " DONE"   
        }
    }
}

# remove Skype startup
New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
$startups = Get-CimInstance Win32_StartupCommand | Select-Object Name,Location
foreach ($startup in $startups){
    if ($startup.Name -eq 'Skype for Desktop'){
        $number = ($startup.Location).IndexOf("\")
        $location = ($startup.Location).Insert("$number",":")
        Write-Output "Disabling $($startup.Name) from $location)"
        Remove-ItemProperty -Path "$location" -Name "$($startup.name)" 
    }
}

Remove-Item $tmppath -Recurse -Force


# creating local account
$answ = [System.Windows.MessageBox]::Show("Create local account?",'ACCOUNT','YesNo','Info')
if ($answ -eq "Yes") {
    
    $form = FormBase -w 350 -h 270 -text "ACCOUNT"
    Label -form $form -x 20 -y 20 -w 80 -text 'Username:' | Out-Null
    $usrname = TxtBox -form $form -x 100 -y 20 
    Label -form $form -x 20 -y 50 -w 80 -text 'Fullname:' | Out-Null
    $fullname = TxtBox -form $form -x 100 -y 50 
    $personal = RadioButton -form $form -checked $true -x 20 -y 80 -text "Set your own password"
    $apass = TxtBox -form $form -x 40 -y 110 -w 260 -masked $true
    $randomic  = RadioButton -form $form -checked $false -x 20 -y 140 -text "Generate random password"
    OKButton -form $form -x 120 -y 190 -text "Ok" | Out-Null
    $result = $form.ShowDialog()
    $username = $usrname.Text
    $completo = $fullname.Text
    if ($personal.Checked) {
        $thepasswd = $apass.Text
    } elseif ($randomic.Checked) {
        Add-Type -AssemblyName 'System.Web'
        $thepasswd = [System.Web.Security.Membership]::GeneratePassword(10, 0)
    }
    Write-Host "Username...: " -NoNewline
    Write-Host $username -ForegroundColor Cyan
    Write-Host "Password...: " -NoNewline
    Write-Host $thepasswd -ForegroundColor Cyan
    Pause
    $pwd = ConvertTo-SecureString $thepasswd -AsPlainText -Force
    
    $ErrorActionPreference= 'Stop'
    Try {
        New-LocalUser -Name $username -Password $pwd -FullName $completo -PasswordNeverExpires -AccountNeverExpires -Description "utente locale"
        Add-LocalGroupMember -Group "Administrators" -Member $username
        Write-Host -ForegroundColor Green "Local account created"
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }    
}

# changing hostname
$serial = Get-WmiObject win32_bios
$hostname = $serial.SerialNumber
$answ = [System.Windows.MessageBox]::Show("Il PC e' dislocato a Torino?",'LOCAZIONE','YesNo','Info')
if ($answ -eq "Yes") {    
    $hostname = $hostname + '-TO'
}
Write-Host "Hostname...: " -NoNewline
Write-Host $hostname -ForegroundColor Cyan

$ErrorActionPreference= 'Stop'
Try {
    Rename-Computer -NewName $hostname
    Write-Host "PC renamed" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}  

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
