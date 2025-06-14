<#
Name......: Init_PC.ps1
Version...: 25.06.1
Author....: Dario CORRADA

This script finalize fresh OS installations:
* install Chrome, AcrobatDC, 7Zip, Supremo, Microsoft 365 apps;
* create a local account with admin privileges;
* set hostname according to the serial number.

+++ KNOWN BUGS +++
* Teams wont be installed machinewide from winget. As temporary workaround a 
  msix installer will be manually downloaded and asked install it afterwards.
  Such bug doesn't affect Windows 11 installation, since Teams app is already 
  onboard.
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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent 

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

<# *******************************************************************************
                                    BODY
******************************************************************************* #>
# fetch and install additional softwares
# modify download paths according to updated software versions (updated on 2021/01/18)

$tmppath = "C:\TEMPSOFTWARE"
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath | Out-Null
}
# get OS release
$info = systeminfo

$download = New-Object net.webclient
$winget_exe = $null
if ($info[2] -match 'Windows 10') {
    # see also https://phoenixnap.com/kb/install-winget
    Write-Host -NoNewline "Installing Desktop Package Manager client (winget)..."
    $url = 'https://github.com/microsoft/winget-cli/releases/latest'
    $ErrorActionPreference= 'Stop'
    try {
        $request = [System.Net.WebRequest]::Create($url)
        $response = $request.GetResponse()
        $realTagUrl = $response.ResponseUri.OriginalString
        $version = $realTagUrl.split('/')[-1]
        $fileName = 'https://github.com/microsoft/winget-cli/releases/download/' + $version + '/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
        $download.Downloadfile("$fileName", "$tmppath\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle")
        Start-Process -FilePath "$tmppath\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
        [System.Windows.MessageBox]::Show("Click Ok once winget will be installed...",'WAIT','Ok','Warning') > $null  
        $winget_exe = Get-ChildItem -Path 'C:\Program Files\WindowsApps\' -Filter 'winget.exe' -Recurse -ErrorAction SilentlyContinue -Force
        Write-Host -ForegroundColor Green " DONE"
    }
    catch {
        Write-Host -ForegroundColor Red " FAILED"    
        Write-Output "Error: $($error[0].ToString())"
    }
    $ErrorActionPreference= 'Inquire'
} elseif ($info[2] -match 'Windows 11') {
    # Here below you can find the thread I opened on such topic
    # https://superuser.com/questions/1858012/winget-wont-upgrade-on-windows-11

    $stdout_winget = winget.exe source update
    $source = ('null', 'null')
    foreach ($newline in $stdout_winget) {
        if (($newline -match "^Aggiornamento origine") -or ($newline -match "^Updating all sources")) {
            $newline -match ": ([A-Za-z0-9]+)(\.{0,3})$" | Out-Null
            $source[0] = $matches[1]
        } elseif (($newline -eq "Fatto") -or ($newline -eq "Done")) {
            $source[1] = 'Ok'
        } elseif (($newline -eq "Annullato") -or ($newline -eq "Cancelled")) {
            $source[1] = 'Ko'
        }
    }

    if ($source[1] -eq 'Ko') {
        Write-Host -ForegroundColor Red "Failed to update [$($source[0])] source"
<# UPDATE April the 14th, 2025
Such issue mainly involve "winget" repository (currently no events collected 
for "msstore"). This bug has been initilally reported on StackExchange[1], for
what concern hosts in which Windows 11 23H2 has been installed.

Recently, such issue raised again onto Windows 11 24H2 installations. Another 
possible solution was found on GitHub issue thread[2]:

    Start a shell as admin:
    Find-PackageProvider -Name NuGet -ForceBootstrap -IncludeDependencies -Force
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-Module -Name PSDownloader -Scope AllUsers
    Import-Module PSDownloader
    Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe
    $ProgressPreference = "SilentlyContinue"
    Start-Download -Url "https://cdn.winget.microsoft.com/cache/source.msix" -Threads 8 -Force -MaxRetry 3 -Destination ".\source.msix" -NoProgress
    Add-AppxPackage -Path ".\source.msix" -ForceApplicationShutdown -ForceUpdateFromAnyVersion

    Then exit, start a new shell as admin:
    winget source update
    winget list --accept-source-agreements


[1] https://superuser.com/questions/1858012/winget-wont-upgrade-on-windows-11/1891855
[2] https://github.com/microsoft/winget-cli/issues/5366
#>
    } else {
        $winget_exe = Get-ChildItem -Path 'C:\Program Files\WindowsApps\' -Filter 'winget.exe' -Recurse -ErrorAction SilentlyContinue -Force
    }
}
if ([string]::IsNullOrEmpty($winget_exe)) {
    [System.Windows.MessageBox]::Show("Winget not configured, some app will not available.`nProceed with manually download and installation for them.",'WINGET','Ok','Warning') | Out-Null
}

$swlist = @{}
$form_panel = FormBase -w 350 -h 470 -text "SOFTWARES"
$swlist['Acrobat Reader'] = CheckBox -form $form_panel -checked $true -x 20 -y 20 -text "Acrobat Reader"
$swlist['BatteryMon'] = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "BatteryMon"
$swlist['Chrome'] = CheckBox -form $form_panel -checked $true -x 20 -y 80 -text "Chrome"
$swlist['OCCT'] = CheckBox -form $form_panel -checked $true -x 20 -y 110 -text "OCCT"
if ($info[2] -match 'Windows 11') {
    $swlist['Office 365 Desktop'] = CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 140 -text "Office 365 Desktop"
} else {
    $swlist['Office 365 Desktop'] = CheckBox -form $form_panel -checked $false -x 20 -y 140 -text "Office 365 Desktop"
}
$swlist['Revo Uninstaller'] = CheckBox -form $form_panel -checked $true -x 20 -y 170 -text "Revo Uninstaller"
$swlist['Supremo'] = CheckBox -form $form_panel -checked $true -x 20 -y 200 -text "Supremo"
if ($info[2] -match 'Windows 11') {
    $swlist['Teams'] = CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 230 -text "Teams"
} else {
    $swlist['Teams'] = CheckBox -form $form_panel -checked $true -x 20 -y 230 -text "Teams"
}
$swlist['TreeSize'] = CheckBox -form $form_panel -checked $true -x 20 -y 260 -text "TreeSize"
if ($info[2] -match 'Windows 11') {
    # l'installazione del client e' gia' gestita via PPPC
    $swlist['VPNnew'] = CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 290 -text "VPN Fortinet"
} else {
    <#
    per qualche motivo su Win10 viene installata una versione vecchia del client
    lascio il flag disponibile per poter scaricare installer e dipendenze da
    gestire manualmente
    #>
    $swlist['VPNnew'] = CheckBox -form $form_panel -checked $false -x 20 -y 290 -text "VPN Fortinet"
}
$swlist['7ZIP'] = CheckBox -form $form_panel -checked $true -x 20 -y 320 -text "7ZIP"
OKButton -form $form_panel -x 100 -y 370 -text "Ok"  | Out-Null
if ([string]::IsNullOrEmpty($winget_exe)) {
    foreach ($item in ('Acrobat Reader', 'BatteryMon', 'Chrome', 'OCCT', 'Revo Uninstaller', 'TreeSize', '7ZIP')) {
        $swlist[$item].Checked = $false
        $swlist[$item].Enabled = $false
    }
}
$result = $form_panel.ShowDialog()

$msstore_opts = '--source msstore --scope machine --accept-package-agreements --accept-source-agreements --silent'
$winget_opts = '--source winget --scope machine --accept-package-agreements --accept-source-agreements --silent'

# for more deeply inspection "Start-Process" cmdlet could be run also with "-RedirectStandardError" option
foreach ($item in ($swlist.Keys | Sort-Object)) {
    if ($swlist[$item].Checked -eq $true) {
        Write-Host -ForegroundColor Blue "[$item]"
        if ($item -eq 'Acrobat Reader') {
            Write-Host -NoNewline "Installing Acrobat Reader..."
            $StagingArgumentList = 'install {0} {1}' -f '--id Adobe.Acrobat.Reader.64-bit', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_Acrobat.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if (($stdout -match "Installazione riuscita") -or ($stdout -match "Successfully installed")) {
                if (Test-Path -Path "$env:PUBLIC\Desktop\Adobe Acrobat.lnk" -PathType Leaf) {
                    Remove-Item -Path "$env:PUBLIC\Desktop\Adobe Acrobat.lnk" -Force
                }
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
                <#
                    On older PC error 3010 can occur during installation: in most cases Acrobat 
                    seems to correctly run without any expected crash or further fails.

                    In the worst scenario try the following steps and see if that works for you:
                    * Run the Acrobat cleaner tool 
                        https://labs.adobe.com/downloads/acrobatcleaner.html
                    * Reboot the computer
                    * Reinstall the application using the link 
                        https://helpx.adobe.com/in/download-install/kb/acrobat-downloads.html
                #>
            }
        } elseif ($item -eq 'BatteryMon') {
            Write-Host -NoNewline "Installing BatteryMon..."
            $StagingArgumentList = 'install {0} {1}' -f '--id PassmarkSoftware.BatteryMon', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_BatteryMon.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if (($stdout -match "Installazione riuscita") -or ($stdout -match "Successfully installed")) {
                if (Test-Path -Path "$env:PUBLIC\Desktop\BatteryMon.lnk" -PathType Leaf) {
                    Remove-Item -Path "$env:PUBLIC\Desktop\BatteryMon.lnk" -Force
                }
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            }
        } elseif ($item -eq 'Chrome') {
            <# 
            There are several packages relayed to various Chrome flavours. 
            I selected 'Google Chrome (EXE)' package and not 'Google Chrome' one, beause of this:
            https://stackoverflow.com/questions/75647313/winget-install-my-app-receives-installer-hash-does-not-match
            #>  
            Write-Host -NoNewline "Installing Google Chrome..."
            $StagingArgumentList = 'install {0} {1}' -f '--id Google.Chrome.EXE', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_Chrome.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if (($stdout -match "Installazione riuscita") -or ($stdout -match "Successfully installed")) {
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            } 
        } elseif ($item -eq 'Revo Uninstaller') {
            Write-Host -NoNewline "Installing Revo Uninstaller..."
            $StagingArgumentList = 'install {0} {1}' -f '--id RevoUninstaller.RevoUninstaller', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_Revo.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if (($stdout -match "Installazione riuscita") -or ($stdout -match "Successfully installed")) {
                if (Test-Path -Path "$env:PUBLIC\Desktop\Revo Uninstaller.lnk" -PathType Leaf) {
                    Remove-Item -Path "$env:PUBLIC\Desktop\Revo Uninstaller.lnk" -Force
                }
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            }
        } elseif ($item -eq 'Office 365 Desktop') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x64&language=it-it&version=O16GA", "$env:USERPROFILE\Downloads\OfficeSetup.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Start-Process -Wait -FilePath "$env:USERPROFILE\Downloads\OfficeSetup.exe"
            Write-Host -ForegroundColor Green " DONE"
            <# there is some hash installation trouble on winget...
            Write-Host -NoNewline "Installing Microsoft Office 365..."
            $StagingArgumentList = 'install  "{0}" {1}' -f 'Microsoft 365 Apps for enterprise', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_o365.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if ($stdout -match "Installazione riuscita") {
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            }
            #>
        } elseif ($item -eq 'OCCT') {
            Write-Host -NoNewline "Installing OCCT..."
            $StagingArgumentList = 'install {0} {1}' -f '--id OCBase.OCCT.Personal', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_OCCT.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if (($stdout -match "Installazione riuscita") -or ($stdout -match "Successfully installed")) {
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            }
        } elseif ($item -eq 'Supremo') {
            Write-Host -NoNewline "Download software..."
            #$download.Downloadfile("https://www.agmsolutions.net/wp-content/uploads/assistenza/Assistenza_Remota.exe", "$env:PUBLIC\Desktop\Supremo.exe")
            Invoke-WebRequest -Uri 'https://www.nanosystems.it/public/download/Supremo.exe' -OutFile "$env:PUBLIC\Desktop\Supremo.exe"
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Write-Host -ForegroundColor Green " DONE"
        } elseif ($item -eq 'Teams') {
            Write-Host -NoNewline "Downloading standalone installer..."                
            $download.Downloadfile("https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix", "$env:PUBLIC\Desktop\MSTeams-x64.msix")
            Write-Host -ForegroundColor Green " DONE"
            [System.Windows.MessageBox]::Show("Please run the installer file [MSTeams-x64.msix] `nonce the target account has been logged in",'INFO','Ok','Info') | Out-Null
            <#
            Write-Host -NoNewline "Installing Microsoft Teams..."
            $StagingArgumentList = 'install  "{0}" {1} {2}' -f 'Microsoft Teams', $winget_opts, '--id Microsoft.Teams'
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_Teams.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if ($stdout -match "Installazione riuscita") {
                Write-Host -ForegroundColor Green " DONE"   
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            }
            #>
        } elseif ($item -eq 'VPNnew') {
            Write-Host -NoNewline "Download software..."
            # download dependencies (no more needed with client versions >= 7.4)
            # see https://community.fortinet.com/t5/Support-Forum/FortiClientVPN-client-doesn-t-work-with-Windows-11-24H2/m-p/366570#M259815)
            $answ = [System.Windows.MessageBox]::Show("Visual C++ Redistributable for Visual Studio 2015-2022`nis already installed?",'DEPENDENCIES','YesNo','Warning')
            if ($answ -eq "No") {
                $download.Downloadfile('https://aka.ms/vs/17/release/vc_redist.x64.exe', "C:\Users\Public\Desktop\vc_redist.x64.exe")
                $answ = [System.Windows.MessageBox]::Show("Keep in mind to install Visual C++ libraries`nBEFORE Fortinet client installation.`n`nPlease run setup once the target account has been logged in",'INFO','Ok','Info')
            }
            #Invoke-WebRequest -Uri 'https://links.fortinet.com/forticlient/win/vpnagent' -OutFile "C:\Users\Public\Desktop\FortiClientVPNOnlineInstaller.exe"
            $download.Downloadfile('https://links.fortinet.com/forticlient/win/vpnagent', "C:\Users\Public\Desktop\FortiClientVPNOnlineInstaller.exe")
            Write-Host -ForegroundColor Green " DONE"
            $answ = [System.Windows.MessageBox]::Show("Fortinet client installer downloaded.`n`nPlease run setup once the target account has been logged in",'INFO','Ok','Info')
        } elseif ($item -eq 'TreeSize') {
            Write-Host -NoNewline "Installing TreeSize Free..."
            $StagingArgumentList = 'install {0} {1}' -f '--id JAMSoftware.TreeSize.Free', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_Treesize.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if (($stdout -match "Installazione riuscita") -or ($stdout -match "Successfully installed")) {
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            }
        } elseif ($item -eq '7ZIP') {            
            Write-Host -NoNewline "Installing 7-Zip..."
            $StagingArgumentList = 'install {0} {1}' -f '--id 7zip.7zip', $winget_opts
            $winget_stdout_file = "$env:USERPROFILE\Downloads\wgetstdout_7zip.log"
            Start-Process -Wait -FilePath $winget_exe -ArgumentList $StagingArgumentList -NoNewWindow -RedirectStandardOutput $winget_stdout_file
            $stdout = Get-Content -Raw $winget_stdout_file
            if (($stdout -match "Installazione riuscita") -or ($stdout -match "Successfully installed")) {
                Write-Host -ForegroundColor Green " DONE"
            } else {
                Write-Host -ForegroundColor Red " FAILED"
                [System.Windows.MessageBox]::Show("Something has gone wrong, check the file `n[$winget_stdout_file]",'OOOPS!','Ok','Error') | Out-Null
            }
        }
    }
}

# removing temp files
Remove-Item $tmppath -Recurse -Force
$answ = [System.Windows.MessageBox]::Show("Remove log files of installations?",'TEMPORARY','YesNo','Info')
if ($answ -eq "Yes") {
    foreach($aLog in (Get-ChildItem "$env:USERPROFILE\Downloads").Name) {
        if ($alog -match "^wgetstdout_.+\.log?") {
            Remove-Item "$env:USERPROFILE\Downloads\$aLog" -Force
        }
    }
}

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
        $WhereIsPedoMellon =  "$workdir\Safety\PedoMellon.ps1"
        if (Test-Path -Path $WhereIsPedoMellon -PathType Leaf) {
            $compliant = $false
            while (!($compliant)) {
                $thepasswd =  PowerShell.exe -file $WhereIsPedoMellon `
                    -UserString $username  `
                    -MinimumLength 12 `
                    -Uppercase `
                    -TransLite `
                    -InsDels `
                    -Reverso `
                    -Rw 3

                if (($thepasswd -cmatch "[A-Z]+") -and ($thepasswd -match "[0-9]+") -and ($thepasswd -match "[!\$\?\*_\+#\@\^=%]+")) {
                    $compliant = $true
                }
            }
        } else { # old method
            Add-Type -AssemblyName 'System.Web'
            $thepasswd = [System.Web.Security.Membership]::GeneratePassword(10, 0)
        }

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
