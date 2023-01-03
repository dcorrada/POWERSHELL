<#
Name......: Reprofiler.ps1
Version...: 22.10.1
Author....: Dario CORRADA

This script will backup the exixsting account profiles and restore a brand new one
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Locales\\Reprofiler\.ps1$" > $null
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
Import-Module -Name "$workdir\Modules\Forms.psm1"

# getting users folders
$userlist = Get-ChildItem C:\Users

# control panel
$hsize = 150 + (25 * $userlist.Count)
$form_panel = FormBase -w 300 -h $hsize -text "USER FOLDERS"
Label -form $form_panel -x 10 -y 20 -w 200 -h 30 -text 'Select profile to be backupped:' | Out-Null
$form_panel.Controls.Add($label)
$vpos = 45
$boxes = @()
foreach ($elem in $userlist) {
    if ($vpos -eq 50) {
        $isfirst = $true
    } else {
        $isfirst = $false
    }
    $boxes += RadioButton -form $form_panel -checked $isfirst -x 20 -y $vpos -text $elem
    $vpos += 25
}
$vpos += 20
OKButton -form $form_panel -x 90 -y $vpos -text "Ok" | Out-Null
$result = $form_panel.ShowDialog()
foreach ($item in $boxes) {
    if ($item.Checked) {
        $theuser = $item.Text
    }
}


# identity check
if ($theuser -eq $env:USERNAME) {
    [System.Windows.MessageBox]::Show("Operating and target users doesn't be the same!",'ABORTING','Ok','Warning') | Out-Null
    Exit    
}

# removing account entries
Write-Host -NoNewline 'Removing local account... '
$ErrorActionPreference = 'Stop'
Try {
    Remove-LocalUser -Name $theuser
    Write-Host -ForegroundColor Green 'OK'                
}
Catch {
    Write-Host -ForegroundColor Red 'KO'
    Write-Output "ERROR: $($error[0].ToString())`n"
    $answ = [System.Windows.MessageBox]::Show("Exception emerge: Proceed anyway?",'WAIT','YesNo','Error')
    # in case we are dealing with an ADuser proceed anyway...
    if ($answ -eq "No") {    
        exit
    }
}
$ErrorActionPreference = 'Inquire'
Write-Host -NoNewline 'Disabling admin... '
$ErrorActionPreference = 'Stop'
Try {
    Remove-LocalGroupMember -Group 'Administrators' -Member $theuser
    Write-Host -ForegroundColor Green 'OK'                
}
Catch {
    Write-Host -ForegroundColor Red 'KO'
    Write-Output "ERROR: $($error[0].ToString())`n"
    $answ = [System.Windows.MessageBox]::Show("Exception emerge: Proceed anyway?",'WAIT','YesNo','Error')
    if ($answ -eq "No") {    
        exit
    }
}
$ErrorActionPreference = 'Inquire'
Write-Host -NoNewline 'Cleaning registry... '
$ErrorActionPreference = 'Stop'
Try {
    $SID =  Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
            Get-ItemProperty | Where-Object {$_.ProfileimagePath -match "C:\\Users\\$theuser" } | Select-Object -Property ProfileimagePath, PSChildName
    if ($SID) {
        $keypath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $SID
        Remove-Item -Path $keypath -Recurse
        Write-Host -ForegroundColor Green 'OK'
    } else {
        Write-Host -ForegroundColor Red 'KO'
    }       
}
Catch {
    Write-Host -ForegroundColor Red 'KO'
    Write-Output "ERROR: $($error[0].ToString())`n"
    $answ = [System.Windows.MessageBox]::Show("Exception emerge: Proceed anyway?",'WAIT','YesNo','Error')
    if ($answ -eq "No") {    
        exit
    }
}
$ErrorActionPreference = 'Inquire'

# backup folder
$tmppath = 'C:\REPROFILER'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
}
Write-Host -NoNewline 'Sweeping data... '
$ErrorActionPreference = 'Stop'
Try {
    Move-Item -Path "C:\Users\$theuser" -Destination $tmppath -Force
    Write-Host -ForegroundColor Green 'OK'
    [System.Windows.MessageBox]::Show("[$theuser] data are moved to [$tmppath]",'INFO','Ok','Info') | Out-Null         
}
Catch {
    Write-Host -ForegroundColor Red 'KO'
    Write-Output "ERROR: $($error[0].ToString())`n"
    $answ = [System.Windows.MessageBox]::Show("Exception emerge: Proceed anyway?",'WAIT','YesNo','Error')
    if ($answ -eq "No") {    
        exit
    }
}
$ErrorActionPreference = 'Inquire'

# select new user mode
$usrform = FormBase -w 350 -h 220 -text "ACCOUNT"
Label -form $usrform -x 10 -y 20 -w 80 -h 30 -text 'Username:' | Out-Null
$usrbox = TxtBox -form $usrform -x 100 -y 20 -w 200 -h 30 -text $theuser
Label -form $usrform -x 10 -y 50 -w 80 -h 30 -text 'Password:' | Out-Null
$pwdbox = TxtBox -form $usrform -x 100 -y 50 -w 200 -h 30 -masked $true
$localusr = RadioButton -form $usrform -x 30 -y 80 -w 120 -checked $true -text 'local user'
$adusr = RadioButton -form $usrform -x 180 -y 80 -checked $false -text 'AD user'
OKButton -form $usrform -x 100 -y 120 -text "Ok" | Out-Null
$result = $usrform.ShowDialog()

$username = $usrbox.Text
$thepasswd = $pwdbox.Text
$pwd = ConvertTo-SecureString $thepasswd -AsPlainText -Force

if ($localusr.Checked) {
    $ErrorActionPreference= 'Stop'
    Try {
        New-LocalUser -Name $username -Password $pwd -FullName $username -PasswordNeverExpires -AccountNeverExpires -Description "local user"
        Add-LocalGroupMember -Group "Administrators" -Member $username
        Write-Host -ForegroundColor Green "Local account created"
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }    
} elseif ($adusr.Checked) {
    [System.Windows.MessageBox]::Show("Connect to your domain before proceed...",'ACCOUNT','Ok','Warning') | Out-Null

    # add domain prefix to username
    $username = $usrname.Text
    $thiscomputer = Get-WmiObject -Class Win32_ComputerSystem
    $fullname = $thiscomputer.Domain + '\' + $username

    # test user
    [reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") > $null
    $principalContext = [System.DirectoryServices.AccountManagement.PrincipalContext]::new([System.DirectoryServices.AccountManagement.ContextType]'Machine',$env:COMPUTERNAME)
    if ($principalContext.ValidateCredentials($fullname,$thepasswd)) {
        # granting local admin privileges
        try {
            Add-LocalGroupMember -Group "Administrators" -Member $fullname
        }
        catch {
            [System.Windows.MessageBox]::Show("Cannot granting admin privilege to $username",'ACCOUNT','Ok','Error') | Out-Null
        }
    } else {
        [System.Windows.MessageBox]::Show("Invalid credentials for $username",'ACCOUNT','Ok','Error') | Out-Null
    }
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
