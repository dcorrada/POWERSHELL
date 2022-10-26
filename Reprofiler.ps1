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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Reprofiler\.ps1$" > $null
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

# create temporary directory
$tmppath = 'C:\TEMPSOFTWARE'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
}

# getting users list
$userlist = Get-ChildItem C:\Users

# control panel
$hsize = 150 + (30 * $userlist.Count)
$form_panel = FormBase -w 300 -h $hsize -text "USER FOLDERS"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(200,30)
$label.Text = "Select profiles to be backupped:"
$form_panel.Controls.Add($label)
$vpos = 50
$boxes = @()
foreach ($elem in $userlist) {
    $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $elem
    $vpos += 30
}
$vpos += 20
OKButton -form $form_panel -x 90 -y $vpos -text "Ok"
$result = $form_panel.ShowDialog()

# get a list of local users
$locales = Get-LocalUser

foreach ($box in $boxes) {
    if ($box.Checked -eq $true) {
        $theuser = $box.Text
        Write-Host -NoNewline "Removing account [$theuser]..."
        $ErrorActionPreference= 'Stop'
        Try {
            # remove local account
            if ($locales.Name -contains $theuser) {
                Remove-LocalUser -Name $theuser
            }
            # search and remove keys
            $record = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
                Get-ItemProperty | Where-Object {$_.ProfileimagePath -match "C:\\Users\\$theuser" } | Select-Object -Property ProfileimagePath, PSChildName
            $keypath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $record.PSChildName
            Remove-Item -Path $keypath -Recurse
            Write-Host -ForegroundColor Green ' DONE'
        }
        Catch {
            Write-Host -ForegroundColor Red ' FAILED'
            Write-Output "`nError: $($error[0].ToString())"
            Pause
            exit
        }    
        Write-Host -NoNewline "Backupping profile of [$theuser]..."
        $ErrorActionPreference= 'Stop'
        Try {
            # renaming user folder
            Rename-Item "C:\Users\$theuser" "C:\Users\OLD-$theuser-OLD"
            Write-Host -ForegroundColor Green ' DONE'
        }
        Catch {
            Write-Host -ForegroundColor Red ' FAILED'
            Write-Output "`nError: $($error[0].ToString())"
            Pause
            exit
        }
    }
}

# creating local account
$answ = [System.Windows.MessageBox]::Show("Create local account?",'ACCOUNT','YesNo','Info')
if ($answ -eq "Yes") {
    $form = FormBase -w 520 -h 270 -text "ACCOUNT"
    $font = New-Object System.Drawing.Font("Arial", 12)
    $form.Font = $font
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(500,30)
    $label.Text = "Username:"
    $form.Controls.Add($label)
    $usrname = New-Object System.Windows.Forms.TextBox
    $usrname.Location = New-Object System.Drawing.Point(10,60)
    $usrname.Size = New-Object System.Drawing.Size(450,30)
    $form.Controls.Add($usrname)
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,100)
    $label2.Size = New-Object System.Drawing.Size(500,30)
    $label2.Text = "Fullname:"
    $form.Controls.Add($label2)
    $fullname = New-Object System.Windows.Forms.TextBox
    $fullname.Location = New-Object System.Drawing.Point(10,140)
    $fullname.Size = New-Object System.Drawing.Size(450,30)
    $form.Controls.Add($fullname)
    $OKButton = New-Object System.Windows.Forms.Button
    OKButton -form $form -x 200 -y 190 -text "Ok"
    $form.Topmost = $true
    $result = $form.ShowDialog()
    $username = $usrname.Text
    $completo = $fullname.Text
    $form_pswd = FormBase -w 450 -h 230 -text "CREATE PASSWORD"
    $personal = RadioButton -form $form_pswd -checked $true -x 30 -y 20 -text "Set your own password"
    $randomic  = RadioButton -form $form_pswd -checked $false -x 30 -y 50 -text "Generate random password"
    OKButton -form $form_pswd -x 90 -y 120 -text "Ok"
    $result = $form_pswd.ShowDialog()
    if ($result -eq "OK") {
        if ($personal.Checked) {
            $form = FormBase -w 520 -h 200 -text "PASSWORD"
            $font = New-Object System.Drawing.Font("Arial", 12)
            $form.Font = $font
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10,20)
            $label.Size = New-Object System.Drawing.Size(500,30)
            $label.Text = "Password:"
            $form.Controls.Add($label)
            $usrname = New-Object System.Windows.Forms.TextBox
            $usrname.Location = New-Object System.Drawing.Point(10,60)
            $usrname.Size = New-Object System.Drawing.Size(450,30)
            $usrname.PasswordChar = '*'
            $form.Controls.Add($usrname)
            $OKButton = New-Object System.Windows.Forms.Button
            OKButton -form $form -x 200 -y 120 -text "Ok"
            $form.Topmost = $true
            $result = $form.ShowDialog()
            $thepasswd = $usrname.Text
        } elseif ($randomic.Checked) {
            Add-Type -AssemblyName 'System.Web'
            $thepasswd = [System.Web.Security.Membership]::GeneratePassword(10, 0)
        }
    }
    $pwd = ConvertTo-SecureString $thepasswd -AsPlainText -Force
    $ErrorActionPreference= 'Stop'
    Try {
        New-LocalUser -Name $username -Password $pwd -FullName $completo -PasswordNeverExpires -AccountNeverExpires -Description "utente locale"
        Add-LocalGroupMember -Group "Administrators" -Member $username
        Write-Host -ForegroundColor Green "Local account created"
        Write-Host "Username...: " -NoNewline
        Write-Host $username -ForegroundColor Cyan
        Write-Host "Password...: " -NoNewline
        Write-Host $thepasswd -ForegroundColor Cyan
        $ErrorActionPreference= 'Inquire'
        Pause
    }
    Catch {
        Write-Host -ForegroundColor Red "Creating local account failed"
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }    
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
