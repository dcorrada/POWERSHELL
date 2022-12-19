<#
Name......: Account_cleaner.ps1
Version...: 22.12.2
Author....: Dario CORRADA

This script removes account from local computer
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Locales\\Account_cleaner\.ps1$" > $null
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

# gathering user profiles candidates to remove
$usrcandidates = @{}
Write-Host -NoNewline 'Fetching data...'

# getting user folders
foreach ($item in Get-ChildItem C:\Users) {
    $usrcandidates[$item.Name] = @{
        FullPath = 'C:\Users\' + $item.Name
        IsAdmin = 'na'
        Domain = 'na'
        Orphan = 'No'
        SID = 'na'
    }
    Write-Host -NoNewline '.'
}

# looking for registry keys
$regges =   Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
            Get-ItemProperty | Where-Object {$_.ProfileimagePath -match "C:\\Users" } | Select-Object -Property ProfileimagePath, PSChildName
foreach ($item in $regges) {
    $item.ProfileImagePath -match "^C:\\Users\\([a-zA-Z_\-\.\\\s0-9:]+)$" > $null
    if ($usrcandidates.ContainsKey($matches[1])) {
        $usrcandidates[$matches[1]].SID = $item.PSChildName
    }
    Write-Host -NoNewline '.'
}

<# checking local or AD

NOTE:   the full AD userlist is shown solely whenever DC is visible AND
        from an active ADuser sesson.
        The cmdlet Get-CimInstance works without importing any AD module
#>
$halloffame = @{}
foreach ($item in Get-CimInstance Win32_UserAccount) {
    $halloffame[$item.Name] = $item.Domain
    Write-Host -NoNewline '.'
}
foreach ($item in $usrcandidates.Keys) {
    if ($halloffame.ContainsKey($item)) {
        if ($halloffame[$item] -eq $env:COMPUTERNAME) {
            $usrcandidates[$item].Domain = 'local'
        } else {
            $usrcandidates[$item].Domain = 'AD'
        }
    } else {
        $usrcandidates[$item].Orphan = 'Yes'
    }
    Write-Host -NoNewline '.'
}

# checking admin privileges
$group = [ADSI] "WinNT://./Administrators,group"
$members = @($group.psbase.Invoke("Members"))
$AdminList = ($members | ForEach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null);Write-Host -NoNewline '.'})
foreach ($item in $AdminList) {
    if ($usrcandidates.ContainsKey($item)) {
        $usrcandidates[$item].IsAdmin = 'Yes'
        $usrcandidates[$item].Orphan = 'No'
    }
    Write-Host -NoNewline '.'
}

Write-Host -ForegroundColor Green 'OK'

# control panel
$hsize = 150 + (30 * $usrcandidates.Count)
$form_panel = FormBase -w 300 -h $hsize -text "USERS LIST"
$label = Label -form $form_panel -x 10 -y 20 -w 200 -h 30 -text 'Select profiles to be deleted:'
$vpos = 50
$boxes = @()
foreach ($item in ($usrcandidates.Keys | Sort-Object)) {
    if ($usrcandidates[$item].Orphan -eq 'Yes') {
        $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $item -enabled $false
    } else {
        $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $item
    }
    $vpos += 30
}
$vpos += 20
OKButton -form $form_panel -x 90 -y $vpos -text "Ok"
$result = $form_panel.ShowDialog()

# perform operations
foreach ($box in $boxes) {
    if ($box.Checked -eq $true) {
        $theuser = $box.Text
        Clear-Host
        Write-Host -ForegroundColor Blue "*** Oblivioning [$theuser] ***"

        if ($usrcandidates[$theuser].IsAdmin -eq 'Yes') {
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
        }

        if ($usrcandidates[$theuser].Domain -eq 'local') {
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
                if ($answ -eq "No") {    
                    exit
                }
            }
            $ErrorActionPreference = 'Inquire'
        }

        if ($usrcandidates[$theuser].SID -ne 'na') {
            Write-Host -NoNewline 'Cleaning registry... '
            $ErrorActionPreference = 'Stop'
            Try {
                $keypath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $usrcandidates[$theuser].SID
                Remove-Item -Path $keypath -Recurse
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
        }

        Write-Host -NoNewline 'Sweeping data... '
        $ErrorActionPreference = 'Stop'
        Try {
            Move-Item -Path $usrcandidates[$theuser].FullPath -Destination $tmppath -Force
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

        Start-Sleep -Seconds 2
    }    
}

# removing TEMPSOFTWARE
$ErrorActionPreference = 'Stop'
Try {
    # elevate Explorer
    taskkill /F /IM explorer.exe
    Start-Sleep 3
    Start-Process C:\Windows\explorer.exe

    $shell = new-object -comobject "Shell.Application"
    $item = $shell.Namespace(0).ParseName($tmppath)
    $item.InvokeVerb("delete")
    Clear-RecycleBin -Force -Confirm:$false
}
Catch {
    $answ = [System.Windows.MessageBox]::Show("Unable to delete temp files",'FAIL','Ok','Error')
    Write-Output "`n***`nError: $($error[0].ToString())"
}
$ErrorActionPreference = 'Inquire'
