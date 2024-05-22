<#
Name......: LocalAdmin-Cshared.ps1
Version...: 21.10.1
Author....: Dario CORRADA

This script will set user as local-admin and share the entire C: volume
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AutoMigration\\LocalAdmin-Cshared\.ps1$" > $null
$repopath = $matches[1]

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$repopath\Modules\Forms.psm1"

# temporary folder
$tmppath = 'C:\AUTOMIGRATION'
if (Test-Path $tmppath) {
    Remove-Item "$tmppath" -Recurse -Force
    Start-Sleep 2
}
New-Item -ItemType directory -Path $tmppath > $null

$backup_list = @{} # list of paths to be migrated
$usrlist = @() # list of user profiles

# select user profiles
$userlist = Get-ChildItem 'C:\Users'
$hsize = 200 + (30 * $userlist.Count)
$form_panel = FormBase -w 300 -h $hsize -text "USER FOLDERS"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(200,30)
$label.Text = "Select users to migrate:"
$form_panel.Controls.Add($label)
$vpos = 50
$boxes = @()
foreach ($elem in $userlist) {
    if ($elem.Name -eq $env:USERNAME) {
        $boxes += CheckBox -form $form_panel -checked $true -enabled $false -x 20 -y $vpos -text $elem.Name
        $vpos += 30
    } else {
        $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $elem.Name
        $vpos += 30
    }
}
$vpos += 20
OKButton -form $form_panel -x 90 -y $vpos -text "Ok"
$result = $form_panel.ShowDialog()
foreach ($box in $boxes) {
    if ($box.Checked -eq $true) {
        $usrname = $box.Text
        $usrlist += "$usrname"
    }
}

# load exclude list file
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select excluded path list file"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Plain text file | *.*'
$OpenFileDialog.ShowDialog() | Out-Null
$excludefile = $OpenFileDialog.filename

# listing paths
Write-Host -NoNewline -ForegroundColor Yellow "Looking for paths to be migrated..."
[string[]]$exclude_list = Get-Content -Path $excludefile
foreach ($usr in $usrlist) {
    $root_usr = 'C:\Users\' + $usr + '\'
    $rooted = Get-ChildItem $root_usr -Attributes D
    $exclude_usrlist = $exclude_list -replace ('\$username', $usr)
    foreach ($candidate in $rooted) {
        $candidate = 'Users\' + $usr + '\' + $candidate
        $decision = $true
        foreach ($item in $exclude_usrlist) {
            if ($item -eq $candidate) {
                $decision = $false
            }
        }
        if ($decision) {
            $full_path = 'C:\' + $candidate
            $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH'
            $StagingLogPath = $tmppath + '\test.log'
            $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $full_path, $StagingLogPath, $CommonRobocopyParams
            Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
            $StagingContent = Get-Content -Path $StagingLogPath
            $output = [system.String]::Join(" ", $StagingContent)
            $output -match "Byte:\s+(\d+)\s+\d+" > $null
            $size = $Matches[1]
            $backup_list[$candidate] = $size
            Remove-Item $StagingLogPath -Force
            Write-Host -NoNewline -ForegroundColor Yellow "."
        }
    }  
}

# resume of target paths selected
Write-Host -ForegroundColor Blue "`n`nRESUME OF TARGET PATHS:"
foreach($key in ($backup_list.Keys | Sort-Object)) {  
    Write-Host -NoNewline "C:\$key "
    $size = $backup_list[$key] / 1GB
    $formattato = '{0:0.0}' -f $size
    if ($size -le 5) {
        Write-Host -ForegroundColor Green "<5,0 GB"
    } elseif ($size -le 40) {
        Write-Host -ForegroundColor Yellow "$formattato GB"
    } else {
        Write-Host -ForegroundColor Red "$formattato GB"
    }
}

$answ = [System.Windows.MessageBox]::Show("Proceed?",'TARGET PATHS','YesNo','Info')
if ($answ -eq "Yes") {
    $logfile = "$tmppath\usrlist.log"
    foreach ($usr in ($usrlist | Sort-Object)) {
        "$usr" | Out-File $logfile -Encoding ASCII -Append 
    }
    $logfile = "$tmppath\targetpaths.log"
    foreach ($key in ($backup_list.Keys | Sort-Object)) {
        "$key" | Out-File $logfile -Encoding ASCII -Append
    }
} elseif ($answ -eq "No") {
    Remove-Item "$tmppath" -Recurse -Force
    Exit
}

# Backup wireless network profile
# see also https://www.tenforums.com/tutorials/3530-backup-restore-wireless-network-profiles-windows-10-a.html
Write-Host -NoNewline -ForegroundColor Yellow "`n`nRetrieving wireless network profile(s)..."
New-Item -ItemType directory -Path "$tmppath\wifi_profiles" > $null
netsh wlan export profile key=clear folder="$tmppath\wifi_profiles" > $null
Write-Host -ForegroundColor Green " DONE"

# shared paths
Write-Host -ForegroundColor Yellow "`nRetrieving shared paths..."
$drives = Get-PSDrive
foreach ($mounted in $drives) {
    if ($mounted.DisplayRoot -match "^\\") {
        New-Item -ItemType file "$tmppath\NetworkDrives.log" > $null
        $string = $mounted.Name + ":;" + $mounted.DisplayRoot
        $string | Out-File "$tmppath\NetworkDrives.log" -Encoding ASCII -Append
    }
}

# setting current user as local admin
Write-Host -ForegroundColor Yellow "`nSetting user(s) as local admin..."
$ErrorActionPreference = 'Stop'
foreach ($usr in $usrlist) {
    Write-Host -NoNewline "[$usr]"
    try {
        Add-LocalGroupMember -Group "Administrators" -Member $usr
        Write-Host -ForegroundColor Green " DONE"
    }
    catch {
        $alright = $error[0].FullyQualifiedErrorId -match "MemberExists"
        if ($alright) {
            # ignore error if the user is already localadmin
            Write-Host -ForegroundColor Green " DONE"
        } else {
            Write-Host -ForegroundColor Red " FAILED"
            Write-Host -ForegroundColor Red "$($error[0].ToString())"
            Pause
        }
    }
}
$ErrorActionPreference = 'Inquire'

# sharing C: volume
Write-Host -NoNewline -ForegroundColor Yellow "`nSharing C: volume..."
$answ = [System.Windows.MessageBox]::Show("This PC is joined to a domain?",'DOMAIN','YesNo','Info')
if ($answ -eq "Yes") {
    $ErrorActionPreference = 'Stop'
    try {
        $fulluser = $env:UserDomain + "\" + $env:UserName
        New-SmbShare –Name "C" –Path "C:\" -FullAccess $fulluser
        Write-Host -ForegroundColor Green " DONE"
    }
    catch {
        Write-Host -ForegroundColor Red " FAILED"
        Write-Host -ForegroundColor Red "$($error[0].ToString())"
        Pause
    }
    $ErrorActionPreference = 'Inquire'
} else {
    $answ = [System.Windows.MessageBox]::Show("Please manually share the C: volume...",'SHARE','Ok','Info')
}


# summary
New-Item -ItemType file "$tmppath\LocalAdmin-Cshared.log" > $null
("*** hostname: " + $env:computername + " ***") | Out-File "$tmppath\LocalAdmin-Cshared.log" -Encoding ASCII -Append
Get-NetIPConfiguration | Out-File "$tmppath\LocalAdmin-Cshared.log" -Encoding ASCII -Append
notepad "$tmppath\LocalAdmin-Cshared.log"
