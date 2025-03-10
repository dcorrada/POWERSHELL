<#
Name......: OneShot.ps1
Version...: 25.03.1
Author....: Dario CORRADA

This script allow to navigate and select single scripts from this repository.
Then it dowload them and launch them locally.

[USING TOKEN]
You can get GitHub API token from https://github.com/settings/tokens/new 
For this script the following scopes suffices:
- repo:status       (Access commit status)
- repo_deployment   (Access deployment status)
- public_repo       (Access public repositories)

Then run the cmdlet "Set-GitHubAuthentication" (no option, interactive mode) 
providing your GitHub API Token in the "Password" field (the "Username field" 
will be ignored).  

The GitHub API Token will be cached on local machine across PowerShell 
sessions.  To clear caching, call "Clear-GitHubAuthentication".
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

# check NuGet
foreach ($pp in (Get-PackageProvider)) {
    if ($pp.Name -eq 'NuGet') {
        $foundit = $pp.Name
    }
}
if ($foundit -ne 'NuGet') {
    $ErrorActionPreference= 'Stop'
    Try {
        Install-PackageProvider -Name "NuGet" -MinimumVersion "2.8.5.208" -Force
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }
    $ErrorActionPreference= 'Inquire'
}

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing third party modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module PowerShellForGitHub
        $ThirdParty = 'Ok'
        
        # comment out the following lines if you cached a GitHub API Token
        Set-GitHubConfiguration -DisableTelemetry -WebRequestTimeoutSec 120 -SuppressNoTokenWarning
        $disclaimer = @"
The module [PowerShellForGitHub] has not yet been configured with a personal GitHub Access token.

The script can still be run, but GitHub will limit your usage to 60 queries per hour.
"@
        [System.Windows.MessageBox]::Show($disclaimer,'DISCLAIMER','Ok','warning') > $null

    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'PowerShellForGitHub')) {
            Install-Module PowerShellForGitHub -Scope AllUsers -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [PowerShellForGitHub] module: click Ok to restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } else {
            Write-Output "`nError: $($error[0].ToString())"
            Pause
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

# GitHub coordinates
$theOwner = 'dcorrada'
$theRepo = 'POWERSHELL'
$theBranch = 'master'
$altBranches = Get-GitHubRepositoryBranch -OwnerName $theOwner -RepositoryName $theRepo 

# get working directory
$workdir = "$env:USERPROFILE\Downloads\dcorrada.OneShot"
if (Test-Path $workdir) {
    Remove-Item -Path $workdir -Recurse -Force > $null
}
New-Item -ItemType directory -Path $workdir > $null
New-Item -ItemType directory -Path "$workdir\Modules" > $null
$download = New-Object net.webclient
foreach ($psm1File in (Get-GitHubContent `
    -OwnerName $theOwner `
    -RepositoryName $theRepo `
    -BranchName  $theBranch `
    -Path 'Modules').entries) {
        if ($psm1File.type -eq 'file') {
            $download.Downloadfile("$($psm1File.download_url)", "$workdir\Modules\$($psm1File.name)")
        }
    }

# graphical stuff
Import-Module -Name "$workdir\Modules\Forms.psm1"

<# *******************************************************************************
                                    HISTORY
******************************************************************************* #>
$cacheFile = "$env:APPDATA\dcorrada.OneShot.csv"
$cachedItems = @()
if (Test-Path $cacheFile -PathType Leaf) {
    $cachedItems += Import-Csv -Path $cacheFile
} else {
    New-Item -ItemType File -Path $cacheFile > $null
    "LABEL,NAME,PATH,URL" | Out-File $cacheFile -Encoding utf8 -Append
}


<# *******************************************************************************
                                    DIALOG
******************************************************************************* #>
$currentFolder = 'root'

# filtering only scripts and paths (nor modules folder)
$CurrentItems = (Get-GitHubContent `
    -OwnerName $theOwner `
    -RepositoryName $theRepo `
    -BranchName  $theBranch `
    ).entries | ForEach-Object {
        if ((($_.type -eq 'dir') -or ($_.name -match "\.ps1$")) -and !($_.name -eq 'Modules')) {
            New-Object -TypeName PSObject -Property @{
                NAME    = $_.name
                PATH    = $_.path
                URL     = $_.download_url
            } | Select NAME, PATH, URL   
        }
    }

$isChecked = "none"
$intoBox = ''
$continueBrowsing = $true
while ($continueBrowsing) {
    $hmin = ((($CurrentItems.Count) * 30) + 90)
    if ($hmin -lt 500) {
        $hmin = 500
    }
    $adialog = FormBase -w 850 -h $hmin -text "SELECT AN ITEM [$currentFolder]@$theBranch"
    $they = 20
    $choices = @()

    $ExLinkLabel = New-Object System.Windows.Forms.LinkLabel
    $ExLinkLabel.Location = New-Object System.Drawing.Size(510,($hmin - 80))
    $ExLinkLabel.Size = New-Object System.Drawing.Size(300,30)
    $ExLinkLabel.Text = "Get more info about my repository"
    $ExLinkLabel.TextAlign = 'MiddleRight'
    $ExLinkLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
    $ExLinkLabel.add_Click({[system.Diagnostics.Process]::start("https://github.com/dcorrada/POWERSHELL")})
    $adialog.Controls.Add($ExLinkLabel)

    # list of items in the current path
    foreach ($ItemName in ($CurrentItems.Name | Sort-Object)) {
        if (($isChecked -eq 'none') -and ($choices.Count -lt 1)) {
            $gotcha = $true
        } elseif ($ItemName -eq $isChecked) {
            $gotcha = $true
        } else {
            $gotcha = $false
        }
        if ($ItemName -match "\.ps1$") {
            $aText = $ItemName
        } else {
            $aText = '[+] ' + $ItemName
        }
        $choices += RadioButton -form $adialog -x 20 -y $they -w 230 -checked $gotcha -text $aText
        $they += 30 
    }

    # preview text box
    TxtBox -form $adialog -x 260 -y 20 -w 550 -h 200 -text $intoBox -multiline $true | Out-Null
    
    # buttons
    $PreviousBut = RETRYButton -form $adialog -x 260 -y 230 -w 75 -text "UP"
    $PreviousBut.DialogResult = [System.Windows.Forms.DialogResult]::CANCEL
    $NextBut = OKButton -form $adialog -x 335 -y 230 -w 75 -text "GO"
    $PreviewBut = RETRYButton -form $adialog -x 710 -y 230 -text "Preview"
    $AbortBut = RETRYButton -form $adialog -x 710 -y 270 -text "Quit"
    $AbortBut.DialogResult = [System.Windows.Forms.DialogResult]::ABORT
    $CacheBut = RETRYButton -form $adialog -x 430 -y 270 -w 230 -text "Clear Cached Items"
    $CacheBut.DialogResult = [System.Windows.Forms.DialogResult]::IGNORE

    # list of alternative branches
    $centar = Label -form $adialog -x 430 -y 230 -h 30 -w 70 -text "Alternative Branches"
    $centar.TextAlign = 'MiddleCenter'
    $DropBranch = DropDown -form $adialog -x 510 -y 235 -w 150 -opts ($altBranches.BranchName | Sort-Object)
    $DropBranch.text = $theBranch
    
    # list of items form history file
    $astoria = Label -form $adialog -x 260 -y 290 -h 25 -text "RECENT LAUNCHES:"
    $astoria.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
    $they = 310
    foreach ($cachedItem in $cachedItems) {
        $choices += RadioButton -form $adialog -x 270 -y $they -w 300 -checked $false -text $cachedItem.LABEL
        $they += 25
    }

    $goahead = $adialog.ShowDialog()

    # looking for selected item properties
    $selectedItem = @{
        NAME = 'none'
        URL = 'null'
        PATH = 'null'
    }
    foreach ($currentOpt in $choices) {
        if ($currentOpt.Checked) {
            $currentOpt.Text -match "^(\[\+\] )*([a-zA-Z_\-\.0-9]+)( \[.+\])*$" | Out-Null
            $isChecked = "$($matches[2])"
            foreach ($anItem in $CurrentItems) {
                if ($anItem.NAME -eq $isChecked) {
                    $selectedItem.NAME = "$($anItem.NAME)"
                    $selectedItem.URL = "$($anItem.URL)"
                    $selectedItem.PATH = "$($anItem.PATH)"
                }
            }
            if ($selectedItem.NAME -eq 'none') {
                foreach ($anItem in $cachedItems) {
                    if ($anItem.NAME -eq $isChecked) {
                        $selectedItem.NAME = "$($anItem.NAME)"
                        $selectedItem.URL = "$($anItem.URL)"
                        $selectedItem.PATH = "$($anItem.PATH)"
                    }
                }                
            }
        }
    }                
    
    # actions for clicking [Preview] button
    if ($goahead -eq 'RETRY') {
        if ([string]::IsNullOrEmpty($selectedItem.URL)) {
            $arrayContent = @("Items in folder [$($currentOpt.Text)]:", "")
            foreach ($entry in (Get-GitHubContent `
                -OwnerName $theOwner `
                -RepositoryName $theRepo `
                -BranchName  $theBranch `
                -Path $selectedItem.PATH
                ).entries) {
                    $arrayContent += "  * $($entry.name)"
                }
            $intoBox = $arrayContent | Out-String
        } else {
            $fileContent = (Invoke-WebRequest -Uri $selectedItem.URL -UseBasicParsing).Content.Split("`n")
            $maxlines = 19
            if ($fileContent.Count -lt $maxlines) {
                $maxlines = $fileContent.Count - 1
            }
            $intoBox = $fileContent[0..$maxlines] + "[...]" | Out-String
        }

    # actions for clicking [GO] button
    } elseif ($goahead -eq 'OK') {
        $intoBox = ''
        if ($theBranch -ne $DropBranch.text) {
            # change branch and reset to the root path
            $isChecked = 'none'
            $theBranch = $DropBranch.text
            $currentFolder = 'root'
            $CurrentItems = (Get-GitHubContent `
                -OwnerName $theOwner `
                -RepositoryName $theRepo `
                -BranchName  $theBranch `
                ).entries | ForEach-Object {
                    if ((($_.type -eq 'dir') -or ($_.name -match "\.ps1$")) -and !($_.name -eq 'Modules')) {
                        New-Object -TypeName PSObject -Property @{
                            NAME    = $_.name
                            PATH    = $_.path
                            URL     = $_.download_url
                        } | Select NAME, PATH, URL   
                    }
                }
        } elseif ([string]::IsNullOrEmpty($selectedItem.URL)) {
            # navigate a path
            $isChecked = 'none'
            $currentFolder = '/' + $selectedItem.PATH
            $CurrentItems = (Get-GitHubContent `
                -OwnerName $theOwner `
                -RepositoryName $theRepo `
                -BranchName  $theBranch `
                -Path $selectedItem.PATH
                ).entries | ForEach-Object {
                    if ((($_.type -eq 'dir') -or ($_.name -match "\.ps1$")) -and !($_.name -eq 'Modules')) {
                        New-Object -TypeName PSObject -Property @{
                            NAME    = $_.name
                            PATH    = $_.path
                            URL     = $_.download_url
                        } | Select NAME, PATH, URL   
                    }
                }
        } else {
            # run the script
            $continueBrowsing = $false
        }

    # actions for clicking [UP] button
    } elseif ($goahead -eq 'CANCEL') {
        $isChecked = 'none'
        $intoBox = ''
        if ($currentFolder -ne 'root') {
            if ($currentFolder -match "(^/.+)/[a-zA-Z_\-\.0-9]+$") {
                $currentFolder = $matches[1]
                $CurrentItems = (Get-GitHubContent `
                -OwnerName $theOwner `
                -RepositoryName $theRepo `
                -BranchName  $theBranch `
                -Path $currentFolder
                ).entries | ForEach-Object {
                    if ((($_.type -eq 'dir') -or ($_.name -match "\.ps1$")) -and !($_.name -eq 'Modules')) {
                        New-Object -TypeName PSObject -Property @{
                            NAME    = $_.name
                            PATH    = $_.path
                            URL     = $_.download_url
                        } | Select NAME, PATH, URL   
                    }
                }
            } else {
                $currentFolder = 'root'
                $CurrentItems = (Get-GitHubContent `
                -OwnerName $theOwner `
                -RepositoryName $theRepo `
                -BranchName  $theBranch `
                ).entries | ForEach-Object {
                    if ((($_.type -eq 'dir') -or ($_.name -match "\.ps1$")) -and !($_.name -eq 'Modules')) {
                        New-Object -TypeName PSObject -Property @{
                            NAME    = $_.name
                            PATH    = $_.path
                            URL     = $_.download_url
                        } | Select NAME, PATH, URL   
                    }
                }
            }
        }
    
    # actions for clicking [Clear Cached Items] button
    } elseif ($goahead -eq 'IGNORE') {
        Remove-Item -Path $cacheFile -Force > $null
        [System.Windows.MessageBox]::Show("Cache will be cleared on `nnext restart of the script",'CLEAR CACHE','Ok','Info')

    # actions for clicking [Quit] button
    } elseif ($goahead -eq 'ABORT') {
        Set-Location $env:USERPROFILE     
        Remove-Item -Path $workdir -Recurse -Force > $null
        exit
    }

}


<# *******************************************************************************
                                    RUNNING
******************************************************************************* #>
Write-Host -NoNewline "Preparing $($selectedItem.NAME)... "
$folders = $selectedItem.PATH -split('/')
$runpath = $workdir
for ($i = 0; $i -lt ($folders.Count - 1); $i++) { # last element is the script filename
    $runpath = $runpath + '\' + $folders[$i]
    if (!(Test-Path $runpath)) {
        New-Item -ItemType Directory -Path $runpath | Out-Null
    }
    Set-Location $runpath
}
$download.Downloadfile("$($selectedItem.URL)", "$runpath\$($selectedItem.NAME)")
Write-Host -ForegroundColor Green "DONE"

Write-Host -NoNewline "Looking for dependencies... "

# requiring PedoMellon password generator
if ($selectedItem.NAME -match 'Init_PC.ps1') {
    New-Item -ItemType Directory -Path "$workdir\Safety" | Out-Null
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/master/Safety/PedoMellon.ps1', "$workdir\Safety\PedoMellon.ps1")    
}

# script dedicated to MS Graph
if ($runpath -match 'Graph') {
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/master/Graph/AppKeyring.ps1', "$workdir\Graph\AppKeyring.ps1")
    New-Item -ItemType Directory -Path "$workdir\Safety" | Out-Null
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/master/Safety/PSWallet.ps1', "$workdir\Safety\PSWallet.ps1")
}

# MOTA integration
if ($selectedItem.NAME -match 'AssignedLicenses') {
    New-Item -ItemType Directory -Path "$workdir\AzureAD" | Out-Null
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/unstable/AzureAD/MOTA.ps1', "$workdir\AzureAD\MOTA.ps1")
}

# stuff scripts adopting PSWallet keyring
$found = Get-content -path "$runpath\$($selectedItem.NAME)" | Select-String -pattern 'Stargate.ps1' -encoding ASCII -CaseSensitive
if ($found.Count -ge 1) {
    if (!(Test-Path "$workdir\Safety")) {
        New-Item -ItemType Directory -Path "$workdir\Safety" | Out-Null
    }
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/master/Safety/PSWallet.ps1', "$workdir\Safety\PSWallet.ps1")
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/master/Safety/Stargate.ps1', "$workdir\Safety\Stargate.ps1")
}
Write-Host -ForegroundColor Green "DONE"

# run the script
if ($selectedItem.NAME -eq 'OneShot.ps1') {
    [System.Windows.MessageBox]::Show("OneShot does not lauch itself, fresh version is downloaded in `n[$workdir]",'UPDATE','Ok','Info') | Out-Null
} else {
    Write-Host -ForegroundColor Cyan "Launching $($selectedItem.NAME)..."
    Start-Sleep -Milliseconds 1000
    Clear-Host
    PowerShell.exe "& ""$runpath\$($selectedItem.NAME)"
}

<# *******************************************************************************
                                    CLEANING
******************************************************************************* #>

# update history (up to 5 items)
if ($cachedItems.URL -cnotcontains $selectedItem.URL) {
    $cachedItems = ,(New-Object -TypeName PSObject -Property @{
        LABEL   = "$($selectedItem.NAME) [$theBranch]"
        NAME    = $selectedItem.NAME
        PATH    = $selectedItem.PATH
        URL     = $selectedItem.URL
    } | Select LABEL, NAME, PATH, URL) + $cachedItems
}
if ($cachedItems.Count -lt 6) {
    $cachedItems | Export-Csv -Path $cacheFile -NoTypeInformation
} else {
    $cachedItems[0..4] | Export-Csv -Path $cacheFile -NoTypeInformation
}

# a quote a day
if (!(Test-Path "$workdir\miscellaneous")) {
    New-Item -ItemType Directory -Path "$workdir\miscellaneous" | Out-Null
}
$download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/tempus/miscellaneous/Quotes.ps1', "$workdir\miscellaneous\Quotes.ps1")
PowerShell.exe "& ""$workdir\miscellaneous\Quotes.ps1"

# delete temps
if ($selectedItem.NAME -cne 'OneShot.ps1') {
    $answ = [System.Windows.MessageBox]::Show("Your script [$($selectedItem.NAME)] is terminated: `ndo you want to locally delete it?",'CLEAN','YesNo','Info')
    if ($answ -eq "Yes") {
        Set-Location $env:USERPROFILE     
        Remove-Item -Path $workdir -Recurse -Force > $null
    }
}
