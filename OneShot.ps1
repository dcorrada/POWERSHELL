<#
Name......: OneShot.ps1
Version...: 24.08.1
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

+++ TODO +++
* Sort and/or mark folders and scripts separately
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
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
    Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:"
    Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force"
    Write-Host "Afterwards restart this script"
    Pause
}
$ErrorActionPreference= 'Inquire'

# importing third party modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module PowerShellForGitHub
        $ThirdParty = 'Ok'
        
        # comment out the following lines if you cached a GitHub API Token
        Set-GitHubConfiguration -DisableTelemetry -SessionOnly
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

# get working directory
$workdir = "$env:USERPROFILE\Downloads\dcorrada.OneShot"
if (Test-Path $workdir) {
    Remove-Item -Path $workdir -Recurse -Force > $null
}
New-Item -ItemType directory -Path $workdir > $null
New-Item -ItemType directory -Path "$workdir\Modules" > $null
$download = New-Object net.webclient
foreach ($psm1File in (Get-GitHubContent `
    -OwnerName 'dcorrada' `
    -RepositoryName 'POWERSHELL' `
    -Path 'Modules').entries) {
        if ($psm1File.type -eq 'file') {
            $download.Downloadfile("$($psm1File.download_url)", "$workdir\Modules\$($psm1File.name)")
        }
    }

# graphical stuff
Import-Module -Name "$workdir\Modules\Forms.psm1"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

<# *******************************************************************************
                                    HISTORY
******************************************************************************* #>
$cacheFile = "$env:APPDATA\dcorrada.OneShot.csv"
$cachedItems = @()
if (Test-Path $cacheFile -PathType Leaf) {
    $cachedItems += Import-Csv -Path $cacheFile
} else {
    New-Item -ItemType File -Path $cacheFile > $null
    "NAME,PATH,URL" | Out-File $cacheFile -Encoding utf8 -Append
}


<# *******************************************************************************
                                    DIALOG
******************************************************************************* #>
$currentFolder = 'root'

# filtering only scripts and paths (nor modules folder)
$CurrentItems = (Get-GitHubContent `
    -OwnerName 'dcorrada' `
    -RepositoryName 'POWERSHELL' `
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
    $adialog = FormBase -w 720 -h $hmin -text "SELECT AN ITEM [$currentFolder]"
    $they = 20
    $choices = @()

    # list of items in the current path
    foreach ($ItemName in ($CurrentItems.Name | Sort-Object)) {
        if (($isChecked -eq 'none') -and ($choices.Count -lt 1)) {
            $gotcha = $true
        } elseif ($ItemName -eq $isChecked) {
            $gotcha = $true
        } else {
            $gotcha = $false
        }
        $choices += RadioButton -form $adialog -x 20 -y $they -checked $gotcha -text $ItemName
        $they += 30 
    }

    # preview text box
    TxtBox -form $adialog -x 230 -y 20 -w 450 -h 200 -text $intoBox -multiline $true | Out-Null
    
    # buttons
    $PreviousBut = RETRYButton -form $adialog -x 230 -y 230 -w 75 -text "UP"
    $PreviousBut.DialogResult = [System.Windows.Forms.DialogResult]::CANCEL
    $NextBut = OKButton -form $adialog -x 305 -y 230 -w 75 -text "GO"
    $PreviewBut = RETRYButton -form $adialog -x 580 -y 230 -text "Preview"
    
    # list of items form history file
    Label -form $adialog -x 230 -y 280 -h 25 -text "RECENT LAUNCHES:" | Out-Null
    $they = 300
    foreach ($cachedItem in $cachedItems) {
        $choices += RadioButton -form $adialog -x 240 -y $they -checked $false -text $cachedItem.NAME
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
            $isChecked = "$($currentOpt.Text)"
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
                -OwnerName 'dcorrada' `
                -RepositoryName 'POWERSHELL' `
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
        if ([string]::IsNullOrEmpty($selectedItem.URL)) {
            $isChecked = 'none'
            $currentFolder = '/' + $selectedItem.PATH
            $CurrentItems = (Get-GitHubContent `
                -OwnerName 'dcorrada' `
                -RepositoryName 'POWERSHELL' `
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
            $continueBrowsing = $false
        }

    # actions for clicking [UP] button
    } elseif ($goahead -eq 'CANCEL') {
        $isChecked = 'none'
        if ($currentFolder -ne 'root') {
            if ($currentFolder -match "(^/.+)/[a-zA-Z_\-\.0-9]+$") {
                $currentFolder = $matches[1]
                $CurrentItems = (Get-GitHubContent `
                -OwnerName 'dcorrada' `
                -RepositoryName 'POWERSHELL' `
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
                -OwnerName 'dcorrada' `
                -RepositoryName 'POWERSHELL' `
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

if ($runpath -match 'Graph') {
    Write-Host -NoNewline "Getting dependencies... "
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/master/Graph/AppKeyring.ps1', "$workdir\Graph\AppKeyring.ps1")
    New-Item -ItemType Directory -Path "$workdir\Safety" | Out-Null
    $download.Downloadfile('https://raw.githubusercontent.com/dcorrada/POWERSHELL/master/Safety/PSWallet.ps1', "$workdir\Safety\PSWallet.ps1")
    Write-Host -ForegroundColor Green "DONE"
}


<# 
DA FARE...

Check aggiuntivo sullo script selezionato e caricare PSWallet e Stargate se 
viene richiamato il wallet (magari greppare nello script una stringa del tipo 
"Error connecting to PSWallet").

Lanciare lo script selezionato
#>

<# *******************************************************************************
                                    CLEANING
******************************************************************************* #>

# update history
if ($cachedItems.NAME -cnotcontains $selectedItem.NAME) {
    $cachedItems += New-Object -TypeName PSObject -Property @{
        NAME    = $selectedItem.NAME
        PATH    = $selectedItem.PATH
        URL     = $selectedItem.URL
    } | Select NAME, PATH, URL
}
$cachedItems[($cachedItems.Count - 5)..($cachedItems.Count - 1)] | Export-Csv -Path $cacheFile -NoTypeInformation

# delete temps
$answ = [System.Windows.MessageBox]::Show("Delete downloaded files?",'CLEAN','YesNo','Info')
if ($answ -eq "Yes") {
    Set-Location $env:USERPROFILE     
    Remove-Item -Path $workdir -Recurse -Force > $null
}