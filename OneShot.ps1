<#
Name......: OneShot.ps1
Version...: 24.08.1
Author....: Dario CORRADA

This script allow to navigate and select single scripts from this repository.
Then it dowload them and launch them locally.
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
        Set-GitHubConfiguration -DisableTelemetry -SessionOnly
        $ThirdParty = 'Ok'
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
                                    BROWSING
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

    TxtBox -form $adialog -x 230 -y 20 -w 450 -h 200 -text $intoBox -multiline $true | Out-Null
    
    $PreviousBut = RETRYButton -form $adialog -x 230 -y 230 -w 75 -text "UP"
    $PreviousBut.DialogResult = [System.Windows.Forms.DialogResult]::CANCEL
    $NextBut = OKButton -form $adialog -x 305 -y 230 -w 75 -text "GO"
    $PreviewBut = RETRYButton -form $adialog -x 580 -y 230 -text "Preview"
    
    Label -form $adialog -x 230 -y 280 -h 25 -text "RECENT LAUNCHES:" | Out-Null
    $they = 300
    foreach ($cachedItem in $cachedItems) {
        $choices += RadioButton -form $adialog -x 240 -y $they -checked $false -text $cachedItem.NAME
        $they += 25
        $CurrentItems += $cachedItem
    }

    $goahead = $adialog.ShowDialog()

    foreach ($currentOpt in $choices) {
        if ($currentOpt.Checked) {
            $isChecked = "$($currentOpt.Text)"
            foreach ($anItem in $CurrentItems) {
                if ($anItem.NAME -eq $isChecked) {
                    if ($goahead -eq 'RETRY') {
                        if ([string]::IsNullOrEmpty($anItem.URL)) {
                            $arrayContent = @("Items in folder [$($currentOpt.Text)]:", "")
                            foreach ($entry in (Get-GitHubContent `
                                -OwnerName 'dcorrada' `
                                -RepositoryName 'POWERSHELL' `
                                -Path $anItem.PATH
                                ).entries) {
                                    $arrayContent += "  * $($entry.name)"
                                }
                            $intoBox = $arrayContent | Out-String
                        } else {
                            $fileContent = (Invoke-WebRequest -Uri $anItem.URL -UseBasicParsing).Content.Split("`n")
                            $maxlines = 19
                            if ($fileContent.Count -lt $maxlines) {
                                $maxlines = $fileContent.Count - 1
                            }
                            $intoBox = $fileContent[0..$maxlines] + "[...]" | Out-String
                        }
                    } elseif ($goahead -eq 'OK') {
                        if ([string]::IsNullOrEmpty($anItem.URL)) {
                            # it's a folder, navigate it
                            $isChecked = 'none'
                            $currentFolder = '/' + $anItem.PATH
                            $CurrentItems = @()
                            $CurrentItems = (Get-GitHubContent `
                                -OwnerName 'dcorrada' `
                                -RepositoryName 'POWERSHELL' `
                                -Path $anItem.PATH
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
                            # selected a script, exit dialog
                            $selectedItem = $anItem
                            $continueBrowsing = $false
                        }
                    } elseif ($goahead -eq 'CANCEL') {
                        $isChecked = 'none'
                        if ($currentFolder -ne 'root') {
                            if ($currentFolder -match "(^/.+)/[a-zA-Z_\-\.0-9]+$") {
                                $currentFolder = $matches[1]
                                $CurrentItems = @()
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
                                $CurrentItems = @()
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
                        } else {
                            [System.Windows.MessageBox]::Show("You are already at root",'ROOT','Ok','info') > $null
                        }
                    }
                }
            }
        }
    }
}


<# *******************************************************************************
                                    RUNNING
******************************************************************************* #>
Write-Host -ForegroundColor Yellow @"
*** SELECTED ITEM ***
NAME...: $($selectedItem.NAME)
PATH...: $($selectedItem.PATH)
URL....: $($selectedItem.URL)
"@
<# 
fare check aggiuntivo sullo script selezionato:
  a) caricare PSWallet e Stargate se viene richiamato il wallet
  b) caricare PSWallet e AppKeyring per gli script in Graph
#>

# update history
if ($cachedItems.NAME -cnotcontains $selectedItem.NAME) {
    $cachedItems += New-Object -TypeName PSObject -Property @{
        NAME    = $selectedItem.NAME
        PATH    = $selectedItem.PATH
        URL     = $selectedItem.URL
    } | Select NAME, PATH, URL
}
$cachedItems[($cachedItems.Count - 5)..($cachedItems.Count - 1)] | Export-Csv -Path $cacheFile -NoTypeInformation

