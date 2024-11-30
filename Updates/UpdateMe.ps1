<#
Name......: UpdateMe.ps1
Version...: 24.11.1
Author....: Dario CORRADA

This script looks for installed Powershell modules and try to update them
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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"


<# *******************************************************************************
                                    BODY
******************************************************************************* #>
$repo = 'PSGallery'

# PowershellGet, on Windows11, seems stuck on version 1.0.0.1 and doesn't work
$ErrorActionPreference= 'Stop'
try {
    $vinstalled = Get-InstalledModule -Name PowershellGet
}
catch {
    Install-Module PowershellGet -Force
}
$ErrorActionPreference= 'Inquire'

# Package Management preliminar check
# https://stackoverflow.com/questions/66305351/powershell-unable-to-update-powershellget-error-the-version-1-4-7-of-modul
$vonline = Find-Module -Name PowershellGet -Repository $repo

if ($vinstalled.Version -ne $vonline.Version) {
    Write-Host -NoNewline "Updating Package Management required..."
    $ErrorActionPreference= 'Stop'
    try {
        Update-Module -Name PowerShellGet -Force
        Write-Host -ForegroundColor Green 'OK'
        [System.Windows.MessageBox]::Show("Package Management updated, please restart it`nClick Ok to close the script",'RESTART','Ok','warning') | Out-Null
        exit
    }
    catch {
        Write-Host -ForegroundColor Red 'KO'
        Write-Output "Error: $($error[0].ToString())`n"
        Pause
    }
    $ErrorActionPreference= 'Inquire'
}

# searching updates
Write-Host -NoNewline "Looking for updates at [$repo]..."
$halloffame = Get-InstalledModule | Foreach-Object{
    Write-Host -NoNewline '.' 
    if ($_.Name -cnotmatch "^Microsoft\.Graph\..+$") { # bypass Graph submodules
        New-Object -TypeName PSObject -Property @{
            Version     = $_.Version
            Name        = $_.Name
            Repository  = $_.Repository
            Description = $_.Description
        } | Select Version, Name
    }
}
$upgradable = @{}
foreach ($item in $halloffame) {
    $online = Find-Module -Name $item.Name -Repository $repo -ErrorAction Stop
    Write-Host -NoNewline '.'
    $upgradable[$online.Name] = $online.Version
}
Write-Host -ForegroundColor Green ' DONE'

# show dialog
$adialog = FormBase -w 425 -h ((($halloffame.Count-1) * 30) + 175) -text "INSTALLED MODULES"
$they = 20
$selmods = @{}
foreach ($item in $halloffame) {
    $desc = $item.Name + ' v' + $item.Version + ' >>> v' + $upgradable[$item.Name]
    if (($item.Name -ceq 'PackageManagement') -or ($item.Name -ceq 'PowerShellGet')) {
        $selmods[$item.Name] = CheckBox -form $adialog -checked $false -enabled $false -x 20 -y $they -w 350 -text $desc
    } elseif ($item.Version -ne $upgradable[$item.Name]) {
        $selmods[$item.Name] = CheckBox -form $adialog -checked $false -x 20 -y $they -w 350 -text $desc
    } else {
        $selmods[$item.Name] = CheckBox -form $adialog -checked $false -enabled $false -x 20 -y $they -w 350 -text $desc
    }
    $they += 30
}
OKButton -form $adialog -x 150 -y ($they + 20) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

# perform installation
$UpdatedItems = 0
foreach ($item in ($selmods.Keys | Sort-Object)) {
    Write-Host -NoNewline "Updating $item... "
    if ($selmods[$item].Checked) {
        $ErrorActionPreference= 'Stop'
        Try {
            Update-Module -Name $item -Force
            Write-Host -ForegroundColor Green 'OK'
            $UpdatedItems++
            $ErrorActionPreference= 'Inquire'
        }
        Catch {
            Write-Host -ForegroundColor Red 'KO'
            Write-Output "Error: $($error[0].ToString())`n"
            Pause
        }
    } else {
        Write-Host -ForegroundColor Yellow 'skipped'
    }
}

if ($UpdatedItems -gt 0) {
    <# 
    Cleaning previous installed versions, thx to Harm Veenstra
    https://powershellisfun.com/2022/07/11/updating-your-powershell-modules-to-the-latest-version-plus-cleaning-up-older-versions/

    For uninstall of Microsoft.Graph and related submodules, thx to Andres Bohren
    https://blog.icewolf.ch/archive/2022/03/18/cleanup-microsoft-graph-powershell-modules/
    #>
    $answ = [System.Windows.MessageBox]::Show("Clean previous installed versions?",'REMOVE','YesNo','Info')
    if ($answ -eq "Yes") {
        
        Write-Host -NoNewline "Checking previous versions..."
        $hallofshame = @()
        foreach ($item in $halloffame) {
            $previous = Get-InstalledModule -Name $item.Name -AllVersions | Sort-Object PublishedDate -Descending
            $current = $previous[0].Version.ToString() # latest version installed
            Write-Host -NoNewline "."
            if ($previous.Count -gt 1) {
                Write-Host -NoNewline "."
                for ($i = 1; $i -lt $previous.Count; $i++) {
                    $old = $item.Name + ' - ' + $previous[$i].Version
                    $hallofshame += $old
                }
            }
        }
        Write-Host -ForegroundColor Green " DONE"

        $adialog = FormBase -w 425 -h ((($hallofshame.Count+1) * 30) + 125) -text "PREVIOUS INSTALLED"
        Label -form $adialog -x 20 -y 20 -w 300 -h 30 -text 'Would you uninstall previous version(s)?' | Out-Null
        $they = 50
        $selmods = @{}
        foreach ($item in $hallofshame) {
            $selmods[$item] = CheckBox -form $adialog -checked $false -w 350 -x 20 -y $they -text $item
            $they += 30
        }
        OKButton -form $adialog -x 150 -y ($they + 20) -text "Ok" | Out-Null
        $result = $adialog.ShowDialog()

        $subitems = @()
        foreach ($item in ($selmods.Keys | Sort-Object)) {
            if ($selmods[$item].Checked) {
                if ($item -match 'Microsoft.Graph') {
                    $item -match "^([a-zA-Z_\-\.0-9]+) - ([0-9\.]+)$" > $null
                    $ReqVer = $matches[2]
                    foreach ($submodule in ((Get-Module Microsoft.Graph* -ListAvailable | Select-Object Name -Unique) | Sort-Object)) {
                        $string = $submodule.Name + ' - ' + [string]$ReqVer
                        $subitems += $string
                    }
                } else {
                    $subitems += $item
                }
            }
        }

        $ErrorActionPreference= 'Stop'
        $thelastone = @()
        foreach ($subitem in ($subitems | Sort-Object)) {
            # keep aside Authentication submodule and the parent, queue them as the lastest to be removed
            if (($subitem -match "^Microsoft.Graph.Authentication - ") -or ($subitem -match "^Microsoft.Graph - ")) {
                $thelastone += $subitem
            } else {
                Write-Host -NoNewline "Uninstalling $subitem... "
                Try {
                    $subitem -match "^([a-zA-Z_\-\.0-9]+) - ([0-9\.]+)$" > $null
                    Uninstall-Module -Name $matches[1] -RequiredVersion $matches[2] -Force:$True -ErrorAction Stop
                    $matches = @()
                    Write-Host -ForegroundColor Green 'OK'
                    $ErrorActionPreference= 'Inquire'
                }
                Catch {
                    Write-Host -ForegroundColor Red 'KO'
                    Write-Output "Error: $($error[0].ToString())`n"
                    Pause
                }
            }
        }
        if ($thelastone.Count -gt 0) {
            foreach ($currentItemName in ($thelastone | Sort-Object -Descending)) {
                Write-Host -NoNewline "Uninstalling $currentItemName... "
                Try {
                    $currentItemName -match "^([a-zA-Z_\-\.0-9]+) - ([0-9\.]+)$" > $null
                    Uninstall-Module -Name $matches[1] -RequiredVersion $matches[2] -Force:$True -ErrorAction Stop
                    $matches = @()
                    Write-Host -ForegroundColor Green 'OK'
                    $ErrorActionPreference= 'Inquire'
                }
                Catch {
                    Write-Host -ForegroundColor Red 'KO'
                    Write-Output "Error: $($error[0].ToString())`n"
                    Pause
                }
            }
        }
    }
}