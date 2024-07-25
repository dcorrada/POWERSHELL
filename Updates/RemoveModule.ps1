<#
Name......: RemoveModule.ps1
Version...: 24.08.1
Author....: Dario CORRADA

This script looks for installed Powershell modules and try to update them
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Updates\\RemoveModule\.ps1$" > $null
$workdir = $matches[1]
<# for testing purposes
$workdir = Get-Location
$workdir = $workdir.Path
#>


# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# searching installed modules
Write-Host -NoNewline "Looking for installed modules..."
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
Write-Host -ForegroundColor Green ' DONE'

# show dialog
$adialog = FormBase -w 425 -h ((($halloffame.Count-1) * 30) + 175) -text "REMOVE MODULES"
$they = 20
$selmods = @{}
foreach ($item in $halloffame) {
    $desc = $item.Name + ' - ' + $item.Version
    $selmods[$desc] = CheckBox -form $adialog -checked $false -x 20 -y $they -w 350 -text $desc
    $they += 30
}
OKButton -form $adialog -x 150 -y ($they + 20) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

<# 
Cleaning previous installed versions, thx to Harm Veenstra
https://powershellisfun.com/2022/07/11/updating-your-powershell-modules-to-the-latest-version-plus-cleaning-up-older-versions/

For uninstall of Microsoft.Graph and related submodules, thx to Andres Bohren
https://blog.icewolf.ch/archive/2022/03/18/cleanup-microsoft-graph-powershell-modules/
#>

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

Start-Sleep -Milliseconds 1000
