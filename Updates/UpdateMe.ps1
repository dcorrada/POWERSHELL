<#
Name......: UpdateMe.ps1
Version...: 22.12.1
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Updates\\UpdateMe\.ps1$" > $null
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

# searching updates
$repo = 'PSGallery'
Write-Host -NoNewline "Looking for updates at [$repo]..."
$halloffame = Get-InstalledModule
$upgradable = @{}
foreach ($item in $halloffame) {
    $online = Find-Module -Name $item.name -Repository $repo -ErrorAction Stop
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
    if ($item.Version -ne $upgradable[$item.Name]) {
        $selmods[$item.Name] = CheckBox -form $adialog -checked $false -x 20 -y $they -text $desc
    } else {
        $selmods[$item.Name] = CheckBox -form $adialog -checked $false -enabled $false -x 20 -y $they -text $desc
    }
    $they += 30 
}
OKButton -form $adialog -x 150 -y ($they + 20) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

# perform instalation
foreach ($item in ($selmods.Keys | Sort-Object)) {
    Write-Host -NoNewline "Updating $item... "
    if ($selmods[$item].Checked) {
        $ErrorActionPreference= 'Stop'
        Try {
            Update-Module -Name $item -Force
            Write-Host -ForegroundColor Green 'OK'
            $ErrorActionPreference= 'Inquire'
        }
        Catch {
            Write-Host -ForegroundColor Red 'KO'
            Write-Output "Error: $($error[0].ToString())`n"
        }
    } else {
        Write-Host -ForegroundColor Yellow 'skipped'
    }
}



