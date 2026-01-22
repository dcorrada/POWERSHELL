<#
Name......: UninstallGraph.ps1
Version...: 26.1.1
Author....: Dario CORRADA

This script uninstall Microsoft.Graph module
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

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework


<# *******************************************************************************
                                    BODY
******************************************************************************* #>

# collecting submodules
Write-Host -NoNewline "Looking for submodules..."
$submodules = ((Get-Module Microsoft.Graph* -ListAvailable | Select-Object Name -Unique) | Sort-Object)
Write-Host -ForegroundColor Green 'Done'

$ErrorActionPreference= 'Stop'
$thelastone = @()
foreach ($subitem in $submodules) {
    # keep aside Authentication submodule and the parent, queue them as the lastest to be removed
    if (($subitem -match "^Microsoft.Graph.Authentication") -or ($subitem -match "^Microsoft.Graph")) {
        $thelastone += $subitem
    } else {
        Write-Host -NoNewline "Uninstalling $($subitem.Name)... "
        Try {
            Uninstall-Module -Name $subitem.Name -Force:$True -ErrorAction Stop
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
    foreach ($subitem in ($thelastone | Sort-Object -Descending)) {
        Write-Host -NoNewline "Uninstalling $currentItemName... "
        Try {
            Uninstall-Module -Name $subitem.Name -Force:$True -ErrorAction Stop
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

[System.Windows.MessageBox]::Show(@"
Module Microsoft.Graph has been removed. In order to 
reinstall such module enter the following commandline:

Install-Module Microsoft.Graph [-RequiredVersion x.yy.z] -Force
"@,'REINSTALL','Ok','Info') | Out-Null
