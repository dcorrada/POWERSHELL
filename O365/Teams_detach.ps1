<#
Name......: Teams_detach.ps1
Version...: 22.12.1
Author....: Dario CORRADA

This script detach any account from Teams desktop app 

See also:
https://learn.microsoft.com/en-us/answers/questions/774931/a-powershell-command-to-disconnect-work-or-school.html
https://answers.microsoft.com/it-it/msteams/forum/all/errore-80090016-di-accesso-a-teams/101e5def-75a7-4dd6-acba-87a08bacf7ca
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\O365$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"

# operations dialog
$adialog = FormBase -w 275 -h 180 -text "CHOOSE"
$clean = RadioButton -form $adialog -x 20 -y 20 -checked $true -text 'Clean cache'
$detach = RadioButton -form $adialog -x 20 -y 50 -checked $false -text 'Detach account(s)'
OKButton -form $adialog -x 75 -y 90 -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

# killing Teams
$ErrorActionPreference= 'SilentlyContinue'
$outproc = Get-Process Teams
if ($outproc -ne $null) {
    $ErrorActionPreference= 'Stop'
    Try {
        Stop-Process -ID $outproc.Id -Force
        Start-Sleep 2
    }
    Catch { 
        [System.Windows.MessageBox]::Show("Check out that all Teams processes have been closed before go ahead",'TASK MANAGER','Ok','Warning') > $null
    }
}
$ErrorActionPreference= 'Inquire'

# perform operations
if ($detach.Checked) {
    $apath = $env:USERPROFILE + '\AppData\Local\Packages'
    foreach ($item in Get-ChildItem $apath) {
        if ($item.Name -match "Microsoft\.AAD.BrokerPlugin.*$") {
            $ErrorActionPreference= 'Stop'
            Try {
                Remove-Item -Path $item.FullName -Recurse -Force
            }
            Catch { 
                [System.Windows.MessageBox]::Show("$($error[0].ToString())",'ERROR','Ok','Error') > $null
            }
            $ErrorActionPreference= 'Inquire'
        }
    } 
} elseif ($clean.Checked) {
    $apath = $env:USERPROFILE + '\AppData\Roaming\Microsoft\teams\'
    foreach ($item in ('application cache\cache', 'blob_storage', 'databases', 'cache', 'gpucache', 'Indexeddb', 'Local Storage', 'tmp')) {
        $targetpath = $apath + $item
        if (Test-Path $targetpath) {
            $ErrorActionPreference= 'Stop'
            Try {
                Remove-Item -Path $targetpath -Recurse -Force
            }
            Catch { 
                [System.Windows.MessageBox]::Show("$($error[0].ToString())",'ERROR','Ok','Error') > $null
            }
            $ErrorActionPreference= 'Inquire'
        } else {
            [System.Windows.MessageBox]::Show("[$targetpath] not found",'WARNING','Ok','Warning') > $null
        }
    }
    $answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
    if ($answ -eq "Yes") {    
        Restart-Computer
    }
}

