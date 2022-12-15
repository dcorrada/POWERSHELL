<#
Name......: Teams_detach.ps1
Version...: 22.12.1
Author....: Dario CORRADA

This script detach any account from Teams desktop app 

See also:
https://learn.microsoft.com/en-us/answers/questions/774931/a-powershell-command-to-disconnect-work-or-school.html
https://answers.microsoft.com/it-it/msteams/forum/all/errore-80090016-di-accesso-a-teams/101e5def-75a7-4dd6-acba-87a08bacf7ca

NOTES:
221215  After running the script the subsequent reload of Teams proposed the usual login
        dialog. Once logged in an error message occurred*, but it was a false positive and 
        don't know if related with this script. Further investigations needed...

[*] https://answers.microsoft.com/en-us/msoffice/forum/all/keyset-does-not-exist-tpm/de690cea-bba8-4260-8985-872e136e76c2
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\O365\\Teams_detach\.ps1$" > $null
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

