<#
Name......: SafetyScan.ps1
Version...: 22.04.1
Author....: Dario CORRADA

This script automatically fetch and performs a scan with Microsoft Safety Scanner Tool
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

Write-Host -ForegroundColor Blue "*** MS Safety Scanner ***`n"

# function for checking if a proscess is alive
function ProCheck {
    param ($procid)
    
    Write-Host -NoNewline -ForegroundColor Gray "`nSCANNING (press Ctrl+X to abort)... "
    [console]::TreatControlCAsInput = $true
    $pidalive = $true
    while ($pidalive -eq $true) {
        # grace period
        Start-Sleep -Milliseconds 500

        # aborting scan
        if ([console]::KeyAvailable) {
            $keypressed = [system.console]::readkey($true)
            if (($keypressed.modifiers -band [consolemodifiers]"control") -and ($keypressed.key -eq "X")) {
                Stop-Process -Id $procid -Force
                Start-Sleep 3
                Write-Host -ForegroundColor Red 'ABORTED'
                Break
            }
        }

        # check if scan is alive
        $pidalive = $false
        $pids = Get-WmiObject Win32_Process | where commandline -NE $null 
        foreach ($item in $pids) {
            if ($item.ProcessID -eq $procid) {
                $pidalive = $true
            }
        }
    }
    if ($pidalive -eq $false) {
        Write-Host -ForegroundColor Green 'DONE'
    }
}

# remove previous log file
$logfile = 'C:\Windows\debug\msert.log'
if (Test-Path $logfile) {
    Remove-Item $logfile -Force    
}

# Download
Write-Host -NoNewline 'Download... '
$download = New-Object net.webclient
$downbin = 'C:\Users\' + $env:USERNAME + '\Downloads\MSERT.exe'
if (Test-Path $downbin) {
    Remove-Item $downbin -Force    
}
# also available at http://definitionupdates.microsoft.com/download/definitionupdates/safetyscanner/amd64/msert.exe
$download.DownloadFile('https://go.microsoft.com/fwlink/?LinkId=212732', $downbin)
#Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?LinkId=212732' -OutFile $downbin
Write-Host -ForegroundColor Green "DONE`n"

<#
SYNOPSIS:
/Q      quiet mode, no UI is shown
/N      detect-only mode
/F      force full scan
/F:Y    force full scan and automatic clean
/H      detect high and severe threats only
#>
# perform a quick scan
Write-Host -NoNewline 'Perform QUICK scan... '
$process = Start-Process $downbin '/N /Q' -PassThru
Write-Host -ForegroundColor Green "STARTED"
ProCheck $process.Id

# log check
Start-Process -Wait notepad $logfile

# perform a full scan
$answ = [System.Windows.MessageBox]::Show("Do you need a full scan?",'FULL-SCAN','YesNo','Info')
if ($answ -eq "Yes") {    
    Write-Host -NoNewline 'Perform FULL scan... '
    $process = Start-Process $downbin  '/F:Y /Q' -PassThru
    Write-Host -ForegroundColor Green "STARTED"
    ProCheck $process.Id

    # log check
    notepad $logfile
}

