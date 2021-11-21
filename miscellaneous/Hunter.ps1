<#
Name......: Hunter.ps1
Version...: 21.11.1
Author....: Dario CORRADA

This script monitor the status of running processes
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Hunter\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# temporary directory
$logpath = 'C:\HUNTER'
$currently = Get-Date -format "yyyyMMdd-HHmmss"
if (Test-Path $logpath) {
    $logpath += "\LOGS_$currently"
    New-Item -ItemType directory -Path $logpath > $null
} else {
    New-Item -ItemType directory -Path $logpath > $null
    $logpath += "\LOGS_$currently"
    New-Item -ItemType directory -Path $logpath > $null
}

# looking at the process to monitor
$prey = Read-Host "Enter the process name to monitor"
$rawdata = Get-WmiObject Win32_Process | where commandline -NE $null 
$catched = $false
foreach ($item in $rawdata) {
    if ($item.Name -match $prey) {
        $pid_found = $item.ProcessId
        $procname = $item.Name
        Write-Host -NoNewline -ForegroundColor Cyan "[$pid_found] "
        Write-Host -ForegroundColor Green "$procname"
        $catched = $true
    }
}
if (!($catched)) {
    Write-Host -ForegroundColor Red "NO process <$prey> found. Aborting..."
    Pause
    Exit
}

# monitoring
Write-Host -NoNewline "Monitoring <$prey> (press Ctrl+X to stop)..."
[console]::TreatControlCAsInput = $true
while ($true) {
    Start-Sleep -Milliseconds 500
    if ([console]::KeyAvailable) {
        $keypressed = [system.console]::readkey($true)
        if (($keypressed.modifiers -band [consolemodifiers]"control") -and ($keypressed.key -eq "X")) {
            Write-Host -ForegroundColor Green " DONE"
            Break
        }
    }
}




$parseddata = @{}
$rawdata = Get-WmiObject Win32_Process | where commandline -NE $null 
foreach ($item in $rawdata) {
    $parseddata["$item.ProcessID"] = @{
        'PID' = $item.ProcessId
        'Name' = $item.Name
        'Cmd' = $item.CommandLine
        'CreationDate' = $item.CreationDate
        'Path' = $item.ExecutablePath
        'ParentPID' = $item.ParentProcessId
        'Time' = $item.UserModeTime
    }
}





# view logs
$answ = [System.Windows.MessageBox]::Show("Monitoring finished. View log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Invoke-Item "$logpath"
}
