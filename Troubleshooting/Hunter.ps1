<#
Name......: Hunter.ps1
Version...: 21.11.1
Author....: Dario CORRADA

This script monitors for all the PIDs which belong to a named process.
Whenever any PID die, this script take a snapshot of the processes list before ([timestamp]_BC.csv) 
and after ([timestamp]_AD.csv) the killing event, plus a collection of the latest event logs ([timestamp]_EventLogs.csv).
#>

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
$WarningPreference = 'SilentlyContinue'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# temporary directory
$logpath = 'C:\HUNTER'
$timestamp = Get-Date -format "yyyyMMdd-HHmmss"
if (Test-Path $logpath) {
    $logpath += "\LOGS_$timestamp"
    New-Item -ItemType directory -Path $logpath > $null
} else {
    New-Item -ItemType directory -Path $logpath > $null
    $logpath += "\LOGS_$timestamp"
    New-Item -ItemType directory -Path $logpath > $null
}

# looking at the process to monitor
$prey = Read-Host "Enter the process name to monitor"
$rawdata = Get-WmiObject Win32_Process | where commandline -NE $null 
$pidofinterest = @()
foreach ($item in $rawdata) {
    if ($item.Name -match $prey) {
        $pid_found = $item.ProcessId
        $procname = $item.Name
        Write-Host -NoNewline -ForegroundColor Cyan "[$pid_found] "
        Write-Host -ForegroundColor Green "$procname"
        $pidofinterest += $pid_found
    }
}
if (!($pidofinterest)) {
    Write-Host -ForegroundColor Red "NO process <$prey> found. Aborting..."
    Pause
    Exit
}

# monitoring
Write-Host -ForegroundColor Blue "`nMonitoring <$prey> (press Ctrl+X to stop)..."
[console]::TreatControlCAsInput = $true
$eventlogs = @{}
$parseddata = @{}
$deepcopy = @{}
$memStream = New-Object -TypeName IO.MemoryStream
$formatter = New-Object -TypeName Runtime.Serialization.Formatters.Binary.BinaryFormatter
while ($true) {

    # grace period
    Start-Sleep -Milliseconds 500

    # shortcut for exiting the loop
    if ([console]::KeyAvailable) {
        $keypressed = [system.console]::readkey($true)
        if (($keypressed.modifiers -band [consolemodifiers]"control") -and ($keypressed.key -eq "X")) {
            Break
        }
    }

    # get a current snapshot of processes
    $parseddata.Clear()
    $rawdata = Get-WmiObject Win32_Process | where commandline -NE $null 
    foreach ($item in $rawdata) {
        $string = $item.CommandLine
        $parseddata[$item.ProcessID] = @{
            'PID' = $item.ProcessId
            'Name' = $item.Name
            'Cmd' = $string.replace(';','|')
            'Path' = $item.ExecutablePath
            'ParentPID' = $item.ParentProcessId
            'Time' = $item.UserModeTime
        }
    }

    # check for alive processes
    $pidalive = $false
    foreach ($item in $pidofinterest) {
        if ($parseddata.ContainsKey($item)) {
            $pidalive = $true
        }
    }

    if ($pidalive) {
        # backup a hashtable deepcopy
        $deepcopy.Clear()
        $formatter.Serialize($memStream, $parseddata)
        $memStream.Position = 0
        $deepcopy = $formatter.Deserialize($memStream)
    } else {
        # take a snapshot
        $timestamp = Get-Date -format "yyyy-MM-dd_HH-mm-ss-fff"
        Write-Host -ForegroundColor Red "[$timestamp] GOTCHA!"

        # saving logs
        Write-Host -NoNewline "Getting process list snapshot... "
        $logfile = $logpath + '\' + $timestamp + '_BC.csv'
        'PID;NAME;CMD;PATH;PARENT_PID;TIME' | Out-File $logfile -Encoding UTF8 -Append
        foreach ($item in ($deepcopy.Keys | Sort-Object)) {
            $new_record = @(
                $parseddata[$item].PID,
                $parseddata[$item].Name,
                $parseddata[$item].Cmd,
                $parseddata[$item].Path,
                $parseddata[$item].ParentPID,
                $parseddata[$item].Time
            )
            $new_string = [system.String]::Join(";", $new_record)
            $new_string | Out-File $logfile -Encoding UTF8 -Append
        }
        $logfile = $logpath + '\' + $timestamp + '_AD.csv' 
        'PID;NAME;CMD;PATH;PARENT_PID;TIME' | Out-File $logfile -Encoding UTF8 -Append
        foreach ($item in ($parseddata.Keys | Sort-Object)) {
            $new_record = @(
                $parseddata[$item].PID,
                $parseddata[$item].Name,
                $parseddata[$item].Cmd,
                $parseddata[$item].Path,
                $parseddata[$item].ParentPID,
                $parseddata[$item].Time
            )
            $new_string = [system.String]::Join(";", $new_record)
            $new_string | Out-File $logfile -Encoding UTF8 -Append
        }
        Write-Host -ForegroundColor Green 'DONE'

        # get events log
        Write-Host -NoNewline "Retrieving events log... "
        $logtime = (Get-Date).AddMinutes(-15)
        $eventlogs.Clear()
        $string = ''
        $ErrorActionPreference= 'SilentlyContinue'
        foreach ($logtype in ('System', 'Security','Application','OAlerts','Setup')) {
            $rawdata = Get-WinEvent -FilterHashTable @{LogName=$logtype; StartTime=$logtime}
            foreach ($record in $rawdata) {
                $logkey = '[' + $record.LogName + ']_'
                $logkey += Get-Date $record.TimeCreated -format "yyyy-MM-dd_HH-mm-ss-fff"
                $record.Message -match "^(.*)\r+" > $null
                if ($matches[1]) {
                    $string = $matches[1]
                    $matches[1] = $null
                } else {
                    $record.Message -match "^(.*)\n+" > $null
                    if ($matches[1]) {
                        $string = $matches[1]
                        $matches[1] = $null
                    } else {
                        $string = $record.Message
                    }
                }
                $eventlogs[$logkey] = @{
                    'Name' = $record.LogName
                    'Time' = Get-Date $record.TimeCreated -format "yyyy-MM-dd_HH-mm-ss"
                    'Id' = $record.Id
                    'Message' = $string
                    'Type' = $record.LevelDisplayName
                }
            }
        }
        $ErrorActionPreference= 'Inquire'
        $logfile = $logpath + '\' + $timestamp + '_EventLogs.csv' 
        'ID;LOGTYPE;NAME;TIME;MESSAGE' | Out-File $logfile -Encoding UTF8 -Append
        foreach ($item in ($eventlogs.Keys | Sort-Object)) {
            $new_record = @(
                $eventlogs[$item].Id,
                $eventlogs[$item].Type,
                $eventlogs[$item].Name,
                $eventlogs[$item].Time,
                $eventlogs[$item].Message  
            )
            $new_string = [system.String]::Join(";", $new_record)
            $new_string | Out-File $logfile -Encoding UTF8 -Append
        }
        Write-Host -ForegroundColor Green 'DONE'

        # waiting for process restart
        Write-Host -NoNewline "Waiting for process restarting... "
        $pidofinterest = @()
        while (!($pidofinterest)) {
            Start-Sleep -Milliseconds 500
            $rawdata = Get-WmiObject Win32_Process | where commandline -NE $null 
            foreach ($item in $rawdata) {
                if ($item.Name -match $prey) {
                    $pid_found = $item.ProcessId
                    $procname = $item.Name
                    #Write-Host -NoNewline -ForegroundColor Cyan "[$pid_found] "
                    #Write-Host -ForegroundColor Green "$procname"
                    $pidofinterest += $pid_found
                }
            }
        }
        Write-Host -ForegroundColor Green 'DONE'
    }
}

# view logs
$answ = [System.Windows.MessageBox]::Show("Monitoring finished. View log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Invoke-Item "$logpath"
}
