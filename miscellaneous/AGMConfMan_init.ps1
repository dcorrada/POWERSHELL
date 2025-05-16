
<#
Name......: AGMConfMan_init.ps1
Version...: 25.05.1
Author....: Dario CORRADA

Questo script si occupa di installare e lanciare i servizio AGM_ConfigManager.

+++ TO DO +++
* integrare e lanciare dalla pipeline PPPC
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


<# *******************************************************************************
                                    INIT
******************************************************************************* #>
# check the existence of target regpath
Write-Host -NoNewline "Check registry path... "
if (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Services\AGM_ConfigManager') {
    Write-Host -ForegroundColor Green "Found"
    $SkipInstallation = $true
} else {
    Write-Host -ForegroundColor Yellow "Undef"
    $SkipInstallation = $false
}

# check if the (previous version of the) service is currently active
Write-Host -NoNewline "Looking for service... "
$ErrorActionPreference= 'Stop'
Try {
    $aService = Get-Service 'AGM_ConfigManager'
    if ($aService.Status -cne 'Stopped') {
        Set-Service -InputObject $aService -Status 'Stopped'
    }
    Write-Host -ForegroundColor Green "Stopped"
    Write-Host -NoNewline "Looking for process... "
    $aProcess = Get-Process 'AGM_ConfigManager'
    Stop-Process -ID $outproc.Id -Force
    Start-Sleep -Milliseconds 1500
    Write-Host -ForegroundColor Green 'Killed'
    $SkipInstallation = $false
}
Catch {
    if ($error[0].InvocationInfo.InvocationName -eq 'Get-Service') { # No service found
        Write-Host -ForegroundColor Yellow "Undef"
        $SkipInstallation = $false
    } elseif ($error[0].InvocationInfo.InvocationName -eq 'Set-Service') { # Service not stopped
        Write-Host -ForegroundColor Yellow "Not stopped"
        $SkipInstallation = $true
    } elseif ($error[0].InvocationInfo.InvocationName -eq 'Get-Process') { # No process found
        Write-Host -ForegroundColor Yellow "Undef"
        $SkipInstallation = $false
    } elseif ($error[0].InvocationInfo.InvocationName -eq 'Stop-Process') { # Process alive
        Write-Host -ForegroundColor Yellow "Alive"
        $SkipInstallation = $true
    } else {
        Write-Host -ForegroundColor Red "`nUnexpected error"
        Write-Host "$($error[0].ToString())`n"
        Pause
        Exit
    }

}
$ErrorActionPreference= 'Inquire'

# create workdirs if not exists
Write-Host -NoNewline "Looking for workdirs..."
foreach ($aPath in ('C:\Program Files\AGM_ConfigManager', 'C:\Windows\SysWOW64\AGM_ConfigManager')) {
    if (!(Test-Path $aPath)) {
        New-Item -ItemType directory -Path $aPath | Out-Null
        Write-Host -NoNewline'.'
    }
}
Write-Host -ForegroundColor Green ' Done'

# downloading resources
Write-Host -NoNewline "Downloading resources..."
$trgets = @{
    'msvbvm60.dll'          = 'C:\Windows\SysWOW64\AGM_ConfigManager\'
    'NTSVC.ocx'             = 'C:\Windows\SysWOW64\AGM_ConfigManager\'
    'AGM_ConfigManager.exe' = 'C:\Program Files\AGM_ConfigManager\'
}
$aUrl = 'https://cm.agmsolutions.net:60443/agent_files/gpo_deploy/'
$ErrorActionPreference= 'Stop'
Try {
    foreach ($aFilename in $trgets.Keys) {
        $aDestination = $trgets[$aFilename] + $aFilename
        if (!(Test-Path $aDestination -PathType Leaf)) {
            $aSource = $aUrl + $aFilename
            Invoke-WebRequest -Uri "$aSource" -OutFile "$aDestination"
            Write-Host -NoNewline '.'
        }
    }
    Write-Host -ForegroundColor Green ' Done'
}
Catch {
    Write-Host -ForegroundColor Red "`nUnexpected error"
    Write-Host "$($error[0].ToString())`n"
    Pause
    Exit
}
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                    INSTALL
******************************************************************************* #>
if ($SkipInstallation) {
    [System.Windows.MessageBox]::Show("Installation skipped",'ABORT','Ok','Warning') | Out-Null
} else {
    Write-Host -ForegroundColor Cyan "`n*** INSTALL ***"

    # registering libraries, see:
    # https://stackoverflow.com/questions/37110533/powershell-to-display-regsvr32-result-in-console-instead-of-dialog
    Write-Host -NoNewline "Registering DLLs..."
    foreach ($aFilename in $trgets.Keys) {
        $aDestination = $trgets[$aFilename] + $aFilename
        $regsvrp = Start-Process regsvr32.exe -ArgumentList "/s $aDestination" -PassThru
        $regsvrp.WaitForExit(1000) # Wait (up to) 1 second
        if($regsvrp.ExitCode -ne 0) {
            Write-Host -ForegroundColor Red "regsvr32 exited with error $($regsvrp.ExitCode)"
            Pause
            Exit
        } else {
            Write-Host -NoNewline '.'
        }
    }
    Write-Host -ForegroundColor Green ' Done'

    # launching installer
    Write-Host -NoNewline "Launching AGM Config Manager..."
    $aInstaller = Start-Process 'C:\Program Files\AGM_ConfigManager\AGM_ConfigManager.exe' -ArgumentList '-I' -PassThru
    $aInstaller.WaitForExit(2000)
    Write-Host -ForegroundColor Green ' Done'

    # creating service
    $ErrorActionPreference= 'Stop'
    try {
        Write-Host -NoNewline "Creating service..."
        $params = @{
            Name            = 'AGM_ConfigManager'
            BinaryPathName  = 'C:\Program Files\AGM_ConfigManager\AGM_ConfigManager.exe'
            DisplayName     = "AGM Config Manager"
            StartupType     = "Automatic"
        }
        New-Service @params
        Set-Service -Name AGM_ConfigManager -Status Running -PassThru
        Write-Host -ForegroundColor Green ' Done'
    }
    catch {
        Write-Host -ForegroundColor Red "`nUnexpected error"
        Write-Host "$($error[0].ToString())`n"
        Pause
    }
    $ErrorActionPreference= 'Inquire'
}

<#
+++ original script +++
Name......: deploy.bat
Author....: Stefano VAILATI

@echo off

reg query HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AGM_ConfigManager
if "%ERRORLEVEL%"=="0" goto skip_installation

sc stop "AGM_ConfigManager"
TASKKILL /F /IM AGM_ConfigManager.exe
if not exist "C:\Program Files\AGM_ConfigManager\" mkdir "C:\Program Files\AGM_ConfigManager"

ECHO Option Explicit > %0\..\tempscr.vbs
ECHO On Error Resume Next >> %0\..\tempscr.vbs
ECHO Sub FILE_Download(ByVal srcFILE, ByVal destPATH) >> %0\..\tempscr.vbs
ECHO Dim objHTTP, objSTREAM >> %0\..\tempscr.vbs
ECHO On Error Resume Next >> %0\..\tempscr.vbs
ECHO Set objHTTP = createobject("Microsoft.XMLHTTP") >> %0\..\tempscr.vbs
ECHO Set objSTREAM = createobject("Adodb.Stream") >> %0\..\tempscr.vbs
ECHO objHTTP.Open "GET", "https://cm.agmsolutions.net:60443/agent_files/gpo_deploy/" ^& srcFILE, False >> %0\..\tempscr.vbs
ECHO objHTTP.Send >> %0\..\tempscr.vbs
ECHO objSTREAM.type = 1 >> %0\..\tempscr.vbs
ECHO objSTREAM.open >> %0\..\tempscr.vbs
ECHO objSTREAM.write objHTTP.responseBody >> %0\..\tempscr.vbs
ECHO objSTREAM.savetofile destPATH, 2 >> %0\..\tempscr.vbs
ECHO End Sub >> %0\..\tempscr.vbs
ECHO Call FILE_Download("msvbvm60.dll", "C:\Windows\SysWOW64\AGM_ConfigManager\msvbvm60.dll") >> %0\..\tempscr.vbs
ECHO Call FILE_Download("NTSVC.ocx", "C:\Windows\SysWOW64\AGM_ConfigManager\NTSVC.ocx") >> %0\..\tempscr.vbs
ECHO Call FILE_Download("AGM_ConfigManager.exe", "C:\Program Files\AGM_ConfigManager\AGM_ConfigManager.exe") >> %0\..\tempscr.vbs
CSCRIPT %0\..\tempscr.vbs
DEL /Q %0\..\tempscr.vbs

regsvr32 /s "%SystemRoot%\SysWow64\msvbvm60.dll"
regsvr32 /s "%SystemRoot%\SysWow64\ntsvc.ocx"
"C:\Program Files\AGM_ConfigManager\AGM_ConfigManager.exe -I"
sc create AGM_ConfigManager binPath="C:\Program Files\AGM_ConfigManager\AGM_ConfigManager.exe" DisplayName="AGM Config Manager" start=auto
sc start "AGM_ConfigManager"
goto end_installation

:skip_installation
echo Skipping, already installed.
goto end_installation


:end_installation
#>