<#
Name......: AGMConfMan_init.ps1
Version...: 26.06.1
Author....: Dario CORRADA

Questo script si occupa di installare e lanciare il servizio AGM_ConfigManager.
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
$serviceExists = $false
$ErrorActionPreference= 'Stop'
Try {
    $aService = Get-Service 'AGM_ConfigManager'
    $serviceExists = $true
    if ($aService.Status -cne 'Stopped') {
        Set-Service -InputObject $aService -Status 'Stopped'
    }
    Write-Host -ForegroundColor Green "Stopped"
    Write-Host -NoNewline "Looking for process... "
    $aProcess = Get-Process 'AGM_ConfigManager'
    Stop-Process -ID $aProcess.Id -Force
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
Write-Host -NoNewline "Looking for workdir..."
$aPath = "$($env:ProgramFiles)\AGM_ConfigManager"
if (!(Test-Path $aPath)) {
    New-Item -ItemType directory -Path $aPath | Out-Null
    Write-Host -NoNewline '.'
}
Write-Host -ForegroundColor Green ' Done'

# downloading resources
Write-Host -NoNewline "Downloading resources..."
$trgets = @{
    'msvbvm60.dll'          = "$($env:SystemRoot)\SysWOW64\"
    'NTSVC.ocx'             = "$($env:SystemRoot)\SysWOW64\"
    'AGM_ConfigManager.exe' = "$aPath\"
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
    $regsvr_codes = @{
        '1' = 'Invalid Argument'
        '2' = 'OleInitialize Failed'
        '3' = 'LoadLibrary Failed'
        '4' = 'GetProcAddress failed'
        '5' = 'DllRegisterServer failed'
    }
    Write-Host -NoNewline "Registering DLLs..."
    foreach ($aFilename in ('msvbvm60.dll', 'NTSVC.ocx')) {
        $aDestination = $trgets[$aFilename] + $aFilename
        $regsvrp = Start-Process regsvr32.exe -ArgumentList "/s $aDestination" -PassThru
        Start-Sleep -Milliseconds 1500
        if ($regsvrp.ExitCode -ne 0) {
            Write-Host -ForegroundColor Red "ERROR with [$aFilename]: regsvr32 ExitCode $($regsvrp.ExitCode) $($regsvr_codes[$regsvrp.ExitCode])"
            Write-Host -ForegroundColor Yellow "You should unregister libraries as follows: regsvr32 /u <dll/ocx_filename>"
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
    Start-Sleep -Milliseconds 2000
    Write-Host -ForegroundColor Green ' Done'

    # creating service
    $ErrorActionPreference= 'Stop'
    Try {
        Write-Host -NoNewline "Check service..."
        if (!$serviceExists) {
            $params = @{
                Name            = 'AGM_ConfigManager'
                BinaryPathName  = "$aPath\AGM_ConfigManager.exe"
                DisplayName     = "AGM Config Manager"
                StartupType     = "Automatic"
            }
            New-Service @params
            Start-Sleep -Milliseconds 1000
        }
        $sObject = Set-Service -Name AGM_ConfigManager -Status Running -StartupType Automatic -PassThru
        Write-Host -ForegroundColor Green ' Done'
    }
    Catch {
        # should investigate inside ServiceController object ($sObject) that displays the results?
        Write-Host -ForegroundColor Red "`nUnexpected error"
        Write-Host "$($error[0].ToString())`n"
        Pause
    }
    $ErrorActionPreference= 'Inquire'
}

<# *******************************************************************************
                                 PATCHES
******************************************************************************* #>
$regpath = 'HKLM:\SYSTEM\CurrentControlSet\Services\AGM_ConfigManager\Parameters'
if (!(Test-Path $regpath)) {
    New-Item -Path $regpath | Out-Null
}
$keyname = 'TenantID'
$ErrorActionPreference= 'SilentlyContinue'
try {
    $exists = Get-ItemProperty $regpath $keyname
    if ($exists.$keyname -ne '1') {
        Set-ItemProperty -Path $regpath -Name $keyname -Value '1'
    }
}
catch {
    New-ItemProperty -Path $regpath -Name $keyname -Value '1' -PropertyType String
}
$ErrorActionPreference= 'Inquire'

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