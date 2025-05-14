
<#
Name......: AGMConfMan_init.ps1
Version...: 25.05.1
Author....: Dario CORRADA
Credits...: Stefano VAILATI

Questo script si occupa di installare e lanciare i servizio AGM_ConfigManager.

Da integrare e lanciare dalla pipeline PPPC
#>


<#
Qui di seguito la traccia dello script di Stefano in versione batch file, 
che crea internamente e richiama uno script vbs.

*** Note di sviluppo ***
Il comando 'sc' serve per creare e gestire l'esecuzione di servizi
https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/cc754599(v=ws.11)

Esiste il cmdlet equivalente 'Set-Service'
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/set-service?view=powershell-5.1
#>

<#
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