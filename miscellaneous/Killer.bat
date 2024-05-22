<# : Batch portion
@rem # The previous line does nothing in Batch, but begins a multiline comment block
@rem # in PowerShell.  This allows a single script to be executed by both interpreters.
@echo off
cls & color 9E & Mode 95,5
Title Running Processes - Scheduled Tasks - Services - Startup items by Hackoo 2021
If [%1] NEQ [Admin] Goto RunAsAdmin

echo(
echo(                ===========================================================
echo(                    Please wait a while ... Working is in progress....
echo(                ===========================================================

Set "Filter_Ext=%Temp%\Filter_Ext"
Call :GetFileNameWithDateTime MyDate
Set "Log=%~dpn0_%Computername%_%MyDate%.txt"
Set "Lnk_Target_Path_Log=%~dp0Lnk_Target_Path_Log.txt"
Set "All_Users=%ProgramData%\Microsoft\Windows\Start Menu\Programs\Startup"
Set "Current_User=%UserProfile%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
set "Winlogonkey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set "ImageFileExec_Key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
Set StartupFolders="%All_Users%" "%Current_User%"
If Exist "%Log%" Del "%Log%"
Set "VbsFile=%Tmp%\%~n0.vbs"
Call :Generate_VBS_File

  Powershell ^
  Get-WmiObject Win32_Process ^
| where commandline -NE $null ^
| Select-Object ProcessID,Name,CommandLine ^
| Out-String -Width 450 ^
| Findstr /I /V "Admin" ^
| Findstr /I /V "Get-WmiObject" ^
| Out-File "%Log%" -Encoding  ASCII 

  Powershell ^
  Get-CimInstance Win32_StartupCommand ^
| Select-Object Name,command,Location,user ^
| Format-List ^
| Out-File -Append "%Log%" -Encoding  ASCII

>"%Lnk_Target_Path_Log%" (
	@For %%A in (%StartupFolders%) Do (
		Call :Execute_VBS_File "%%~A"
	)
)

>> "%Log%" (Type "%Lnk_Target_Path_Log%")

> "%Filter_Ext%" (
	echo .vbs
	echo .vbe
	echo .js
	echo .jse
	echo .bat
	echo .cmd
	echo .ps1
)

@for /f "delims=" %%a in ('Type "%Lnk_Target_Path_Log%" ^| Findstr /I /G:"%Filter_Ext%"') do (
	@for /f "tokens=2 delims==" %%b in ('echo %%a') do (
		>> "%Log%" 2>&1 (
			echo(
			echo ===================================================================================
			echo( Source code of TargetPath=%%b
			echo ===================================================================================
			Type %%b
		)
	)
)

Del "%Filter_Ext%" /F >nul 2>&1
Del "%Lnk_Target_Path_Log%" >nul 2>&1
SetLocal EnableDelayedExpansion
>> "%Log%" (
	echo(
	echo ****************************************************************************************************
	echo(                                 No Microsoft Scheduled Tasks List
	echo ****************************************************************************************************
	@for /f "delims=" %%I in ('powershell -noprofile "iex (${%~f0}|out-string)"') do echo %%I
	REM @For /F "tokens=2,9,17,19,20,21,22 delims=," %%a in ('SCHTASKS /Query /NH /FO CSV /V ^|find /I /V "Microsoft" ^|findstr /I /C:"VBS" /C:"EXE"') do (
	REM	Set TaskName=%%~a
	REM	Set TaskPath=%%~b
	REM	Call :Trim_Dequote !TaskName! TaskName
	REM	Call :Trim_Dequote !TaskPath! TaskPath
	REM	echo "!TaskName!"
	REM	echo "!TaskPath!"
	REM	echo %%c;%%d;%%f;%%g
	REM	echo( ---------------------------------------------------------------------------------------------------
	REM )
)

>> "%Log%" (
	echo(
	echo ****************************************************************************************************
	echo(                                 No Microsoft Services List
	echo ****************************************************************************************************
@for /f "tokens=*" %%a in (
	'WMIC service where "Not PathName like '%%Micro%%' AND Not PathName like '%%Windows%%'" get Name^,DisplayName^,PathName^,Status'
	) do (
		@for /f "delims=" %%b in ("%%a") do (
			echo %%b
			)
	)
)

>> "%Log%" (
	echo(
	echo ****************************************************************************************************
	echo %Winlogonkey%
	Reg Query "%Winlogonkey%" | find /I "userinit"
	@for /f "delims=" %%a in ('Reg Query "%ImageFileExec_Key%" /f "*.exe" ^|findstr /I /V ":"') do (
		@for /f "delims=" %%b in ('Reg Query "%%~a" /s /f "Debugger" ^|findstr /I /V "0" ^|findstr /I /V "1"') do (
			echo %%b
		)
	)
)

Call :ExtractCmdLine_Hashes
If Exist "%Log%" Start /MAX "Log" "%Log%" & Exit 
::-----------------------------------------------------------------------------------
:Trim_Dequote <Var> <NewVar>
(
	echo	Wscript.echo Trim_Dequote("%~1"^)
	echo	Function Trim_Dequote(S^)
	echo	If Left(S, 1^) = """" And Right(S, 1^) = """" Then Trim_Dequote = Trim(Mid(S, 2, Len(S^) - 2^)^) Else Trim_Dequote = Trim(S^)
	echo	End Function
)>"%VbsFile%"
for /f "delims=" %%a in ('Cscript //nologo "%VbsFile%"') do ( 
	set "%2=%%a" 
)
Del "%VbsFile%" /F >nul 2>&1
exit /b
REM ------------------------------------------------------------------------------
:GetFileNameWithDateTime <FileName>
for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set "MyDate=%%x"
set "%1=%MyDate:~0,4%-%MyDate:~4,2%-%MyDate:~6,2%-%MyDate:~8,2%-%MyDate:~10,2%"
Exit /B
REM -----------------------------------------------------------------------------
:Generate_VBS_File
>"%VbsFile%" ( 
	echo	Option Explicit
	echo	Dim Ws,objStartFolder,objFSO,objFolder,colFiles
	echo	Dim objFile,strFilePath,Lnk,Title
	echo	Title = "Extracting Target Path from .lnk and .url files by Hackoo 2020"
	echo	Set Ws = CreateObject("Wscript.Shell"^)
	echo	If WSH.Arguments.Count = 0 Then MsgBox "Missing Arguments",vbExclamation,Title : Wscript.Quit(1^)
	echo	objStartFolder = WSH.Arguments(0^)
	echo	Set objFSO = CreateObject("Scripting.FileSystemObject"^)
	echo	Set objFolder = objFSO.GetFolder(objStartFolder^)
	echo	Set colFiles = objFolder.Files
	echo	For Each objFile in colFiles
	echo	strFilePath = objFile.Path
	echo	  If Ucase(objFSO.GetExtensionName(strFilePath^)^) = "LNK"_
	echo	   Or Ucase(objFSO.GetExtensionName(strFilePath^)^) = "URL" Then
	echo	      Call ExtractTargetPath(strFilePath^)
	echo	  End If
	echo	Next
	echo	'-------------------------------------------------------------
	echo	Sub ExtractTargetPath(Lnk^)
	echo	set Lnk = Ws.Createshortcut(Lnk^)
	echo	WScript.echo "Link="^& DblQuote(Lnk^) ^& vbcrlf ^&_
	echo	"Target="^& DblQuote(Lnk.TargetPath^) ^& vbcrlf ^&_
	echo	String(100,"-"^)
	echo	End Sub
	echo	'-------------------------------------------------------------
	echo	Function DblQuote(Str^)
	echo	    DblQuote = Chr(34^) ^& Str ^& Chr(34^)
	echo	End Function
	echo	'-------------------------------------------------------------
)
Exit /B
REM -----------------------------------------------------------------------------
:Execute_VBS_File
cscript //nologo "%VbsFile%" "%~1"
Exit /B
REM -----------------------------------------------------------------------------
:RunAsAdmin
cls & color 9E & Mode 95,5
echo(
echo(               ===========================================================
echo(                    Please wait a while ... Running as Admin ....
echo(               ===========================================================
Powershell start -verb runas '%0' Admin & Exit
REM -----------------------------------------------------------------------------
:ExtractCmdLine_Hashes
Rem Killing all Process that have a status not responding
Taskkill /f /fi "status eq not responding">nul 2>&1
Set "LogScan=%~dp0Log_Scan"
If Not Exist %LogScan%\ MD %LogScan%
Set "Abs_cmdline=%LogScan%\%~n0_Abs_cmdline.txt"
Set "Tmp_cmdline=%LogScan%\%~n0_Tmp_cmdline.txt
Set "cmdline=%LogScan%\%~n0_cmdline.txt
Set "TmpHashes=%LogScan%\%~n0_TmpHashes.txt"
Set "Hashes=%LogScan%\%~n0_Hashes.txt"
Set "Hash2Check_VirusTotal=%LogScan%\Hash2Check_VirusTotal.txt"
For %%a in ("%Abs_cmdline%" "%Tmp_cmdline%" "%TmpHashes%" "%Hash2Check_VirusTotal%") Do If Exist "%%a" Del "%%a"
Set ProcessNames="wscript.exe" "cmd.exe" "powershell.exe" "cscript.exe" "svchost.exe"
SetLocal EnableDelayedExpansion
for %%A in (%ProcessNames%) Do (
	REM echo(
	REM echo Please Wait a while ... Looking for any instance of %%A ...
	Call :GetCommandLine %%A>nul 2>&1 
)
Timeout /T 1 /NoBreak>nul
Call :Extract "%Abs_cmdline%" "%Tmp_cmdline%"
for /f "delims=" %%a in ('Type "%Tmp_cmdline%"') do (
	for /f "skip=1 delims=" %%H in ('CertUtil -hashfile "%%~a" SHA256 ^| findstr /i /v "CertUtil"') do set "H=%%H"
		REM echo %%a=!H: =!
		echo %%a=!H: =! >> "%TmpHashes%"
)

Call :RemoveDuplicateEntry %TmpHashes% %Hashes%
Call :RemoveDuplicateEntry %Tmp_cmdline% %cmdline%
If exist "%TmpHashes%" Del "%TmpHashes%" & If exist "%Tmp_cmdline%" Del "%Tmp_cmdline%"

for /f "tokens=1,2 delims==" %%a in ('Type "%Hashes%"') do (
	If /I "%%~xa"==".vbs" MD "%LogScan%\VBS">nul 2>&1 & Type "%%a" > "%LogScan%\VBS\%%~nxa.txt"
	If /I "%%~xa"==".vbe" MD "%LogScan%\VBE">nul 2>&1 & Type "%%a" > "%LogScan%\VBE\%%~nxa.txt"
	If /I "%%~xa"==".js"  MD "%LogScan%\JS">nul  2>&1 & Type "%%a" > "%LogScan%\JS\%%~nxa.txt"
	If /I "%%~xa"==".jse" MD "%LogScan%\JSE">nul 2>&1 & Type "%%a" > "%LogScan%\JSE\%%~nxa.txt"
	If /I "%%~xa"==".bat" MD "%LogScan%\BAT">nul 2>&1 & Type "%%a" > "%LogScan%\BAT\%%~nxa.txt"
	If /I "%%~xa"==".cmd" MD "%LogScan%\CMD">nul 2>&1 & Type "%%a" > "%LogScan%\CMD\%%~nxa.txt"
	If /I "%%~xa"==".ps1" MD "%LogScan%\PS1">nul 2>&1 & Type "%%a" > "%LogScan%\PS1\%%~nxa.txt"
	If /I "%%~xa"==".wsf" MD "%LogScan%\WSF">nul 2>&1 & Type "%%a" > "%LogScan%\WSF\%%~nxa.txt"
	Set "Hash=%%b"
	Set "Hash=!Hash: =!
	IF {!Hash!} NEQ {!CMD_HASH!} (
		IF {!Hash!} NEQ {!PS_HASH!} (
			echo https://www.virustotal.com/#/file/%%b>>"%Hash2Check_VirusTotal%"
			Start "Chek SHA256 on VIRUSTOTAL" "https://www.virustotal.com/old-browsers/file/%%b"
	 	)
	)
)
::Start "" /MAX "%Hashes%" 
::Start "" /MAX "%cmdline%"
Exit /B
::********************************************************************************************************
:GetCommandLine <ProcessName>
Set "ProcessCmd="
for /f "tokens=2 delims==" %%P in ('wmic process where caption^="%~1" get commandline /format:list ^| findstr /I "%~1" ^| find /I /V "%~nx0" 2^>nul') do (
	Set "ProcessCmd=%%P"
	REM echo !ProcessCmd!
	echo !ProcessCmd! >> "%Abs_cmdline%"
)
Exit /b
::********************************************************************************************************
:Extract <InputData> <OutPutData>
(
echo Data = WScript.StdIn.ReadAll
echo Data = Extract(Data,"(^?^!.*(\x22\w\W^)^).*(\.ps1^|\.vbs^|\.vbe^|\.js^|\.jse^|\.cmd^|\.bat^|\.wsf^|\.exe^)(^?^!.*(\x22\w\W^)^)"^)
echo WScript.StdOut.WriteLine Data
echo Function Extract(Data,Pattern^)
echo    Dim oRE,oMatches,Match,Line
echo    set oRE = New RegExp
echo    oRE.IgnoreCase = True
echo    oRE.Global = True
echo    oRE.Pattern = Pattern
echo    set oMatches = oRE.Execute(Data^)
echo    If not isEmpty(oMatches^) then
echo        For Each Match in oMatches  
echo            Line = Line ^& Trim(Match.Value^) ^& vbcrlf
echo        Next
echo        Extract = Line 
echo    End if
echo End Function
)>"%tmp%\%~n0.vbs"
cscript /nologo "%tmp%\%~n0.vbs" < "%~1" > "%~2"
If Exist "%tmp%\%~n0.vbs" Del "%tmp%\%~n0.vbs"
exit /b
::****************************************************
::----------------------------------------------------
:RemoveDuplicateEntry <InputFile> <OutPutFile>
Powershell  ^
$Contents=Get-Content '%1';  ^
$LowerContents=$Contents.ToLower(^);  ^
$LowerContents ^| select -unique ^| Out-File '%2'
Exit /b
::----------------------------------------------------
: end Batch / begin PowerShell hybrid code #>
Function getTasks($path) {
    $out = @()
    # Get root tasks
    $schedule.GetFolder($path).GetTasks(0) | % {
        $xml = [xml]$_.xml
        $out += New-Object psobject -Property @{
            "Name" = $_.Name
            "Path" = $_.Path
            "LastRunTime" = $_.LastRunTime
            "NextRunTime" = $_.NextRunTime
            "Actions" = ($xml.Task.Actions.Exec | % { "$($_.Command) $($_.Arguments)" }) -join "`n"
"==============" = "===================================================================================="
        }
    }
    # Get tasks from subfolders
    $schedule.GetFolder($path).GetFolders(0) | % {
        $out += getTasks($_.Path)
    }
    #Output
    $out
}
$tasks = @()
$schedule = New-Object -ComObject "Schedule.Service"
$schedule.Connect() 
# Start inventory
$tasks += getTasks("\")
# Close com
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedule) | Out-Null
Remove-Variable schedule

# To show All No Microsoft Scheduled Tasks
$tasks | ? { $_.Path -notmatch "Micro*" } | Out-String -Width 450