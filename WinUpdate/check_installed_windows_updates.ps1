############################################################################################
#
# NAME: 	check_installed_windows_updates.ps1
#
# COMMENT:  Script to check for installed windows updates splitting between important and recommended
#
# BUILDING DATE: 03/03/2016
#
# AUTHOR: Mentor
#
# SITE: www.internauta37.altervista.org
#
# EMAIL: internauta37@altervista.org
#
############################################################################################
#
# BASED ON:
# https://exchange.nagios.org/directory/Plugins/Operating-Systems/Windows-NRPE/Check-Windows-Updates-using-Powershell/details 
# http://tomtalks.uk/2013/09/list-all-microsoftwindows-updates-with-powershell-sorted-by-kbhotfixid-get-microsoftupdate/
#
############################################################################################

#Begin Transcript
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference="Continue"
Start-Transcript -path 'Installed_Updates.txt' -append

$htReplace = New-Object hashtable
foreach ($letter in (Write-Output ä ae ö oe ü ue Ä Ae Ö Oe Ü Ue ß ss)) {
    $foreach.MoveNext() | Out-Null
    $htReplace.$letter = $foreach.Current
}
$pattern = "[$(-join $htReplace.Keys)]"

$returnStateOK = 0
$returnStateWarning = 1
$returnStateCritical = 2
$returnStateUnknown = 3
$returnStatePendingReboot = $returnStateWarning
$returnStateOptionalUpdates = $returnStateWarning

$updateCacheFile = "check_installed_windows_updates-cache.xml"
$updateCacheExpireHours = "24"

$logFile = "check_installed_windows_update.log"

function LogLine(	[String]$logFile = $(Throw 'LogLine:$logFile unspecified'),
					[String]$row = $(Throw 'LogLine:$row unspecified')) {
	$logDateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
	Add-Content -Encoding UTF8 $logFile ($logDateTime + " - " + $row)
}

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){
	Write-Host "updates installed, reboot required"
	if (Test-Path $logFile) {
		Remove-Item $logFile | Out-Null
	}
	if (Test-Path $updateCacheFile) {
		Remove-Item $updateCacheFile | Out-Null
	}
	exit $returnStatePendingReboot
}

if (-not (Test-Path $updateCacheFile)) {
	LogLine -logFile $logFile -row ("$updateCacheFile not found, creating....")
	$updateSession = new-object -com "Microsoft.Update.Session"
	$updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=1 and Type='Software'")).Updates
	Export-Clixml -InputObject $updates -Encoding UTF8 -Path $updateCacheFile
}

if ((Get-Date) -gt ((Get-Item $updateCacheFile).LastWriteTime.AddHours($updateCacheExpireHours))) {
	LogLine -logFile $logFile -row ("update cache expired, updating....")
	$updateSession = new-object -com "Microsoft.Update.Session"
	$updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=1 and Type='Software'")).Updates
	Export-Clixml -InputObject $updates -Encoding UTF8 -Path $updateCacheFile
} else {
	LogLine -logFile $logFile -row ("using valid cache file....")
	$updates = Import-Clixml $updateCacheFile
}

if ($updates.Count -eq 0) {
	Write-Host "OK - no pending updates."
	exit $returnStateOK
}

$Total=  @()
$Critical=  @()
$Optional=  @()

foreach ($update in $updates) {

	$string = $update.title
	
	$Regex = “KB\d*”
	$KB = $string | Select-String -Pattern $regex | Select-Object { $_.Matches }
	
	$updateformat = New-Object -TypeName PSobject
	$updateformat | add-member NoteProperty “HotFixID” -value $KB.‘ $_.Matches ‘.Value
	$updateformat | add-member NoteProperty “Title” -value $string
	
	$Total += $updateformat
	
	if ($update.AutoSelectOnWebSites) {
		$Critical += $updateformat
	} else {
		$Optional += $updateformat
	}

}

#Pending Updates Tally
write-host "`n"
Write-Host "=====Updates Tally====="
write-host "`n"
Write-Host "Total Installed Updates: $($Total.Count)"  -foregroundcolor "green"
write-host "`n"
Write-Host "Important Updates: $($Critical.Count)" -foregroundcolor "red"
Write-Host "Optional Updates: $($Optional.Count)" -foregroundcolor "yellow"
write-host "`n"

if (($($Critical.Count) + $($Optional.Count)) -gt 0) {
	Write-Host “$($Critical.Count) Important Updates Found” -foregroundcolor "red"
	$Critical | Sort-Object HotFixID | Format-Table -AutoSize
	write-host "`n"
	Write-Host “$($Optional.Count) Optional Updates Found” -foregroundcolor "yellow"
	$Optional | Sort-Object HotFixID | Format-Table -AutoSize
}

#if ($($Critical.Count) -gt 0 -or $($Optional.Count) -gt 0) {
#	Start-Process "wuauclt.exe" -ArgumentList "/detectnow" -WindowStyle Hidden
#}

if ($($Critical.Count) -gt 0) {
	exit $returnStateCritical
}

if ($($Optional.Count) -gt 0) {
	exit $returnStateOptionalUpdates
}

Write-Host "UNKNOWN script state"
exit $returnStateUnknown