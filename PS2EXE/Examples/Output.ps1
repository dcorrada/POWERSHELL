Write-Output "Write-Output"
$DebugPreference = "Continue"
Write-Debug "Write-Debug"
$VerbosePreference = "Continue"
Write-Verbose "Write-Verbose"
Write-Warning "Write-Warning"
Write-Error "Write-Error"

# keep following windows in foreground with -noConsole:
$Host.UI.RawUI.FlushInputBuffer()
ipconfig | Out-String
$Host.UI.RawUI.FlushInputBuffer()

Read-Host "Read-Host: Press key to exit"

Write-Host "Write-Host: Done"



