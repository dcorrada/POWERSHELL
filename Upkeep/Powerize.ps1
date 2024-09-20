<#
Name......: Powerize.ps1
Version...: 21.06.1
Author....: Dario CORRADA

This script set power management profile
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

function Get-PowerScheme {
<#
Get the currently active PowerScheme
This will query the current power scheme and return the GUID and user friendly name
#>

    [CmdletBinding()][OutputType([object])]
    param ()

    #Get the currently active power scheme
    $Query = powercfg.exe /getactivescheme
    #Get the alias name of the active power scheme
    $ActiveSchemeName = ($Query.Split("()").Trim())[1]
    #Get the GUID of the active power scheme
    $ActiveSchemeGUID = ($Query.Split(":(").Trim())[1]
    $Query = powercfg.exe /query $ActiveSchemeGUID
    try {
        $GUIDAlias = ($Query | where { $_.Contains("GUID Alias:") }).Split(":")[1].Trim()
    }
    catch {
        $GUIDAlias = ''
    }
    $Scheme = New-Object -TypeName PSObject
    $Scheme | Add-Member -Type NoteProperty -Name PowerScheme -Value $ActiveSchemeName
    $Scheme | Add-Member -Type NoteProperty -Name GUIDAlias -Value $GUIDAlias
    $Scheme | Add-Member -Type NoteProperty -Name GUID -Value $ActiveSchemeGUID
    Return $Scheme
}

function Set-PowerSchemeSettings {
<#
Modify current power scheme
This will modify settings of the currently active power scheme.
#>
    
    [CmdletBinding()]
    param
    (
        [string]
        $MonitorTimeoutAC,
        [string]
        $MonitorTimeoutDC,
        [string]
        $DiskTimeoutAC,
        [string]
        $DiskTimeoutDC,
        [string]
        $StandbyTimeoutAC,
        [string]
        $StandbyTimeoutDC,
        [string]
        $HibernateTimeoutAC,
        [string]
        $HibernateTimeoutDC
    )
    
    $Scheme = Get-PowerScheme
    If (($MonitorTimeoutAC -ne $null) -and ($MonitorTimeoutAC -ne "")) {
        Write-Host "Setting monitor timeout on AC to"$MonitorTimeoutAC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "monitor-timeout-ac" + [char]32 + $MonitorTimeoutAC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\7516b95f-f776-4464-8c53-06167f40cc99\3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"
        $TestValue = $MonitorTimeoutAC
        $PowerIndex = "ACSettingIndex"
    }
    If (($MonitorTimeoutDC -ne $null) -and ($MonitorTimeoutDC -ne "")) {
        Write-Host "Setting monitor timeout on DC to"$MonitorTimeoutDC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "monitor-timeout-dc" + [char]32 + $MonitorTimeoutDC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\7516b95f-f776-4464-8c53-06167f40cc99\3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"
        $TestValue = $MonitorTimeoutDC
        $PowerIndex = "DCSettingIndex"
    }
    If (($DiskTimeoutAC -ne $null) -and ($DiskTimeoutAC -ne "")) {
        Write-Host "Setting disk timeout on AC to"$DiskTimeoutAC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "disk-timeout-ac" + [char]32 + $DiskTimeoutAC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\0012ee47-9041-4b5d-9b77-535fba8b1442\6738e2c4-e8a5-4a42-b16a-e040e769756e"
        $TestValue = $DiskTimeoutAC
        $PowerIndex = "ACSettingIndex"
    }
    If (($DiskTimeoutDC -ne $null) -and ($DiskTimeoutDC -ne "")) {
        Write-Host "Setting disk timeout on DC to"$DiskTimeoutDC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "disk-timeout-dc" + [char]32 + $DiskTimeoutDC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\0012ee47-9041-4b5d-9b77-535fba8b1442\6738e2c4-e8a5-4a42-b16a-e040e769756e"
        $TestValue = $DiskTimeoutDC
        $PowerIndex = "DCSettingIndex"
    }
    If (($StandbyTimeoutAC -ne $null) -and ($StandbyTimeoutAC -ne "")) {
        Write-Host "Setting standby timeout on AC to"$StandbyTimeoutAC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "standby-timeout-ac" + [char]32 + $StandbyTimeoutAC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\238c9fa8-0aad-41ed-83f4-97be242c8f20\29f6c1db-86da-48c5-9fdb-f2b67b1f44da"
        $TestValue = $StandbyTimeoutAC
        $PowerIndex = "ACSettingIndex"
    }
    If (($StandbyTimeoutDC -ne $null) -and ($StandbyTimeoutDC -ne "")) {
        Write-Host "Setting standby timeout on DC to"$StandbyTimeoutDC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "standby-timeout-dc" + [char]32 + $StandbyTimeoutDC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\238c9fa8-0aad-41ed-83f4-97be242c8f20\29f6c1db-86da-48c5-9fdb-f2b67b1f44da"
        $TestValue = $StandbyTimeoutDC
        $PowerIndex = "DCSettingIndex"
    }
    If (($HibernateTimeoutAC -ne $null) -and ($HibernateTimeoutAC -ne "")) {
        Write-Host "Setting hibernate timeout on AC to"$HibernateTimeoutAC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "hibernate-timeout-ac" + [char]32 + $HibernateTimeoutAC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\238c9fa8-0aad-41ed-83f4-97be242c8f20\9d7815a6-7ee4-497e-8888-515a05f02364"
        [int]$TestValue = $HibernateTimeoutAC
        $PowerIndex = "ACSettingIndex"
    }
    If (($HibernateTimeoutDC -ne $null) -and ($HibernateTimeoutDC -ne "")) {
        Write-Host "Setting hibernate timeout on DC to"$HibernateTimeoutDC" minutes....." -NoNewline
        $Switches = "/change" + [char]32 + "hibernate-timeout-dc" + [char]32 + $HibernateTimeoutDC
        $TestKey = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" + $Scheme.GUID + "\238c9fa8-0aad-41ed-83f4-97be242c8f20\9d7815a6-7ee4-497e-8888-515a05f02364"
        $TestValue = $HibernateTimeoutDC
        $PowerIndex = "DCSettingIndex"
    }
    $ErrCode = (Start-Process -FilePath "powercfg.exe" -ArgumentList $Switches -WindowStyle Minimized -Wait -Passthru).ExitCode
    $RegValue = (((Get-ItemProperty $TestKey).$PowerIndex) /60)
    #Round down to the nearest tenth due to hibernate values being 1 decimal off
    $RegValue = $RegValue - ($RegValue % 10)
    If (($RegValue -eq $TestValue) -and ($ErrCode -eq 0)) {
        Write-Host "Success" -ForegroundColor Yellow
        $Errors = $false
    } else {
        Write-Host "Failed" -ForegroundColor Red
        $Errors = $true
    }
    Return $Errors
}

Set-PowerSchemeSettings -MonitorTimeoutAC 15
Set-PowerSchemeSettings -MonitorTimeoutDC 15
Set-PowerSchemeSettings -DiskTimeoutAC 0
Set-PowerSchemeSettings -DiskTimeoutDC 0
Set-PowerSchemeSettings -StandbyTimeoutAC 0
Set-PowerSchemeSettings -StandbyTimeoutDC 30
Set-PowerSchemeSettings -HibernateTimeoutAC 0
Set-PowerSchemeSettings -HibernateTimeoutDC 0

# print out a battery report
$outfile = "$env:USERPROFILE\Downloads\BATTERY_REPORT.html"
powercfg.exe /batteryreport /output $outfile