Param([string]$InFile='NULL')

<#
Name......: MOTA.ps1
Version...: 25.03.3
Author....: Dario CORRADA

This script would be an add on for all of the AssignedLicenses*.ps1 flavours. 
It reads, as input, the Excel reference file outputted from the last ones, then 
produce summary aomunt report of individual licenses assigned vs. that ones 
recovered (aka available to be assigned again).


    Everyday, well it's the same
    That bong that's on the table starts to call my name
                                                                [The Offspring]
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
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        Install-Module ImportExcel -Confirm:$False -Force
        [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: click Ok restart the script",'RESTART','Ok','warning') > $null
        $ThirdParty = 'Ko'
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'


<# *******************************************************************************
                               FETCHING RAW DATA
******************************************************************************* #>
if ($InFile -ceq 'NULL') {
    # looking for Excel reference file
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Open File"
    $OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
    $OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
    $OpenFileDialog.ShowDialog() | Out-Null
    $xlsx_file = $OpenFileDialog.filename
} else {
    $xlsx_file = $InFile
}

$Worksheet_list = Get-ExcelSheetInfo -Path $xlsx_file
$RawData = @{}

if ($Worksheet_list.Name -contains 'Assigned_Licenses') {
    if ($InFile -ceq 'NULL') {
        Write-Host -NoNewline "Importing [Assigned_Licenses] worksheet..."
    }
    foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Assigned_Licenses')) {
        if ($InFile -ceq 'NULL') {
            Write-Host -NoNewline '.'
        }
        $aUSER      = $history.USRNAME
        $aDATE      = $history.TIMESTAMP | Get-Date -format "yyyy/MM/dd"
        $aLICENSE   = $history.LICENSE
        $aNOTE      = 'null'

        # looking for those users whom any license has been currently assigned
        if ($aLICENSE -cne 'NONE') {
            if (!($RawData.ContainsKey($aUSER))) {
                $RawData["$aUSER"] = @{}
            }
            if (!($RawData["$aUSER"].ContainsKey($aDATE))) {
                ($RawData["$aUSER"])["$aDATE"] = @()
            }
            ($RawData["$aUSER"])["$aDATE"] += "$aLICENSE>>$aNOTE"
        }
    }
    if ($InFile -ceq 'NULL') {
        Write-Host -ForegroundColor Green ' DONE'
    }
} else {
    [System.Windows.MessageBox]::Show("The Excel file does not contain necessary data`n`n[$xlsx_file]",'ABORTING','Ok','Error') > $null
    Exit
}

if ($Worksheet_list.Name -contains 'Orphaned') {
    if ($InFile -ceq 'NULL') {
        Write-Host -NoNewline "Importing [Orphaned] worksheet..."
    }
    foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Orphaned')) {
        if ($InFile -ceq 'NULL') {
            Write-Host -NoNewline '.'
        }
        $aUser      = $history.USRNAME
        $aDATE      = $history.TIMESTAMP | Get-Date -format "yyyy/MM/dd"
        $aLICENSE   = $history.LICENSE
        $aNOTE      = "$($history.NOTES) since>>$($history.CREATED | Get-Date -format "yyyy/MM/dd")"

        # gathering dismissed license data
        if (!($RawData.ContainsKey($aUSER))) {
            $RawData["$aUSER"] = @{}
        }
        if (!($RawData["$aUSER"].ContainsKey($aDATE))) {
            ($RawData["$aUSER"])["$aDATE"] = @()
        }
        ($RawData["$aUSER"])["$aDATE"] += "$aLICENSE>>$aNOTE"
    }
    if ($InFile -ceq 'NULL') {
        Write-Host -ForegroundColor Green ' DONE'
    }
} else {
    [System.Windows.MessageBox]::Show("Worksheet 'Orphaned not found in the Excel file`n`n[$xlsx_file]",'WARNING','Ok','Warning') > $null
}

<# *******************************************************************************
                                PARSING
******************************************************************************* #>
# looking for those accounts that need to be further investigated
$ToBeParsed = @()
foreach ($currentUser in $RawData.Keys) {
    foreach ($currentDate in $RawData["$currentUser"].Keys) {
        foreach ($currentEvent in ($RawData["$currentUser"])["$currentDate"]) {
            if ($currentEvent -match ">>null") {
                # data which came from Assigned_Licenses worksheet
            } else {
                $ToBeParsed += "$currentUser"
            }
        }
    }
}

$ParsedData = @()
foreach ($currentUser in $RawData.Keys) {
    if ($ToBeParsed -contains $currentUser) {
        foreach ($currentDate in $RawData["$currentUser"].Keys) {
            foreach ($currentEvent in ($RawData["$currentUser"])["$currentDate"]) {
                if ($currentEvent -match ">>null") {
                    # data which came from AssignedLicenses worksheet but
                    # this account was previously involved (also found in Orphaned worksheet)
                    $currentEvent -match "^(.+)>>" | Out-Null
                    $currentLicense = $matches[1]
                    $ParsedData += [pscustomobject]@{
                        TIMESTAMP   = "$currentDate"
                        ACCOUNT     = $currentUser
                        LICENSE     = $currentLicense
                        STATUS      = 'ASSIGNED'
                    }
                } elseif ($currentEvent -match "user no longer exists on tenant") {
                    $currentEvent -match "^(.+)>>user no longer exists on tenant since>>(.+)$" | Out-Null
                    $currentLicense = $matches[1]
                    $birthday = $matches[2]
                    if ($currentLicense -cne 'NONE') {
                        # dismissed account
                        $ParsedData += [pscustomobject]@{
                            TIMESTAMP   = "$currentDate"
                            ACCOUNT     = $currentUser
                            LICENSE     = $currentLicense
                            STATUS      = 'RECOVERED'
                        }
                        # previous hypotetic assignement (based on account creation date)
                        $ParsedData += [pscustomobject]@{
                            TIMESTAMP   = "$birthday"
                            ACCOUNT     = $currentUser
                            LICENSE     = $currentLicense
                            STATUS      = 'ASSIGNED'
                        }
                    } else {
                        # dismissed account in which license have been recovered before (nothing to do)
                    }
                } elseif ($currentEvent -match 'license dismissed for this user') {
                    # dismissed license
                    $currentEvent -match "^(.+)>>license dismissed for this user" | Out-Null
                    $currentLicense = $matches[1]
                    $ParsedData += [pscustomobject]@{
                        TIMESTAMP   = "$currentDate"
                        ACCOUNT     = $currentUser
                        LICENSE     = $currentLicense
                        STATUS      = 'RECOVERED'
                    }
                } elseif ($currentEvent -match "assigned license\(s\) to this user") {
                    # account created without any license, assignement was done afterwards
                    # and already collected from AssignedLicenses worksheet

                    $currentEvent -match "^(.+)>>assigned license\(s\) to this user" | Out-Null
                    $currentLicense = $matches[1]
                    if ($currentLicense -cne 'NONE') {
                        # just handle any exception due to old releases of AssignedLicenses*.ps1 script
                        Write-Host -ForegroundColor Yellow @"

Uhmmmm... Something screwy is going on!
Maybe you are using an old version of AssignedLicenses*.ps1 script?

Here there are the info extracted from Orphaned worksheet:

    TIMESTAMP   = $currentDate
    ACCOUNT     = $currentUser
    LICENSE     = $currentLicense

"@
                        Pause
                    }
                }
            }            
        } 
    } else {
        foreach ($currentDate in $RawData["$currentUser"].Keys) {
            foreach ($assignement in ($RawData["$currentUser"])["$currentDate"]) {
                $assignement -match "^(.+)>>" | Out-Null
                $currentLicense = $matches[1]
                $ParsedData += [pscustomobject]@{
                    TIMESTAMP   = "$currentDate"
                    ACCOUNT     = $currentUser
                    LICENSE     = $currentLicense
                    STATUS      = 'ASSIGNED'
                }
            }
        }
    }
}


<# *******************************************************************************
                                OUTPUT
******************************************************************************* #>
$OutFile = "C:$env:HOMEPATH\Downloads\mota.csv"
if (Test-Path $OutFile -PathType Leaf) {
    Remove-Item -Path $OutFile -Force | Out-Null
}

$ParsedData | Export-Csv -Path $OutFile -NoTypeInformation

if ($InFile -cne 'NULL') {
    Write-Host $OutFile
}
