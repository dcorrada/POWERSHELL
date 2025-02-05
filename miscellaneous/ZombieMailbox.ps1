<#
Name......: ZombieMailbox.ps1
Version...: 25.2.3
Author....: Dario CORRADA

This script look for any [user|shared] mailbox present on ExchangeOnLine. Then 
it seek for any [SeandAs|SendOnBehalfTo|FullAccess] permission for each one. 
The aim is to find any account delegated revealed as dismissed (aka zombie).

More details about ExOv3 module cmdlets are available at:
https://learn.microsoft.com/en-us/powershell/module/exchange/?view=exchange-ps#powershell-v3-module
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script."
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

# importing third party modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module ExchangeOnlineManagement
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'ExchangeOnlineManagement')) {
            Install-Module ExchangeOnlineManagement -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ExchangeOnlineManagement] module: click Ok to restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } elseif (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
            Install-Module ImportExcel -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: click Ok restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') > $null
            Write-Output "`nError: $($error[0].ToString())"
            Pause
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                    QUERYING
******************************************************************************* #>
Write-Host -NoNewline "Connecting to ExchangeOnLine... "
try {
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host -ForegroundColor Green "Ok"
}
catch {
    Write-Host -ForegroundColor Red "Ko"
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

Write-Host -NoNewline "Fetching mailbox list..."
<#
The Id (Identity) attribute obtained from Get-EXOMailbox cmdlet may store 
different kind of id types. The most adopted ones are:
* Object ID           ie: 5c9bc377-137d-4b04-83d4-248b1ee6c705
* Security ID         ie: S-1-5-21-2262606477-833299019-2854854355-8406544
* User Principal Name ie: username@domain
* Legacy User Name    ie: username
#>

$EXOlist = Get-EXOMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize unlimited | ForEach-Object {
    Write-Host -NoNewline '.'        
    New-Object -TypeName PSObject -Property @{
        ID          = "$($_.Id)"
        TYPE        = "$($_.RecipientType)"
        OBJ_ID      = "$($_.ExternalDirectoryObjectId)"
        UPN         = "$($_.UserPrincipalName)"
        DISPLAY     = "$($_.DisplayName)"
    } | Select ID, TYPE, OBJ_ID, UPN, DISPLAY
}
Write-Host -ForegroundColor Green " Done"

Write-Host "Searching orphaned..."
$orphaned = Get-EXOMailbox -InactiveMailboxOnly -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize unlimited | ForEach-Object {
    Write-Host -NoNewline '.'        
    New-Object -TypeName PSObject -Property @{
        OBJ_ID      = "$($_.ExternalDirectoryObjectId)"
        UPN         = "$($_.UserPrincipalName)"
    } | Select OBJ_ID, UPN
}
if ($orphaned.Count -gt 0) {
    Write-Host "`n"
    foreach ($item in $orphaned) {
        Write-Host -ForegroundColor Blue "[$($item.OBJ_ID)] $($item.UPN)"
    }
    Write-Host -ForegroundColor Red "`nThe mailboxes listed above result inactive"
    Pause
} else {
    Write-Host -ForegroundColor Green "No inactive mailbox found"
}

<# *******************************************************************************
                                FILTERING
******************************************************************************* #>
$answ = [System.Windows.MessageBox]::Show("Load exclude list text file?",'INFILE','YesNo','Info')
$ExcludeList = @{ 
    NONE = 'True'
}
if ($answ -eq 'yes') {
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Open File"
    $OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
    $OpenFileDialog.filter = 'Text file (*.txt)| *.txt'
    $OpenFileDialog.ShowDialog() | Out-Null
    $ExcludeFile = $OpenFileDialog.filename

    foreach ($aUPN in (Get-Content -Path $ExcludeFile)) {
        $ExcludeList["$aUPN"] = 'True'
    }
}

Write-Host "Gathering mailbox details..."
$EXOdetailed = @{}
$totKeys = $EXOlist.Count
$counterKeys = 0
$parsebar = ProgressBar
foreach ($entity in $EXOlist) {
    $counterKeys ++
    $EXOdetailed["$($entity.OBJ_ID)"] = @{
        UPN             = "$($entity.UPN)"
        DISPLAYNAME     = "$($entity.DISPLAY)"
        TYPE            = "$($entity.TYPE)"
        FULLACCESS      = @()
        SENDAS          = @()
        SENDONBEHALF    = @()
        NOMINAL         = 'False'
        LASTLOGON       = 'na'
    }

    $mbs = Get-MailboxStatistics -Identity $entity.UPN | Select LastLogonTime
    $EXOdetailed["$($entity.OBJ_ID)"].LASTLOGON = "$(($mbs.LastLogonTime | Get-Date -format 'yyyy-MM-dd HH:mm:ss').ToString())"

    if ($ExcludeList.ContainsKey($entity.UPN)) {
        Write-Host -ForegroundColor Black "$($entity.UPN)"
        $EXOdetailed["$($entity.OBJ_ID)"].NOMINAL = 'True'
        $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += 'SELF'
    } else {
        Write-Host -ForegroundColor Cyan "$($entity.UPN)"

        foreach ($sandman in (Get-EXOrecipientPermission -id $entity.OBJ_ID)) {
            if ((($sandman.AccessRights -join ',') -cmatch 'SendAs') -and ($sandman.Trustee -cne 'NT AUTHORITY\SELF')) {
                if ($sandman.Trustee -notmatch '@') {
                    if ($sandman.Trustee -notmatch "^S\-1\-") {
                        $ErrorActionPreference= 'Stop'
                        try {
                            $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$((Get-EXOMailbox -Identity $sandman.Trustee).UserPrincipalName)"
                        }
                        catch {
                            [System.Windows.MessageBox]::Show("There is something nasty with grant [SendAs] `nassigned to [$($sandman.Trustee)]","MAILBOX $($entity.UPN)",'Ok','Warning') | Out-Null
                        }
                        $ErrorActionPreference= 'Inquire'
                    } else {
                        Write-Host -ForegroundColor Yellow "  Value [SendAs]'$($sandman.Trustee)' needs to be fixed"
                        # temporary fill of a value that needs to be fixed
                        $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$($sandman.Trustee)"
                    }
                } else {
                    $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$($sandman.Trustee)"
                }
            }
        }

        foreach ($fuller in (Get-EXOMailboxPermission -id $entity.OBJ_ID)) {
            if (($fuller.AccessRights -join ',') -cmatch 'FullAccess') {
                if ($fuller.User -ceq 'NT AUTHORITY\SELF') {
                    $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += 'SELF'
                } else {
                    if ($fuller.User -notmatch '@') {
                        if ($fuller.User -notmatch "^S\-1\-") {
                            $ErrorActionPreference= 'Stop'
                            try {
                                $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += "$((Get-EXOMailbox -Identity $fuller.User).UserPrincipalName)"
                            }
                            catch {
                                [System.Windows.MessageBox]::Show("There is something nasty with grant [FullAccess] `nassigned to [$($fuller.User)]","MAILBOX $($entity.UPN)",'Ok','Warning') | Out-Null
                            }
                            $ErrorActionPreference= 'Inquire'
                        } else {
                            Write-Host -ForegroundColor Yellow "  Value [FullAccess]'$($fuller.User)' needs to be fixed"
                            # temporary fill of a value that needs to be fixed
                            $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$($fuller.User)"
                        }
                    } else {
                        $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += "$($fuller.User)"
                    }
                }
            }
        }

        foreach ($Beowulf in (Get-Mailbox -Identity $entity.OBJ_ID).GrantSendOnBehalfTo) {
            if ($Beowulf -notmatch '@') {
                if ($Beowulf -notmatch "^S\-1\-") {
                    $ErrorActionPreference= 'Stop'
                    try {
                        $EXOdetailed["$($entity.OBJ_ID)"].SENDONBEHALF += "$((Get-EXOMailbox -Identity $Beowulf).UserPrincipalName)"
                    }
                    catch {
                        [System.Windows.MessageBox]::Show("There is something nasty with grant [SendOnBehalf] `nassigned to [$Beowulf]","MAILBOX $($entity.UPN)",'Ok','Warning') | Out-Null
                    }
                    $ErrorActionPreference= 'Inquire'
                } else {
                    Write-Host -ForegroundColor Yellow "  Value [SendOnBehalf]'$($Beowulf)' needs to be fixed"
                    # temporary fill of a value that needs to be fixed
                    $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$Beowulf"
                }
            } else {
                $EXOdetailed["$($entity.OBJ_ID)"].SENDONBEHALF += $Beowulf
            }
        }
    }

    # progressbar
    $percent = ($counterKeys / $totKeys)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Mailbox {0} out of {1} parsed [{2}%]" -f ($counterKeys, $totKeys, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
Write-Host -ForegroundColor Green "Done"
$parsebar[0].Close()

Disconnect-ExchangeOnline -Confirm:$false

<# *******************************************************************************
                                GET OUTPUT
******************************************************************************* #>
$xlsx_file = "C:$env:HOMEPATH\Downloads\ZombieMailbox-" + (Get-Date -format "yyMMddHHmm") + '.xlsx'
$XlsPkg = Open-ExcelPackage -Path $xlsx_file -Create

# fare 3 worksheet? (nominali, non nominali, shared)

$ErrorActionPreference= 'Stop'
try {  
    $label = 'MailBox'
    Write-Host -NoNewline "Writing worksheet [$label]..."
    $inData = $EXOdetailed.Keys | Foreach-Object {
        Write-Host -NoNewline '.'
        $granted = $EXOdetailed[$_]
        foreach ($aGrant in $granted.FULLACCESS) {
            New-Object -TypeName PSObject -Property @{
                OBJECTID        = "$_"
                UPN             = "$($granted.UPN)"
                DISPLAYNAME     = "$($granted.DISPLAYNAME)"
                TYPE            = "$($granted.TYPE)"
                NOMINAL         = "$($granted.NOMINAL)"
                LASTLOGON       = [DateTime]$granted.LASTLOGON
                FULLACCESS      = "$aGrant"
                SENDAS          = ""
                SENDONBEHALF    = ""
            } | Select OBJECTID, UPN, DISPLAYNAME, TYPE, NOMINAL, LASTLOGON, FULLACCESS, SENDAS, SENDONBEHALF
        }

        foreach ($aGrant in $granted.SENDAS) {
            New-Object -TypeName PSObject -Property @{
                OBJECTID        = "$_"
                UPN             = "$($granted.UPN)"
                DISPLAYNAME     = "$($granted.DISPLAYNAME)"
                TYPE            = "$($granted.TYPE)"
                NOMINAL         = "$($granted.NOMINAL)"
                LASTLOGON       = [DateTime]$granted.LASTLOGON
                FULLACCESS      = ""
                SENDAS          = "$aGrant"
                SENDONBEHALF    = ""
            } | Select OBJECTID, UPN, DISPLAYNAME, TYPE, NOMINAL, LASTLOGON, FULLACCESS, SENDAS, SENDONBEHALF
        }

        foreach ($aGrant in $granted.SENDONBEHALF) {
            New-Object -TypeName PSObject -Property @{
                OBJECTID        = "$_"
                UPN             = "$($granted.UPN)"
                DISPLAYNAME     = "$($granted.DISPLAYNAME)"
                TYPE            = "$($granted.TYPE)"
                NOMINAL         = "$($granted.NOMINAL)"
                LASTLOGON       = [DateTime]$granted.LASTLOGON
                FULLACCESS      = ""
                SENDAS          = ""
                SENDONBEHALF    = "$aGrant"
            } | Select OBJECTID, UPN, DISPLAYNAME, TYPE, NOMINAL, LASTLOGON, FULLACCESS, SENDAS, SENDONBEHALF
        }
    }
    $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium2' -AutoSize -PassThru
    Write-Host -ForegroundColor Green ' DONE'
} catch {
    [System.Windows.MessageBox]::Show("Error updating data",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red ' FAIL'
    Write-Host -ForegroundColor Yellow "ERROR: $($error[0].ToString())"
    exit
}
$ErrorActionPreference= 'Inquire'

Close-ExcelPackage -ExcelPackage $XlsPkg

[System.Windows.MessageBox]::Show("File [$xlsx_file] has been created",'OUTPUT','Ok','Info') | Out-Null
