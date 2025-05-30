<#
Name......: ZombieMailboxSDK.ps1
Version...: 25.5.3
Author....: Dario CORRADA

This script look for any [user|shared] mailbox present on ExchangeOnLine. Then 
it seek for any [SeandAs|SendOnBehalfTo|FullAccess] permission for each one. 
The aim is to find any account delegated revealed as dismissed (aka zombie).

More details about ExOv3 module cmdlets are available at:
https://learn.microsoft.com/en-us/powershell/module/exchange/?view=exchange-ps#powershell-v3-module

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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent | Split-Path -Parent

# graphical stuff
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
        } elseif (!(((Get-InstalledModule).Name) -contains 'Microsoft.Graph')) {
            Install-Module Microsoft.Graph -Scope AllUsers -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [Microsoft.Graph] module: click Ok restart the script",'RESTART','Ok','warning') > $null
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

# available only for local domain joined hosts and AD module installed
Write-Host -NoNewline "SID filter..."
$SIDfilter = $false
if ((Get-CimInstance win32_computersystem).PartOfDomain) { 
    if ((Get-Module -Name ActiveDirectory -ListAvailable) -ne $null) {
        Write-Host -ForegroundColor Green " enabled"
        $SIDfilter = $true
    } else {
        $ErrorActionPreference= 'Stop'
        try {
            Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online
            Write-Host -ForegroundColor Green " enabled"
            $SIDfilter = $true
        } catch {
            [System.Windows.MessageBox]::Show("Unable to install RSAT",'WARNING','Ok','Warning') | Out-Null
            Write-Host -ForegroundColor Red " disabled"
        }
        $ErrorActionPreference= 'Inquire'
    }
} else { 	
    Write-Host -ForegroundColor Red " disabled"
}

if ($SIDfilter) {
    # retrieve a list of disabled users
    $DisabledUsers = @{}
    foreach ($DisabledItem in (Search-ADAccount -AccountDisabled -UsersOnly -ResultSetSize $null | Select-Object SID, UserPrincipalName)) {
        $DisabledUsers["$($DisabledItem.SID)"] = "$($DisabledItem.UserPrincipalName)"
    }
}

<# *******************************************************************************
                            CREDENTIALS MANAGEMENT
******************************************************************************* #>
Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Graph\AppKeyring.ps1"
if ($pswout.Count -eq 4) {
    $UPN = $pswout[0]
    $clientID = $pswout[1]
    $tenantID = $pswout[2]
    Write-Host -ForegroundColor Green ' ok'
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# connect to Tenant
Write-Host -NoNewline "Connecting to the Tenant..."
$ErrorActionPreference= 'Stop'
Try {
    $splash = Connect-MgGraph -ClientId $clientID -TenantId $tenantID 
    Write-Host -ForegroundColor Green ' ok'
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("Error connecting to the Tenant",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

<# *******************************************************************************
                                LICENSES
******************************************************************************* #>
Write-Host -NoNewline "Gathering assigned licenses..."
$avail_lics = @()
$AccountName = (Get-MgSubscribedSku)[1].AccountName
foreach ($item in (Get-MgSubscribedSku | Select-Object -Property SkuPartNumber, ConsumedUnits -ExpandProperty PrepaidUnits)) {
    if ($item.Enabled -lt 10000) { # excluding the broadest licenses
        $avail_lics += $item.SkuPartNumber
    }
}

# SkuID to SkuPartNumber hash table
$SkuID2name = @{}
Get-MgSubscribedSku | Select-Object -Property SkuId, SkuPartNumber | foreach { 
    $SkuID2name[$_.SkuId] = $_.SkuPartNumber 
}

# retrieve all users list
$MsolUsrData = @{}
$tot = (Get-MgUser -All).Count
$usrcount = 0
$parsebar = ProgressBar
$MgUsrs = Get-MgUser -All -Property UserPrincipalName, DisplayName, UserType, accountEnabled, CreatedDateTime, assignedLicenses  `
    | Select-Object UserPrincipalName, DisplayName, UserType, accountEnabled, CreatedDateTime, assignedLicenses
foreach ($item in ($MgUsrs | Sort-Object DisplayName)) {
    $usrcount ++
    if (!($MsolUsrData.ContainsKey($item.UserPrincipalName))) {
        $MsolUsrData[$item.UserPrincipalName] = 'null'    
    }
    
    if ($item.AssignedLicenses.Count -ge 1) {
        $licenses = @()
        foreach ($accountsku in $item.AssignedLicenses) {
            $SkuName = $SkuID2name[$accountsku.SkuID]
            if ($avail_lics -contains $SkuName) { # filtering only managed licenses
                $licenses += $SkuName
            }
        }
        $MsolUsrData[$item.UserPrincipalName] = $licenses -join ' '
    } else {
        $MsolUsrData[$item.UserPrincipalName] = 'NONE'
    }

    # progressbar
    $percent = ($usrcount / $tot)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("User {0} out of {1} parsed [{2}%]" -f ($usrcount, $tot, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()

    Start-Sleep -Milliseconds 10
}
Write-Host -ForegroundColor Green " done"
$parsebar[0].Close()

<# *******************************************************************************
                                    QUERYING
                                DISTRIBUTION LISTS
******************************************************************************* #>
$DLs = @{}

# Getting static DLs
Write-Host "Getting static DLs..."
$GroupList = Get-MgGroup -All  -Property Id, DisplayName, Description, Mail, CreatedDateTime, GroupTypes, mailEnabled, securityEnabled `
    | Select-Object Id, DisplayName, Description, Mail, CreatedDateTime, GroupTypes, mailEnabled, securityEnabled
foreach ($aDL in $GroupList) {
    if (($aDL.GroupTypes -cnotcontains 'Unified') -and ($aDL.SecurityEnabled -ne 'True')) {
        Write-Host -ForegroundColor Yellow "  $($aDL.DisplayName)"
        $DLs["$($aDL.Id)"] = @{
            OBJECTID        = "$($aDL.Id)"
            UPN             = "$($aDL.Mail)"
            DISPLAYNAME     = "$($aDL.DisplayName)"
            TYPE            = 'static'
            CREATED         = $aDL.CreatedDateTime | Get-Date -format "yyyy/MM/dd"
            NOTES           = "$($aDL.Description)"
        }
    }
}
Write-Host -ForegroundColor Green 'done'

# disconnect from Tenant
Write-Host -NoNewline "Disconnecting from Tenant... "
$infoLogout = Disconnect-Graph
Start-Sleep -Milliseconds 3000
Write-Host -ForegroundColor Green "done"

Write-Host -NoNewline "Connecting to ExchangeOnLine... "
try {
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host -ForegroundColor Green "ok"
}
catch {
    Write-Host -ForegroundColor Red "Ko"
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

Write-Host -NoNewline "Getting dynamic DLs..."
foreach ($dyndl in Get-DynamicDistributionGroup) {
    Write-Host -NoNewline '.'
    $DLs["$($dyndl.ExchangeObjectId)"] = @{
        OBJECTID        = "$($dyndl.ExchangeObjectId)"
        UPN             = "$($dyndl.PrimarySmtpAddress)"
        DISPLAYNAME     = "$($dyndl.DisplayName)"
        TYPE            = 'dynamic'
        CREATED         =  $dynDL.WhenCreated | Get-Date -format "yyyy/MM/dd"
        NOTES           = "$($dyndl.Notes)"
    }
}
Write-Host -ForegroundColor Green ' done'

<# *******************************************************************************
                                    QUERYING
******************************************************************************* #>
Write-Host -NoNewline "Fetching mailbox list..."
<#
The Id (Identity) attribute obtained from Get-EXOMailbox cmdlet may vary, 
based on different kind of id types:
* Object ID           ie: 5c9bc377-137d-4b04-83d4-248b1ee6c705
* Security ID         ie: S-1-5-21-2262606477-833299019-2854854355-8406544
* User Principal Name ie: username@domain
* Legacy User Name    ie: username
#>

$EXOlist = Get-EXOMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize unlimited | ForEach-Object {
    Write-Host -NoNewline '.'        
    New-Object -TypeName PSObject -Property @{
        ID          = "$($_.Id)"
        TYPE        = "$($_.RecipientTypeDetails)"
        OBJ_ID      = "$($_.ExternalDirectoryObjectId)"
        UPN         = "$($_.UserPrincipalName)"
        DISPLAY     = "$($_.DisplayName)"
    } | Select ID, TYPE, OBJ_ID, UPN, DISPLAY
}
Write-Host -ForegroundColor Green " done"

Write-Host -NoNewline "Searching inactive mailbox..."
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
    Write-Host -ForegroundColor Green " none found"
}

<# *******************************************************************************
                                FILTERING
******************************************************************************* #>
# exclude list file is a simply list of UPNs
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
        EXCLUDED        = 'No'
        LASTLOGON       = 'na'
    }

    if ($ExcludeList.ContainsKey($entity.UPN)) {
        Write-Host -ForegroundColor DarkGray "  $($entity.UPN)"
        $EXOdetailed["$($entity.OBJ_ID)"].EXCLUDED = 'Yes'
        $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += 'SELF'
        $EXOdetailed["$($entity.OBJ_ID)"].LASTLOGON = '1980-02-07 12:00:00'
        Start-Sleep -Milliseconds 100
    } else {
        Write-Host -ForegroundColor Cyan "  $($entity.UPN)"

        $ErrorActionPreference= 'Stop'
        try {
            $mbs = Get-MailboxStatistics -Identity $entity.UPN | Select LastLogonTime
            $EXOdetailed["$($entity.OBJ_ID)"].LASTLOGON = "$(($mbs.LastLogonTime | Get-Date -format 'yyyy-MM-dd HH:mm:ss').ToString())"
        }
        catch {
            $EXOdetailed["$($entity.OBJ_ID)"].LASTLOGON = '1980-02-07 12:00:00'
        }
        $ErrorActionPreference= 'Inquire'

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
                            $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "'$($sandman.Trustee)',"
                        }
                        $ErrorActionPreference= 'Inquire'
                    } else {
                        if ($SIDfilter) {
                            $UPNfound = Get-ADUser -Filter * | Select-Object -Property SID,UserPrincipalName | Where-Object -Property SID -like "$($sandman.Trustee)"
                            if ($UPNfound -ne $null) {
                                $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$($UPNfound.UserPrincipalName)"
                            } elseif ($DisabledUsers.ContainsKey("$($sandman.Trustee)")) {
                                $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "'$($DisabledUsers["$($sandman.Trustee)"])',"
                            } else {
                                Write-Host -ForegroundColor Red "  granted '$($sandman.Trustee)' not found"
                                $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "'$($sandman.Trustee)',"
                            }
                        } else {
                            Write-Host -ForegroundColor Yellow "  granted '$($sandman.Trustee)' needs to be fixed"
                            $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$($sandman.Trustee)"
                        }
                    }
                } else {
                    $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "$($sandman.Trustee)"
                }
            }
        }

        foreach ($fuller in (Get-EXOMailboxPermission -id $entity.OBJ_ID)) {
            if ((($fuller.AccessRights -join ',') -cmatch 'FullAccess') -and ($fuller.InheritanceType -cne 'None')) {
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
                                $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += "'$($fuller.User)',"
                                
                            }
                            $ErrorActionPreference= 'Inquire'
                        } else {
                            if ($SIDfilter) {
                                $UPNfound = Get-ADUser -Filter * | Select-Object -Property SID,UserPrincipalName | Where-Object -Property SID -like "$($fuller.User)"
                                if ($UPNfound -ne $null) {
                                    $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += "$($UPNfound.UserPrincipalName)"
                                } elseif ($DisabledUsers.ContainsKey("$($fuller.User)")) {
                                     $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "'$($DisabledUsers["$($fuller.User)"])',"
                                } else {
                                    Write-Host -ForegroundColor Red "  granted '$($fuller.User)' not found"
                                    $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += "'$($fuller.User)',"
                                }
                            } else {
                                Write-Host -ForegroundColor Yellow "  granted '$($fuller.User)' needs to be fixed"
                                $EXOdetailed["$($entity.OBJ_ID)"].FULLACCESS += "$($fuller.User)"
                            }
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
                        $EXOdetailed["$($entity.OBJ_ID)"].SENDONBEHALF += "'$Beowulf',"
                    }
                    $ErrorActionPreference= 'Inquire'
                } else {
                    if ($SIDfilter) {
                        $UPNfound = Get-ADUser -Filter * | Select-Object -Property SID,UserPrincipalName | Where-Object -Property SID -like "$Beowulf"
                        if ($UPNfound -ne $null) {
                            $EXOdetailed["$($entity.OBJ_ID)"].SENDONBEHALF += "$Beowulf"
                        } elseif ($DisabledUsers.ContainsKey("$Beowulf")) {
                            $EXOdetailed["$($entity.OBJ_ID)"].SENDAS += "'$($DisabledUsers["$Beowulf"])',"
                        } else {
                            Write-Host -ForegroundColor Red "  granted '$Beowulf' not found"
                            $EXOdetailed["$($entity.OBJ_ID)"].SENDONBEHALF += "'$Beowulf',"
                        }
                    } else {
                        Write-Host -ForegroundColor Yellow "  granted '$Beowulf' needs to be fixed"
                        $EXOdetailed["$($entity.OBJ_ID)"].SENDONBEHALF += "$Beowulf"
                    }
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
Write-Host -ForegroundColor Green "done"
$parsebar[0].Close()

Disconnect-ExchangeOnline -Confirm:$false


<# *******************************************************************************
                                UPDATING NOTES
******************************************************************************* #>
$ManualNotes = @{}
$answ = [System.Windows.MessageBox]::Show("Do you have a xlsx template for updating notes?",'UPDATES','YesNo','Info')
if ($answ -eq 'Yes') {
    Write-Host "`n"
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Open File"
    $OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
    $OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
    $OpenFileDialog.ShowDialog() | Out-Null
    $xlsx_template = $OpenFileDialog.filename

    $Worksheet_list = Get-ExcelSheetInfo -Path $xlsx_template
    foreach ($templateWorksheet in ('UserMailbox', 'SharedMailbox', 'DistributionList')) {
        if ($Worksheet_list.Name -contains "$templateWorksheet") {
            $ManualNotes["$templateWorksheet"] = @{}
            Write-Host -NoNewline "Fetching [$templateWorksheet] data..."
            foreach ($history in (Import-Excel -Path $xlsx_template -WorksheetName "$templateWorksheet")) {
                $objID = $history.OBJECTID
                $note = $history.NOTES

                if (!($ManualNotes["$templateWorksheet"].ContainsKey("$objID"))) {
                    $($ManualNotes["$templateWorksheet"])["$objID"] = ($note | Out-String)
                    Write-Host -NoNewline '.'
                }
            }
            Write-Host -ForegroundColor Green ' done'
        } else {
            Write-Host -ForegroundColor Magenta "No [$templateWorksheet] worksheet found"
        }
    }   
}


<# *******************************************************************************
                                GET OUTPUT
******************************************************************************* #>
$xlsx_file = "C:$env:HOMEPATH\Downloads\ZombieMailboxSDK-" + (Get-Date -format "yyMMddHHmm") + '.xlsx'
$XlsPkg = Open-ExcelPackage -Path $xlsx_file -Create

$ErrorActionPreference= 'Stop'
try {  
    $label = 'ExcludedMailbox'
    Write-Host -NoNewline "Writing worksheet [$label]..."
    $inData = $EXOdetailed.Keys | Foreach-Object {
        if ($EXOdetailed[$_].EXCLUDED -eq 'Yes') {
            Write-Host -NoNewline '.'
            New-Object -TypeName PSObject -Property @{
                OBJECTID        = "$_"
                UPN             = "$($EXOdetailed[$_].UPN)"
                DISPLAYNAME     = "$($EXOdetailed[$_].DISPLAYNAME)"
            } | Select OBJECTID, UPN, DISPLAYNAME
        }
    }
    $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium1' -AutoSize -PassThru
    Write-Host -ForegroundColor Green ' done'

    $label = 'UserMailbox'
    Write-Host -NoNewline "Writing worksheet [$label]..."
    $inData = $EXOdetailed.Keys | Foreach-Object {
        if (($EXOdetailed[$_].EXCLUDED -eq 'No') -and ($EXOdetailed[$_].TYPE -eq 'UserMailbox')) {
            Write-Host -NoNewline '.'
            foreach ($GrantType in ('FULLACCESS', 'SENDAS', 'SENDONBEHALF')) {
                foreach ($Granted in (($EXOdetailed[$_])[$GrantType])) {
                    $aNote = 'null'
                    if ($ManualNotes.ContainsKey('UserMailbox')) {
                        if ($($ManualNotes.UserMailbox).ContainsKey("$_")) {
                            $aNote = $ManualNotes.UserMailbox["$_"]
                        }
                    }

                    New-Object -TypeName PSObject -Property @{
                        OBJECTID        = "$_"
                        UPN             = "$($EXOdetailed[$_].UPN)"
                        DISPLAYNAME     = "$($EXOdetailed[$_].DISPLAYNAME)"
                        LASTLOGON       = [DateTime]$EXOdetailed[$_].LASTLOGON
                        GRANT           = "$GrantType"
                        GRANTED         = "$Granted"
                        LICENSES        = "$($MsolUsrData["$($EXOdetailed[$_].UPN)"])"
                        NOTES           = "$($aNote.Trim())"
                    } | Select OBJECTID, UPN, DISPLAYNAME, LASTLOGON, GRANT, GRANTED, LICENSES, NOTES
                }
            }
        }
    }
    $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium2' -AutoSize -PassThru
    Write-Host -ForegroundColor Green ' done'

    $label = 'SharedMailbox'
    Write-Host -NoNewline "Writing worksheet [$label]..."
    $inData = $EXOdetailed.Keys | Foreach-Object {
        if (($EXOdetailed[$_].EXCLUDED -eq 'No') -and ($EXOdetailed[$_].TYPE -eq 'SharedMailbox')) {
            Write-Host -NoNewline '.'
            foreach ($GrantType in ('FULLACCESS', 'SENDAS', 'SENDONBEHALF')) {
                foreach ($Granted in (($EXOdetailed[$_])[$GrantType])) {
                    $aNote = 'null'
                    if ($ManualNotes.ContainsKey('SharedMailbox')) {
                        if ($($ManualNotes.SharedMailbox).ContainsKey("$_")) {
                            $aNote = $ManualNotes.SharedMailbox["$_"]
                        }
                    }

                    New-Object -TypeName PSObject -Property @{
                        OBJECTID        = "$_"
                        UPN             = "$($EXOdetailed[$_].UPN)"
                        DISPLAYNAME     = "$($EXOdetailed[$_].DISPLAYNAME)"
                        LASTLOGON       = [DateTime]$EXOdetailed[$_].LASTLOGON
                        GRANT           = "$GrantType"
                        GRANTED         = "$Granted"
                        LICENSES        = "$($MsolUsrData["$($EXOdetailed[$_].UPN)"])"
                        NOTES           = "$($aNote.Trim())"
                    } | Select OBJECTID, UPN, DISPLAYNAME, LASTLOGON, GRANT, GRANTED, LICENSES, NOTES
                }
            }
        }
    }
    $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium3' -AutoSize -PassThru
    Write-Host -ForegroundColor Green ' done'
    
    $label = 'DistributionList'
    Write-Host -NoNewline "Writing worksheet [$label]..."
    $inData = $DLs.Keys | Foreach-Object {
        Write-Host -NoNewline '.'

        $aNote = "$($DLs[$_].NOTES)"
        if ($ManualNotes.ContainsKey('DistributionList')) {
            if ($($ManualNotes.DistributionList).ContainsKey("$_")) {
                $aNote = $ManualNotes.DistributionList["$_"]
            }
        }

        New-Object -TypeName PSObject -Property @{
            OBJECTID        = "$_"
            UPN             = "$($DLs[$_].UPN)"
            DISPLAYNAME     = "$($DLs[$_].DISPLAYNAME)"
            TYPE            = "$($DLs[$_].TYPE)"
            CREATED         = "$($DLs[$_].CREATED)"
            NOTES           = "$($aNote.Trim())"
        } | Select OBJECTID, UPN, DISPLAYNAME, TYPE, CREATED, NOTES
    }
    $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium4' -AutoSize -PassThru
    Write-Host -ForegroundColor Green ' done'
} catch {
    [System.Windows.MessageBox]::Show("Error updating data",'ABORTING','Ok','Error') | Out-Null
    Write-Host -ForegroundColor Red ' FAIL'
    Write-Host -ForegroundColor Yellow "ERROR: $($error[0].ToString())"
    exit
}
$ErrorActionPreference= 'Inquire'

Close-ExcelPackage -ExcelPackage $XlsPkg

[System.Windows.MessageBox]::Show("File [$xlsx_file] has been created",'OUTPUT','Ok','Info') | Out-Null
