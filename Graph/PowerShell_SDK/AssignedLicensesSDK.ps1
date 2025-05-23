<#
Name......: AssignedLicensesSDK.ps1
Version...: 25.04.1
Author....: Dario CORRADA

This script will connect to the Microsoft 365 tenant and query a list of which 
license(s) are assigned to each user, then create/edit an excel report file.

Thx to Ali TAJRAN for the useful notes about Get-MgUser on:
https://www.alitajran.com/get-mguser/ 

TODO LIST:
* to build new pivot template 
  (happy ending option reserved for initial Excel reference files)
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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent | Split-Path -Parent

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'Microsoft.Graph')) {
            Install-Module Microsoft.Graph -Scope AllUsers -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [Microsoft.Graph] module: click Ok to restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } elseif (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
            Install-Module ImportExcel -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: click Ok restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') > $null
            Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                            CREDENTIALS MANAGEMENT
******************************************************************************* #>
Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Graph\AppKeyring.ps1"
if ($pswout.Count -eq 4) {
    $UPN = $pswout[0]
    $clientID = $pswout[1]
    $tenantID = $pswout[2]
    Write-Host -ForegroundColor Green ' Ok'
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# connect to Tenant
Write-Host -NoNewline "Connecting to the Tenant..."
$ErrorActionPreference= 'Stop'
Try {
    $splash = Connect-MgGraph -ClientId $clientID -TenantId $tenantID 
    Write-Host -ForegroundColor Green ' Ok'
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("Error connecting to the Tenant",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}


<# *******************************************************************************
                            FETCHING DATA FROM TENANT
******************************************************************************* #>
# retrieve the available licenses
Write-Host -NoNewline "Looking for available licenses..."
$avail_lics = @{}
$AccountName = (Get-MgSubscribedSku)[1].AccountName
foreach ($item in (Get-MgSubscribedSku | Select-Object -Property SkuPartNumber, ConsumedUnits -ExpandProperty PrepaidUnits)) {
    if ($item.Enabled -lt 10000) { # excluding the broadest licenses
        $avail_lics[$item.SkuPartNumber] = @{
            TOTAL   = $item.Enabled
            AVAIL   = ($item.Enabled - $item.ConsumedUnits) 
        }
    }
}
Write-Host -ForegroundColor Cyan " Found $($avail_lics.Count) active SKU"

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
    Write-Host -NoNewline "Getting data from [$($item.DisplayName)]... "     

    $MsolUsrData[$item.UserPrincipalName] = @{
        BLOCKED         = !($item.AccountEnabled)
        DESC            = $item.DisplayName
        USRNAME         = $item.UserPrincipalName
        LICENSED        = ($item.AssignedLicenses.Count -ge 1)
        LICENSES        = @{ # default values assuming no license assigned
            'NONE'        = Get-Date -format "yyyy/MM/dd"
        }
        USRTYPE         = $item.UserType
        CREATED         = $item.CreatedDateTime | Get-Date -format "yyyy/MM/dd"
    }

    if ($MsolUsrData[$item.UserPrincipalName].LICENSED -eq "True") {
        $MsolUsrData[$item.UserPrincipalName].LICENSES = @{} # re-init for updating licenses
        foreach ($accountsku in $item.AssignedLicenses) {            
            $SkuName = $SkuID2name[$accountsku.SkuID]
            if ($avail_lics.ContainsKey($SkuName)) { # filtering only managed licenses
                $MsolUsrData[$item.UserPrincipalName].LICENSES[$SkuName] = Get-Date -format "yyyy/MM/dd"
            }
        }
        Write-Host -ForegroundColor Blue "$($MsolUsrData[$item.UserPrincipalName].LICENSES.Count) license(s) assigned"
    } else {
        Write-Host -ForegroundColor Yellow "NO license assigned"
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
Write-Host -ForegroundColor Green " DONE"
$parsebar[0].Close()

# disconnect from Tenant
$infoLogout = Disconnect-Graph

<# *******************************************************************************
                            CREATING UPDATED DATAFRAMES
******************************************************************************* #>
# looking for Excel reference file
$UseRefFile = [System.Windows.MessageBox]::Show("Would you load an existing Excel reference file?",'UPDATING','YesNo','Info')
if ($UseRefFile -eq "Yes") {
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Open File"
    $OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
    $OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
    $OpenFileDialog.ShowDialog() | Out-Null
    $xlsx_file = $OpenFileDialog.filename
    $Worksheet_list = Get-ExcelSheetInfo -Path $xlsx_file
} else {
    $xlsx_file = "C:$env:HOMEPATH\Downloads\$($AccountName)_licenses.xlsx"
    [System.Windows.MessageBox]::Show("File [$xlsx_file] will be created",'CREATING','Ok','Info') | Out-Null
}
Write-Host -ForegroundColor Yellow "`nExcel reference file is [$xlsx_file]`n"

# [Licenses_Pool]
if ($UseRefFile -eq "Yes") {
    $timeline = Import-Excel -Path $xlsx_file -WorksheetName 'Licenses_Pool' | Select UPTIME | Get-Unique -AsString
    $adialog = FormBase -w 300 -h (($timeline.Count * 30) + 150) -text "TIMELINE"
    Label -form $adialog -x 20 -y 20 -w 250 -h 25 -text "[Licenses_Pool] records to keep:" | Out-Null
    $they = 40
    $choices = @()
    foreach ($adate in $timeline) {
        $choices += CheckBox -form $adialog -checked $true -x 50 -y $they -w 150 -text $($adate.UPTIME | Get-Date -Format "dd-MM-yyyy")
        $they += 30
    }
    OKButton -form $adialog -x 60 -y ($they + 15) -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
    $SaveTheDate = @()
    foreach ($item in $choices) {
        if ($item.Checked) {
            $SaveTheDate += $item.Text
        }
    }    
}

$Licenses_Pool_dataframe = @()
if ($UseRefFile -eq "Yes") { # appending older data
    if ($Worksheet_list.Name -contains 'Licenses_Pool') {
        Write-Host "Appending [Licenses_Pool] data..."
        foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Licenses_Pool')) {
            if ($SaveTheDate -contains ($history.UPTIME | Get-Date -format "dd-MM-yyyy")) {
                $Licenses_Pool_dataframe += ,@(
                    ($history.UPTIME | Get-Date -format "yyyy/MM/dd"),
                    $history.LICENSE,
                    $history.AVAILABLE,
                    $history.TOTAL
                )
            }
        }
    } else {
        Write-Host -ForegroundColor Magenta "No [Licenses_Pool] worksheet found"
    }
}
foreach ($item in $avail_lics.Keys) {
    $Licenses_Pool_dataframe += ,@(
        (Get-Date -format "yyyy/MM/dd"),
        $item,
        $avail_lics[$item].AVAIL,
        $avail_lics[$item].TOTAL
    )
}

# [Assigned_Licenses]
$orphanedrecords = @()
if ($UseRefFile -eq "Yes") {
    if ($Worksheet_list.Name -contains 'Assigned_Licenses') {
        Write-Host "Merging [Assigned_Licenses] data..."
        foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Assigned_Licenses')) {
            $aUser = $history.USRNAME
            if ($MsolUsrData.ContainsKey($aUser)) {
                $aLicense = $history.LICENSE
                if ($MsolUsrData[$aUser].LICENSES.ContainsKey($aLicense)) {
                    $OldTime = $history.TIMESTAMP | Get-Date -format "yyyy/MM/dd"
                    $NewTime = $MsolUsrData[$aUser].LICENSES[$aLicense]
                    <#
                    By default, the field TIMESTAMP refers to any change in license assignement.
                    If would like to also track the time in changing account proprerties substitue the
                    if clause as follows, for instance:
                    
                    if (($OldTime -lt $NewTime) -and ($MsolUsrData[$aUser].BLOCKED -eq $history.BLOCKED)) {
                    #>
                    if ($OldTime -lt $NewTime) {
                        $MsolUsrData[$aUser].LICENSES[$aLicense] = $OldTime
                    }
                } else {
                    Write-Host -ForegroundColor Yellow "[$aLicense] no longer assigned to [$aUser]"
                    if ($history.LICENSE -eq 'NONE') {
                        $anote = 'assigned license(s) to this user'
                    } else {
                        $anote = 'license dismissed for this user'
                    }
                    $orphanedrecords += ,@(
                        $history.USRNAME,
                        $history.DESC,
                        $history.USRTYPE,
                        ($history.CREATED | Get-Date -format "yyyy/MM/dd"),
                        $history.BLOCKED,
                        $history.LICENSED,
                        $history.LICENSE,
                        (Get-Date -format "yyyy/MM/dd"),
                        $anote
                    )
                }
            } else {
                Write-Host -ForegroundColor Yellow "[$aUser] no longer exists on tenant"
                $orphanedrecords += ,@(
                    $history.USRNAME,
                    $history.DESC,
                    $history.USRTYPE,
                    ($history.CREATED | Get-Date -format "yyyy/MM/dd"),
                    'NULL',
                    'NULL',
                    $history.LICENSE,
                    (Get-Date -format "yyyy/MM/dd"),
                    'user no longer exists on tenant'
                )
            }
        }
    } else {
        Write-Host -ForegroundColor Magenta "No [Assigned_Licenses] worksheet found"
    }
}
$Assigned_Licenses_dataframe = @()
foreach ($item in $MsolUsrData.Keys) {
    foreach ($subitem in $MsolUsrData[$item].LICENSES.Keys) {
        $Assigned_Licenses_dataframe += ,@(
            $MsolUsrData[$item].USRNAME,
            $MsolUsrData[$item].DESC,
            $MsolUsrData[$item].USRTYPE.ToString(),
            $MsolUsrData[$item].CREATED.ToString(),
            $MsolUsrData[$item].BLOCKED.ToString(),
            $MsolUsrData[$item].LICENSED.ToString(),
            $subitem,
            $MsolUsrData[$item].LICENSES[$subitem]
        )
    }
}

# [Orphaned]
$newOrphans = $false
if (($orphanedrecords.Count) -ge 1) {
    $newOrphans = $true
    if ($UseRefFile -eq "Yes") {
        if ($Worksheet_list.Name -contains 'Orphaned') {
            foreach ($currentRec in (Import-Excel -Path $xlsx_file -WorksheetName 'Orphaned')) {
                $orphanedrecords += ,@(
                    $currentRec.USRNAME,
                    $currentRec.DESC,
                    $currentRec.USRTYPE,
                    ($currentRec.CREATED | Get-Date -format "yyyy/MM/dd"),
                    $currentRec.BLOCKED,
                    $currentRec.LICENSED,
                    $currentRec.LICENSE,
                    ($currentRec.TIMESTAMP | Get-Date -format "yyyy/MM/dd"),
                    $currentRec.NOTES
                )
            }
        } else {
            Write-Host -ForegroundColor Magenta "No [Orphaned] worksheet found"
        }
    }
}


<# *******************************************************************************
                            WRITING REFERENCE FILE
******************************************************************************* #>
if ($UseRefFile -eq 'Yes') { 
    # create backup file
    $LastOfUs = $(Get-ChildItem $xlsx_file | Select-Object -Property LastWriteTime).LastWriteTime | Get-Date -format "yyMMdd_HHmm"
    $bkp_file = $xlsx_file + '-' + "$lastOfUs" + '.bkp'
    $bkp_file = $bkp_file.Replace('.xlsx', '')
    if (Test-Path -Path $bkp_file -PathType Leaf) {
        Remove-Item -Path $bkp_file -Force
    }
    Copy-Item -Path $xlsx_file -Destination $bkp_file -Force

    # remove older worksheets
    foreach ($currentSheet in $Worksheet_list) {
        if (($currentSheet.Name -eq 'Assigned_Licenses') `
        -or ($currentSheet.Name -eq 'Licenses_Pool') `
        -or (($newOrphans -eq $true) -and ($currentSheet.Name -eq 'Orphaned')) `
        -or ($currentSheet.Name -match "SkuCatalog")) {
            Remove-Worksheet -Path $xlsx_file -WorksheetName $currentSheet.Name
        }    
    }
    $XlsPkg = Open-ExcelPackage -Path $xlsx_file
} else {
    $XlsPkg = Open-ExcelPackage -Path $xlsx_file -Create
}


# writing SkuCatalog worksheet
$ErrorActionPreference= 'Stop'
try {
    $label = 'SkuCatalog'
    $csvdestfile = "C:$($env:HOMEPATH)\Downloads\$label.csv"
    if (Test-Path -Path $csvdestfile -PathType Leaf) { Remove-Item -Path $csvdestfile -Force }
    Invoke-WebRequest -Uri 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv' -OutFile "$csvdestfile"
    $SkuCatalog_rawdata = @{}
    foreach ($currentItem in (Import-Csv -Path $csvdestfile)) {
        if (!($SkuCatalog_rawdata.ContainsKey("$($currentItem.GUID)"))) {
            $SkuCatalog_rawdata["$($currentItem.GUID)"] = @{
                SKUID   = "$($currentItem.String_Id)"
                DESC    = "$($currentItem.Product_Display_Name)"
            }
        }
    }
    Write-Host -ForegroundColor Green "$($SkuCatalog_rawdata.Keys.Count) license type found"

    Write-Host -NoNewline "Writing worksheet [$label]..."
    $now = Get-Date -Format  "yyyy/MM/dd"
    $inData = $SkuCatalog_rawdata.Keys | Foreach-Object{
        Write-Host -NoNewline '.'        
        New-Object -TypeName PSObject -Property @{
            TIMESTAMP   = [DateTime]$now
            ID          = "$_"
            SKU         = "$($SkuCatalog_rawdata[$_].SKUID)"
            DESCRIPTION = "$($SkuCatalog_rawdata[$_].DESC)"
        } | Select TIMESTAMP, ID, SKU, DESCRIPTION
    }
    $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium1' -AutoSize -PassThru
    Write-Host -ForegroundColor Green ' DONE'
} catch {
    [System.Windows.MessageBox]::Show("Error updating data",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red ' FAIL'
    Write-Host -ForegroundColor Yellow "ERROR: $($error[0].ToString())"
    exit
}
$ErrorActionPreference= 'Inquire'


# writing Licenses_Pool worksheet
$ErrorActionPreference= 'Stop'
try {
    $label = 'Licenses_Pool'
    Write-Host -NoNewline "Writing worksheet [$label]..."
    $inData = 0..($Licenses_Pool_dataframe.Count - 1) | Foreach-Object{
        Write-Host -NoNewline '.'        
        New-Object -TypeName PSObject -Property @{
            UPTIME      = [DateTime]$Licenses_Pool_dataframe[$_][0]
            LICENSE     = $Licenses_Pool_dataframe[$_][1]
            AVAILABLE   = $Licenses_Pool_dataframe[$_][2]
            TOTAL       = $Licenses_Pool_dataframe[$_][3]
        } | Select UPTIME, LICENSE, AVAILABLE, TOTAL
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

# writing Assigned_Licenses worksheet
$ErrorActionPreference= 'Stop'
try {
    $label = 'Assigned_Licenses'
    Write-Host -NoNewline "Writing worksheet [$label]..."
    $inData = 0..($Assigned_Licenses_dataframe.Count - 1) | Foreach-Object{
        Write-Host -NoNewline '.'        
        New-Object -TypeName PSObject -Property @{
            USRNAME     = $Assigned_Licenses_dataframe[$_][0]
            DESC        = $Assigned_Licenses_dataframe[$_][1]
            USRTYPE     = $Assigned_Licenses_dataframe[$_][2]
            CREATED     = [DateTime]$Assigned_Licenses_dataframe[$_][3]
            BLOCKED     = $Assigned_Licenses_dataframe[$_][4]
            LICENSED    = $Assigned_Licenses_dataframe[$_][5]
            LICENSE     = $Assigned_Licenses_dataframe[$_][6]
            TIMESTAMP   = [DateTime]$Assigned_Licenses_dataframe[$_][7]
        } | Select USRNAME, DESC, USRTYPE, CREATED, BLOCKED, LICENSED, LICENSE, TIMESTAMP
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

# writing Orphaned worksheet
$ErrorActionPreference= 'Stop'
try {
    if ($orphanedrecords.Count -ge 1) {
        $label = 'Orphaned'
        Write-Host -NoNewline "Writing worksheet [$label]..."
        $inData = 0..($orphanedrecords.Count - 1) | Foreach-Object{
            Write-Host -NoNewline '.'        
            New-Object -TypeName PSObject -Property @{
                USRNAME     = $orphanedrecords[$_][0]
                DESC        = $orphanedrecords[$_][1]
                USRTYPE     = $orphanedrecords[$_][2]
                CREATED     = [DateTime]$orphanedrecords[$_][3]
                BLOCKED     = $orphanedrecords[$_][4]
                LICENSED    = $orphanedrecords[$_][5]
                LICENSE     = $orphanedrecords[$_][6]
                TIMESTAMP   = [DateTime]$orphanedrecords[$_][7]
                NOTES       = $orphanedrecords[$_][8]
            } | Select USRNAME, DESC, USRTYPE, CREATED, BLOCKED, LICENSED, LICENSE, TIMESTAMP, NOTES
        }
        $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium3' -AutoSize -PassThru
        Write-Host -ForegroundColor Green ' DONE'
    }
} catch {
    [System.Windows.MessageBox]::Show("Error updating data",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red ' FAIL'
    Write-Host -ForegroundColor Yellow "ERROR: $($error[0].ToString())"
    exit
}
$ErrorActionPreference= 'Inquire'

<# MOTA add-on
Add a worksheet which produce summary amount report of individual licenses 
assigned vs. that ones recovered (aka available to be assigned again).
See also the "/AzureAD/MOTA.ps1" script for more details.
#>
Close-ExcelPackage -ExcelPackage $XlsPkg

# remove older worksheet
if ($Worksheet_list.Name -contains 'MOTA') {
    Remove-Worksheet -Path $xlsx_file -WorksheetName 'MOTA'
    Write-Host -NoNewline "Updating worksheet [MOTA]..."
} else {
    Write-Host -NoNewline "Adding worksheet [MOTA]..."
}

# launch MOTA and import related csv file
$MOTAcsvfile = PowerShell.exe -file "$workdir\AzureAD\MOTA.ps1" -InFile $xlsx_file
# by default $MOTAcsvfile should be "C:$($env:HOMEPATH)\Downloads\mota.csv"

$inData = Import-Csv $MOTAcsvfile | ForEach-Object {
    Write-Host -NoNewline '.'        
    New-Object -TypeName PSObject -Property @{
        TIMESTAMP   = [DateTime]$_.TIMESTAMP
        ACCOUNT     = $_.ACCOUNT
        LICENSE     = $_.LICENSE
        STATUS      = $_.STATUS
    } | Select TIMESTAMP, ACCOUNT, LICENSE, STATUS
}
$XlsPkg = Open-ExcelPackage -Path $xlsx_file
$XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName 'MOTA' -TableName 'MOTA' -TableStyle 'Medium4' -AutoSize -PassThru
Write-Host -ForegroundColor Green ' DONE'

# resorting worksheets
$XlsPkg.Workbook.Worksheets.MoveToStart('SkuCatalog')
$XlsPkg.Workbook.Worksheets.MoveAfter('Licenses_Pool', 'Skucatalog')
$XlsPkg.Workbook.Worksheets.MoveAfter('Assigned_Licenses', 'Licenses_Pool')
if ($XlsPkg.Workbook.Worksheets.Name -contains 'Orphaned') {
    $XlsPkg.Workbook.Worksheets.MoveAfter('Orphaned', 'Assigned_Licenses')
    $XlsPkg.Workbook.Worksheets.MoveAfter('MOTA', 'Orphaned')
} else {
    $XlsPkg.Workbook.Worksheets.MoveAfter('MOTA', 'Assigned_Licenses')
}

<# *******************************************************************************
                                HAPPY ENDING
******************************************************************************* #>
$HappyEnding = @{}
$form_panel = FormBase -w 300 -h 220 -text "HAPPY ENDING"
if ($UseRefFile -eq 'No') {
    # enable this feature once implemented the new release
    $HappyEnding['PIVOT'] = CheckBox -form $form_panel -checked $true -x 50 -y 20 -text "Add summary pivot template"
    $HappyEnding['PREVIEW'] = CheckBox -form $form_panel -checked $true -x 50 -y 50 -text "Open Excel formatted file"
    $HappyEnding['CLEANSWEEP'] = CheckBox -form $form_panel -checked $false -enabled $false -x 50 -y 80 -text "Remove temporary backup"
} else {
    $HappyEnding['PIVOT'] = CheckBox -form $form_panel -checked $false -x 50 -y 20 -text "Add summary pivot template"
    $HappyEnding['PREVIEW'] = CheckBox -form $form_panel -checked $false -x 50 -y 50 -text "Open Excel formatted file"
    $HappyEnding['CLEANSWEEP'] = CheckBox -form $form_panel -checked $true -x 50 -y 80 -text "Remove temporary backup"
}
OKButton -form $form_panel -x 80 -y 130 -text "Ok"  | Out-Null
$result = $form_panel.ShowDialog()

if ($HappyEnding['PIVOT'].Checked -eq $true) {
    #Close-ExcelPackage -ExcelPackage $XlsPkg
    #$XlsPkg = Open-ExcelPackage -Path $xlsx_file

    Add-Worksheet -ExcelPackage $XlsPkg -WorksheetName 'SUMMARY' > $null

    Add-PivotTable -ExcelPackage $XlsPkg `
    -PivotTableName 'POOL' -Address $XlsPkg.SUMMARY.cells["B3"] `
    -SourceWorksheet 'Licenses_Pool' `
    -PivotRows 'LICENSE' -PivotColumns 'UPTIME' -PivotData @{AVAILABLE="Sum";TOTAL="Sum"} `
    -PivotTableStyle 'Dark7' -PivotTotals 'Rows'

    $placeholder = ($avail_lics.Count * 3) + 11
    Add-PivotTable -ExcelPackage $XlsPkg `
    -PivotTableName 'ASSIGNEMENTS' -Address $XlsPkg.SUMMARY.cells["B$placeholder"] `
    -SourceWorksheet 'MOTA' `
    -PivotRows ('LICENSE', 'ACCOUNT') -PivotColumns 'TIMESTAMP' -PivotData 'LICENSE' `
    -PivotFilter 'STATUS' `
    -GroupDateColumn 'TIMESTAMP' -GroupDatePart 'Years,Months' `
    -PivotTableStyle 'Dark3' -PivotTotals 'Rows' 
}

if ($HappyEnding['PREVIEW'].Checked -eq $true) {
    Close-ExcelPackage -ExcelPackage $XlsPkg -Show
} else {
    Close-ExcelPackage -ExcelPackage $XlsPkg
}

if (($HappyEnding['CLEANSWEEP'].Checked -eq $true) -and ($UseRefFile -eq 'Yes'))  {
    Remove-Item -Path $bkp_file -Force
}

# REMOVE RENMANTS
# for those cases in which the xlsx file is mapped onto OneDrive synced paths
$twodots = Split-Path -Parent $xlsx_file | Split-Path -Parent
$filename = Split-Path -Leaf $xlsx_file
if (Test-Path "$twodots/$filename" -PathType Leaf) {
    Remove-Item -Path "$twodots/$filename" -Force
}
