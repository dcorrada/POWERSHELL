<#
Name......: AssignedLicenses_nologin.ps1
Version...: 24.06.2
Author....: Dario CORRADA

This script is a recovery version of <AssigneLicenses.ps1>, it requires either 
a .csv file retrieved from the Microsoft 365 tenant web interface and .xlsx 
history reference file.
#>


<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AzureAD\\AssignedLicenses_nologin\.ps1$" > $null
$workdir = $matches[1]
<# for testing purposes
$workdir = Get-Location
$workdir = $workdir.Path
#>

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
        if (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
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
                            FETCHING DATA FROM TENANT
******************************************************************************* #>
$answ = [System.Windows.MessageBox]::Show("Load .csv file from tenant",'INFILE','Ok','Info')
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Open File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'CSV file (*.csv)| *.csv'
$OpenFileDialog.ShowDialog() | Out-Null
$infile = $OpenFileDialog.filename
# sostituisco l'header del file per non avere lettere accentate e spazi
$Heather = 'blocked', 'city', 'state', 'dept', 'DirSyncEnabled', 'DisplayName',
    'Fax', 'name', 'LastDirSync', 'surname', 'LastPasswordSet', 'details', 'licenses',
    'phone', 'MetadataTokenOATH', 'ObjID', 'Office', 'NeverExpire', 'phone2', 'CAP',
    'preferred', 'language','proxy', 'deploy', 'eliminazione_temporanea', 'country',
    'addres', 'pswdComplex', 'title', 'zone', 'UPN', 'created'
$A = Get-Content -Path $infile
$A = $A[1..($A.Count - 1)]
$A | Out-File -FilePath "C:\Users\$env:USERNAME\Downloads\headless.csv"
$UsrRawdata = Import-Csv -Path "C:\Users\$env:USERNAME\Downloads\headless.csv" -Header $Heather

# retrieve all users list
$MsolUsrData = @{} 
$tot = $UsrRawdata.Count
$usrcount = 0
$parsebar = ProgressBar
foreach ($item in $UsrRawdata) {
    $usrcount ++
    Write-Host -NoNewline "Getting data from [$($item.DisplayName)]... "     

    if ([string]::IsNullOrEmpty($item.licenses)) {
        $fired = 'False'
    } else {
        $fired = 'True'
    }

    $MsolUsrData[$item.UPN] = @{
        BLOCKED         = $item.blocked
        DESC            = $item.DisplayName
        USRNAME         = $item.UPN
        LICENSED        = $fired
        LICENSES        = @{ # default values assuming no license assigned
            'NONE'        = Get-Date -format "yyyy/MM/dd"
        }
        USRTYPE         = 'na' # no info available from .csv
        CREATED         = ($item.created | Get-Date -format "yyyy/MM/dd")
    }

    if ($MsolUsrData[$item.UPN].LICENSED -eq "True") {
        $MsolUsrData[$item.UPN].LICENSES = @{} # re-init for updating licenses
        $LicList = $item.licenses.Split('+')

        foreach ($aLic in $LicList) {
            $MsolUsrData[$item.UPN].LICENSES[$aLic] = Get-Date -format "yyyy/MM/dd"
        }
        Write-Host -ForegroundColor Blue "$($MsolUsrData[$item.UPN].LICENSES.Count) license(s) assigned"
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


<# *******************************************************************************
                            CREATING UPDATED DATAFRAMES
******************************************************************************* #>
$answ = [System.Windows.MessageBox]::Show("Load .xlsx reference file",'INFILE','Ok','Info')
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Open File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
$OpenFileDialog.ShowDialog() | Out-Null
$xlsx_file = $OpenFileDialog.filename

# [Licenses_Pool]
$timeline = Import-Excel -Path $xlsx_file -WorksheetName 'Licenses_Pool' | Select UPTIME | Get-Unique -AsString
$adialog = FormBase -w 300 -h (($timeline.Count * 30) + 150) -text "TIMELINE"
Label -form $adialog -x 20 -y 20 -w 200 -h 25 -text "[Licenses_Pool] records to keep:" | Out-Null
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
$Licenses_Pool_dataframe = @()

Write-Host -NoNewline "Fetching data from [Licenses_Pool]..."
$avail_lics = @{}
foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Licenses_Pool')) {
    if ($SaveTheDate -contains ($history.UPTIME | Get-Date -format "dd-MM-yyyy")) {
        $Licenses_Pool_dataframe += ,@(
            ($history.UPTIME | Get-Date -format "yyyy/MM/dd"),
            $history.LICENSE,
            $history.AVAILABLE,
            $history.TOTAL
        )
    }
    if ($history.UPTIME -eq $timeline[$timeline.Count-1].UPTIME) {
        $avail_lics[$history.LICENSE] = @{
            TOTAL   = $history.TOTAL
            AVAIL   = $history.AVAILABLE 
        }
    }
}
Write-Host -ForegroundColor Green " DONE"

Write-Host -NoNewline "Update data to [Licenses_Pool]..."
$they = 60
$UpPool = @{}
$UpdatePoolForm = FormBase -w 580 -h ((($avail_lics.Count) * 40) + 180)  -text 'UPDATE LICENSE POOL'
Label -form $UpdatePoolForm -x 20 -y 15 -w 300 -text 'LICENSE' | Out-Null
Label -form $UpdatePoolForm -x 330 -y 15 -w 80 -text 'TOTAL' | Out-Null
Label -form $UpdatePoolForm -x 430 -y 15 -w 80 -text 'AVAILABLE' | Out-Null
foreach ($currentLic in ($avail_lics.Keys | Sort-Object)) {
    Label -form $UpdatePoolForm -x 20 -y $they -w 300 -text $currentLic | Out-Null
    $UpPool[$currentLic] = @(
        (TxtBox -form $UpdatePoolForm -x 330 -y $they -w 50 -text $avail_lics[$currentLic].TOTAL),
        (TxtBox -form $UpdatePoolForm -x 430 -y $they -w 50 -text $avail_lics[$currentLic].AVAIL)
    )
    $they += 40
}
OKButton -form $UpdatePoolForm -x 210 -y ($they + 20) -text "Ok" | Out-Null
$result = $UpdatePoolForm.ShowDialog()
foreach ($currentLic in ($UpPool.Keys | Sort-Object)) {
    $Licenses_Pool_dataframe += ,@(
        (Get-Date -format "yyyy/MM/dd"),
        $currentLic,
        ($UpPool[$currentLic])[1].text,
        ($UpPool[$currentLic])[0].text
    )
}
Write-Host -ForegroundColor Green " DONE"

# [SkuCatalog]
Write-Host -NoNewline 'Getting SKUs...'
$csvdestfile = "C:$($env:HOMEPATH)\Downloads\SkuCatalog.csv"
if (Test-Path -Path $csvdestfile -PathType Leaf) { Remove-Item -Path $csvdestfile -Force }
Invoke-WebRequest -Uri 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv' -OutFile "$csvdestfile"
$SkuCatalog_rawdata = @{}
$MsolUsrData_translates = @{}
$tot = (Import-Csv -Path $csvdestfile).Count - 1
$usrcount = 0
$parsebar = ProgressBar
foreach ($currentItem in (Import-Csv -Path $csvdestfile)) {
    $usrcount ++
    if (!($SkuCatalog_rawdata.ContainsKey("$($currentItem.GUID)"))) {
        $SkuCatalog_rawdata["$($currentItem.GUID)"] = @{
            SKUID   = "$($currentItem.String_Id)"
            DESC    = "$($currentItem.Product_Display_Name)"
        }
        if ($avail_lics.Keys -contains "$($currentItem.String_Id)") {
            $MsolUsrData_translates["$($currentItem.Product_Display_Name)"] = "$($currentItem.String_Id)"
        }
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
}
Write-Host -ForegroundColor Blue " $($SkuCatalog_rawdata.Keys.Count) license type found"
$parsebar[0].Close()
Start-Sleep -Milliseconds 1000

# Patch $MsolUsrData (convert license description into SKUID)
foreach ($aUPN in $MsolUsrData.Keys) {
    if ($MsolUsrData[$aUPN].LICENSED -eq 'True') {
        $SKUlist = @()
        foreach ($aDesc in $MsolUsrData[$aUPN].LICENSES.Keys) {
            # specific translation from italian tenant version
            $aDesc = $aDesc.replace('Piano','Plan')
            $aDesc = $aDesc.replace('Visio - Plan 2','Visio Online Plan 2')
            $aDesc = $aDesc.replace('Planner Plan 1','Project Plan 1')

            if ($MsolUsrData_translates.ContainsKey($aDesc)) {
                $SKUlist += "$($MsolUsrData_translates[$aDesc])"
            } elseif (!(('Microsoft Fabric (gratuito)', # blacklisted SKUs
                'Microsoft Power Automate Free', 
                'Microsoft Teams Exploratory') -contains "$aDesc")) {

                # raise warning: this SKU is not translated nor blacklisted
                Write-Host -ForegroundColor Yellow "untracked [$aDesc] assigned to [$aUPN]"
                [System.Windows.MessageBox]::Show(@"
license [$aDesc]
assigned to [$aUPN]
NOT TRACKED YET

Please edit your Excel reference file once script have been closed
"@,'NEW SKU','Ok','Warning') | Out-Null
                $SKUlist += "[TO_FIX] $aDesc"
            }
        }
        $MsolUsrData[$aUPN].LICENSES = @{} # re-init for patching SKUs
        if ($SKUlist.Count -eq 0) {
            $MsolUsrData[$aUPN].LICENSES = @{
                'NONE'        = Get-Date -format "yyyy/MM/dd"
            }
        } else {
            foreach ($currentSKU in $SKUlist) {
                $MsolUsrData[$aUPN].LICENSES[$currentSKU] = Get-Date -format "yyyy/MM/dd"
            }
        }
    }
}


# [Assigned_Licenses]
$orphanedrecords = @()
Write-Host "Merging [Assigned_Licenses] data..."
foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Assigned_Licenses')) {
    $aUser = $history.USRNAME
    if ($MsolUsrData.ContainsKey($aUser)) {
        if ($history.USRTYPE -ne 'na') {
            $MsolUsrData[$aUser].USRTYPE = $history.USRTYPE
        }
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
Pause

foreach ($currentUPN in $MsolUsrData.Key) {
    if ($MsolUsrData[$currentUPN].USRTYPE -eq 'na') {
        $typefixForm = FormBase -w 400 -h 160  -text 'NEW USER TYPE'
        Label -form $typefixForm -x 20 -y 15 -w 200 -h 20 -text 'UPN' | Out-Null
        Label -form $typefixForm -x 240 -y 15 -w 60 -h 20 -text 'Member' | Out-Null
        Label -form $typefixForm -x 310 -y 15 -w 60 -h 20 -text 'Guest' | Out-Null
        Label -form $typefixForm -x 20 -y 40 -w 220 -text $currentUPN | Out-Null
        $okMember = RadioButton -form $typefixForm -x 255 -y 33 -w 30 -checked $true
        $okGuest = RadioButton -form $typefixForm -x 325 -y 33 -w 30 -checked $false
        OKButton -form $typefixForm -x 130 -y 80 -text "Ok" | Out-Null
        $result = $typefixForm.ShowDialog()
        if ($okMember.Checked) {
            $MsolUsrData[$currentUPN].USRTYPE = 'Member'
        } elseif ($okGuest.Checked) {
            $MsolUsrData[$currentUPN].USRTYPE = 'Guest'
        }
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



<# *******************************************************************************
                            WRITING REFERENCE FILE
******************************************************************************* #>
# create backup file    
$bkp_file = $xlsx_file + '.bkp'
if (Test-Path -Path $bkp_file -PathType Leaf) {
    Remove-Item -Path $bkp_file -Force
}
Copy-Item -Path $xlsx_file -Destination $bkp_file -Force

# remove older worksheets
foreach ($currentSheet in $Worksheet_list) {
    if (($currentSheet.Name -eq 'Assigned_Licenses') `
    -or ($currentSheet.Name -eq 'Licenses_Pool') `
    -or ($currentSheet.Name -eq 'Orphaned') `
    -or ($currentSheet.Name -match "SkuCatalog")) {
        Remove-Worksheet -Path $xlsx_file -WorksheetName $currentSheet.Name
    }    
}
$XlsPkg = Open-ExcelPackage -Path $xlsx_file


# writing SkuCatalog worksheet
$ErrorActionPreference= 'Stop'
try {
    $label = 'SkuCatalog'
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

# resorting worksheets
$XlsPkg.Workbook.Worksheets.MoveToStart('SkuCatalog')
$XlsPkg.Workbook.Worksheets.MoveAfter('Licenses_Pool', 'Skucatalog')
$XlsPkg.Workbook.Worksheets.MoveAfter('Assigned_Licenses', 'Licenses_Pool')
if ($XlsPkg.Workbook.Worksheets.Name -contains 'Orphaned') {
    $XlsPkg.Workbook.Worksheets.MoveAfter('Orphaned', 'Assigned_Licenses')
}

Close-ExcelPackage -ExcelPackage $XlsPkg

# remove renmants
$twodots = Split-Path -Parent $xlsx_file | Split-Path -Parent
$filename = Split-Path -Leaf $xlsx_file
if (Test-Path "$twodots/$filename" -PathType Leaf) {
    Remove-Item -Path "$twodots/$filename" -Force
}
