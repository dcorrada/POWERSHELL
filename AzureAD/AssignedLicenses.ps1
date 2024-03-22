<#
Name......: AssignedLicenses.ps1
Version...: 24.04.alfa
Author....: Dario CORRADA

This script will connect to the Microsoft 365 tenant and query a list of which 
license(s) are assigned to each user, then create/edit an excel report file.

*** ONGOING: ***
-> adopt the PSExcel module to wrote excel file, instead of using COM object.
    https://ramblingcookiemonster.github.io/PSExcel-Intro/
    https://www.powershellgallery.com/packages/PSExcel

*** TODO LIST: ***
-> add pivot tables into excel file
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

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AzureAD\\AssignedLicenses\.ps1$" > $null
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
try {
    Import-Module -Name "$workdir\Modules\Gordian.psm1"
    Import-Module -Name "$workdir\Modules\Forms.psm1"
    Import-Module MSOnline
    Import-Module PSExcel
} catch {
    if (!(((Get-InstalledModule).Name) -contains 'MSOnline')) {
        Install-Module MSOnline -Confirm:$False -Force
        [System.Windows.MessageBox]::Show("Installed [MSOnline] module: please restart the script",'RESTART','Ok','warning')
        exit
    } elseif (!(((Get-InstalledModule).Name) -contains 'PSExcel')) {
        Install-Module PSExcel -Confirm:$False -Force
        [System.Windows.MessageBox]::Show("Installed [PSExcel] module: please restart the script",'RESTART','Ok','warning')
        exit
    } else {
        [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error')
        Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
        Pause
        exit
    }
}
$ErrorActionPreference= 'Inquire'


<# *******************************************************************************
                            CREDENTIALS MANAGEMENT
******************************************************************************* #>
# the db files
$usedb = $true
$dbfile = $env:LOCALAPPDATA + '\AssignedLicenses.encrypted'
$keyfile = $env:LOCALAPPDATA + '\AssignedLicenses.key'

# looking for existing DB file
if (Test-Path $dbfile -PathType Leaf) {
    $adialog = FormBase -w 350 -h 170 -text "DATABASE"
    if (Test-Path $keyfile -PathType Leaf) {
        $enterDB = RadioButton -form $adialog -checked $true -x 20 -y 20 -w 500 -h 30 -text "Enter the DB file"
        $cleanDB = RadioButton -form $adialog -checked $false -x 20 -y 50 -w 500 -h 30 -text "Delete the DB file"
    } else {
        $enterDB = RadioButton -form $adialog -enabled $false -checked $false -x 20 -y 20 -w 500 -h 30 -text "Enter the DB file (NO key to decrypt!)"
        $cleanDB = RadioButton -form $adialog -checked $true -x 20 -y 50 -w 500 -h 30 -text "Delete the DB file"
    }
    OKButton -form $adialog -x 100 -y 90 -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
    if ($cleanDB.Checked -eq $true) {
        $answ = [System.Windows.MessageBox]::Show("Really delete DB file?",'DELETE','YesNo','Warning')
        if ($answ -eq "Yes") {    
            Remove-Item -Path $dbfile
        }
    }
}

if (!(Test-Path $dbfile -PathType Leaf)) {
    $answacc = [System.Windows.MessageBox]::Show("No DB file found.`nDo you want accessing tenant directly?",'ACCESS','YesNo','Info')
    if ($answacc -eq "Yes") {    
        $usedb = $false
        $singlelogin = LoginWindow
    } else {
            # creating key file if not available
        if (!(Test-Path $keyfile -PathType Leaf)) {
            CreateKeyFile -keyfile "$keyfile" | Out-Null
        }

        # creating DB file
        $adialog = FormBase -w 400 -h 300 -text "DB INIT"
        Label -form $adialog -x 20 -y 20 -w 500 -h 30 -text "Initialize your DB as follows (NO space allowed)" | Out-Null
        $dbcontent = TxtBox -form $adialog -x 20 -y 50 -w 300 -h 150 -text ''
        $dbcontent.Multiline = $true;
        $dbcontent.Text = @'
USR;PWD
user1@foobar.baz;password1
user2@foobar.baz;password2
'@
        $dbcontent.AcceptsReturn = $true
        OKButton -form $adialog -x 100 -y 220 -text "Ok" | Out-Null
        $result = $adialog.ShowDialog()
        $tempusfile = $env:LOCALAPPDATA + '\AssignedLicenses.csv'
        $dbcontent.Text | Out-File $tempusfile
        EncryptFile -keyfile "$keyfile" -infile "$tempusfile" -outfile "$dbfile" | Out-Null
    }
}

if ($usedb -eq $true) {
    # reading DB file
    $filecontent = (DecryptFile -keyfile "$keyfile" -infile "$dbfile").Split(" ")
    $allowed = @{}
    foreach ($newline in $filecontent) {
        if ($newline -ne 'USR;PWD') {
            ($username, $passwd) = $newline.Split(';')
            $allowed[$username] = $passwd
        }
    }

    # select the account to access
    $adialog = FormBase -w 350 -h (($allowed.Count * 30) + 120) -text "SELECT AN ACCOUNT"
    $they = 20
    $choices = @()
    foreach ($username in $allowed.Keys) {
        if ($they -eq 20) {
            $isfirst = $true
        } else {
            $isfirst = $false
        }
        $choices += RadioButton -form $adialog -x 20 -y $they -w 300 -checked $isfirst -text $username
        $they += 30
    }
    OKButton -form $adialog -x 100 -y ($they + 10) -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
}


<# *******************************************************************************
                            FETCHING DATA FROM TENANT
******************************************************************************* #>
# get credentials for accessing
if ($usedb -eq $true) {
    foreach ($item in $choices) {
        if ($item.Checked) {
            $usr = $item.Text
            $plain_pwd = $allowed[$usr]
        }
    }
    $pwd = ConvertTo-SecureString $plain_pwd -AsPlainText -Force
    $credits = New-Object System.Management.Automation.PSCredential($usr, $pwd)
} else {
    $credits = $singlelogin
}


# connect to Tenant
Write-Host -NoNewline "Connecting to the Tenant..."
$ErrorActionPreference= 'Stop'
Try {
    Connect-MsolService -Credential $credits
    Write-Host -ForegroundColor Green "Ok"
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("Error connecting to the Tenant",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# retrieve the available licenses, complete list of possible account sku is available on:
# https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
Write-Host -NoNewline "Looking for available licenses..."
$avail_lics = @{}
$AccountName = (Get-MsolAccountSku)[1].AccountName
foreach ($item in (Get-MsolAccountSku)) {
    if ($item.ActiveUnits -lt 10000) { # excluding the broadest licenses
        $avail_lics[$item.SkuPartNumber] = @{
            TOTAL   = $item.ActiveUnits
            AVAIL   = ($item.ActiveUnits - $item.ConsumedUnits) 
        }
    }
}
Write-Host " Found $($avail_lics.Count) active SKU"

# retrieve all users list
$MsolUsrData = @{} 
$tot = (Get-MsolUser -All).Count
$usrcount = 0
$parsebar = ProgressBar
foreach ($item in (Get-MsolUser -All | Sort-Object DisplayName)) {
    $usrcount ++
    Write-Host -NoNewline "Getting data from [$($item.DisplayName)]... "     

    $MsolUsrData[$item.UserPrincipalName] = @{
        BLOCKED         = $item.BlockCredential
        DESC            = $item.DisplayName
        USRNAME         = $item.UserPrincipalName
        LICENSED        = $item.IsLicensed
        LICENSES        = @{ # default values assuming no license assigned
            'NONE'        = Get-Date -format "yyyy/MM/dd"
        }
        USRTYPE         = $item.UserType
        CREATED         = $item.WhenCreated | Get-Date -format "yyyy/MM/dd"
    }

    if ($MsolUsrData[$item.UserPrincipalName].LICENSED -eq "True") {
        $MsolUsrData[$item.UserPrincipalName].LICENSES = @{} # re-init for updating licenses
        foreach ($accountsku in $item.Licenses.AccountSku.SkuPartNumber) {
            if ($avail_lics.ContainsKey($accountsku)) { # filtering only managed licenses
                $MsolUsrData[$item.UserPrincipalName].LICENSES[$accountsku] = Get-Date -format "yyyy/MM/dd"
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
}
Write-Host -ForegroundColor Green " DONE"
$parsebar[0].Close()



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
} else {
    $xlsx_file = "C:$env:HOMEPATH\Downloads\$($AccountName)_licenses.xlsx"
    [System.Windows.MessageBox]::Show("File [$xlsx_file] will be created",'CREATING','Ok','Info') | Out-Null
}
Write-Host -ForegroundColor Yellow "`nExcel reference file is [$xlsx_file]`n"

# Open Excel reference file
$XlsxObj = New-Excel -Path $xlsx_file
$Worksheet_list = $XlsxObj | Get-Worksheet

# [Licenses_Pool]
$Licenses_Pool_dataframe = @()
if ($UseRefFile -eq "Yes") { # appending older data
    if ($Worksheet_list.Name -contains 'Licenses_Pool') {
        Write-Host "Appending [Licenses_Pool] data..."
        foreach ($history in (Import-XLSX -Path $xlsx_file -Sheet 'Licenses_Pool')) {
            $Licenses_Pool_dataframe += ,@(
                ($history.UPTIME | Get-Date -format "yyyy/MM/dd"),
                $history.LICENSE,
                $history.AVAILABLE,
                $history.TOTAL
            )
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
        foreach ($history in (Import-XLSX -Path $xlsx_file -Sheet 'Assigned_Licenses')) {
            $aUser = $history.USRNAME
            if ($MsolUsrData.ContainsKey($aUser)) {
                $aLicense = $history.LICENSE
                if ($MsolUsrData[$aUser].LICENSES.ContainsKey($aLicense)) {
                    $OldTime = $history.TIMESTAMP | Get-Date -format "yyyy/MM/dd"
                    $NewTime = $MsolUsrData[$aUser].LICENSES[$aLicense]
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
                    'NONE',
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
            foreach ($currentRec in (Import-XLSX -Path $xlsx_file -Sheet 'Orphaned')) {
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
$Myexcel = New-Object -ComObject excel.application
$Myexcel.Visible = $false
$Myexcel.DisplayAlerts = $false
$FetchSkuCatalog = $false
if ($UseRefFile -eq 'Yes') { # remove older worksheets
    $ReplaceSkuCatalog = [System.Windows.MessageBox]::Show("Update [SkuCatalog] worksheet",'UPDATING','YesNo','Info')
    if ($ReplaceSkuCatalog -eq 'Yes') {
        $FetchSkuCatalog = $true
    }
    $Myworkbook = $Myexcel.Workbooks.Open($xlsx_file)
    foreach ($currentSheet in ($Myworkbook.Worksheets)) {
        if (($currentSheet.Name -eq 'Assigned_Licenses') -or ($currentSheet.Name -eq 'Licenses_Pool')) {
            $currentSheet.Delete()
        } 
        if (($ReplaceSkuCatalog -eq 'Yes') -and ($currentSheet.Name -match "SkuCatalog")) {
            $currentSheet.Delete()
        }
        if (($newOrphans -eq $true) -and ($currentSheet.Name -match "Orphaned")) {
            $currentSheet.Delete()
        }       
    }
} else { # create new file
    $Myworkbook = $Myexcel.Workbooks.Add()
    $FetchSkuCatalog = $true
}

# writing SkuCatalog worksheet
if ($FetchSkuCatalog -eq $true) {
    $label = 'SkuCatalog_' + (Get-Date -Format "yyMMdd")
    $csvdestfile = "C:$($env:HOMEPATH)\Downloads\$label.csv"
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
    $Sheet3 = $Myworkbook.Worksheets.add()
    $Sheet3.name = "$label"
    $i = 1
    foreach ($item in ('ID','SKU','DESCRIPTION')) {
        $Sheet3.cells.item(1,$i) = $item
        $i++        
    }
    $i = 2
    $tot = $SkuCatalog_rawdata.Keys.Count
    $usrcount = 0
    $parsebar = ProgressBar
    foreach ($currentID in $SkuCatalog_rawdata.Keys) {
        $new_record = @(
            "$currentID",
            "$($SkuCatalog_rawdata[$currentID].SKUID)",
            "$($SkuCatalog_rawdata[$currentID].DESC)"
        )
        $j = 1
        foreach ($value in $new_record) {
            $Sheet3.cells.item($i,$j) = $value
            $j++
        }
        $i++

        # progressbar
        $usrcount++
        $percent = ($usrcount / $tot)*100
        if ($percent -gt 100) {
            $percent = 100
        }
        $formattato = '{0:0.0}' -f $percent
        [int32]$progress = $percent   
        $parsebar[2].Text = ("Record {0} out of {1} written [{2}%]" -f ($usrcount, $tot, $formattato))
        if ($progress -ge 100) {
            $parsebar[1].Value = 100
        } else {
            $parsebar[1].Value = $progress
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    $parsebar[0].Close()
    $i--
    $Myworkbook.Activesheet.Cells.EntireColumn.Autofit() | Out-Null
    $Table3 = $Sheet3.ListObjects.Add(
    [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
    $Sheet3.Range("A1:C$i"), "$label",
    [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
    )
    $Table3.name = "SkuCatalog"
    Write-Host -ForegroundColor Green "Ok"
}

# writing Orphaned worksheet
if ($newOrphans -eq $true) {
    Write-Host -NoNewline "Writing worksheet [Orphaned]..."
    $Sheet4 = $Myworkbook.Worksheets.add()
    $Sheet4.name = "Orphaned"
    $i = 1
    foreach ($item in ('USRNAME','DESC','USRTYPE','CREATED', 'BLOCKED', 'LICENSED', 'LICENSE', 'TIMESTAMP', 'NOTES')) {
        $Sheet4.cells.item(1,$i) = $item
        $i++        
    }
    $i = 2
    foreach ($new_record in $orphanedrecords) {
        $j = 1
        foreach ($value in $new_record) {
            $Sheet4.cells.item($i,$j) = $value
            $j++
        }
        $i++
    }
    $i--
    $Myworkbook.Activesheet.Cells.EntireColumn.Autofit() | Out-Null
    $Table4 = $Sheet4.ListObjects.Add(
    [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
    $Sheet4.Range("A1:I$i"), "Orphaned",
    [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
    )
    $Table4.name = "Orphaned"
    Write-Host -ForegroundColor Green "Ok"
}

# writing Licenses_Pool worksheet
Write-Host -NoNewline "Writing worksheet [Licenses_Pool]..."
$Sheet1 = $Myworkbook.Worksheets.add()
$Sheet1.name = "Licenses_Pool"
$i = 1
foreach ($item in ('UPTIME','LICENSE','AVAILABLE','TOTAL')) {
    $Sheet1.cells.item(1,$i) = $item
    $i++        
}
$i = 2
foreach ($new_record in $Licenses_Pool_dataframe) {
    $j = 1
    foreach ($value in $new_record) {
        $Sheet1.cells.item($i,$j) = $value
        $j++
    }
    $i++
}
$i--
$Myworkbook.Activesheet.Cells.EntireColumn.Autofit() | Out-Null
$Table1 = $Sheet1.ListObjects.Add(
[Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
$Sheet1.Range("A1:D$i"), "Licenses_Pool",
[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
)
$Table1.name = "Licenses_Pool"
Write-Host -ForegroundColor Green "Ok"

# writing Assigned_Licenses worksheet
Write-Host -NoNewline "Writing worksheet [Assigned_Licenses]..."
$Sheet2 = $Myworkbook.Worksheets.add()
$Sheet2.name = "Assigned_Licenses"
$i = 1
foreach ($item in ('USRNAME','DESC','USRTYPE','CREATED', 'BLOCKED', 'LICENSED', 'LICENSE', 'TIMESTAMP')) {
    $Sheet2.cells.item(1,$i) = $item
    $i++        
}
$i = 2
$tot = $Assigned_Licenses_dataframe.Count
$usrcount = 0
$parsebar = ProgressBar
foreach ($new_record in $Assigned_Licenses_dataframe) {
    $j = 1
    foreach ($value in $new_record) {
        $Sheet2.cells.item($i,$j) = $value
        $j++
    }
    $i++

    # progressbar
    $usrcount++
    $percent = ($usrcount / $tot)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Record {0} out of {1} written [{2}%]" -f ($usrcount, $tot, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
$parsebar[0].Close()
$i--
$Myworkbook.Activesheet.Cells.EntireColumn.Autofit() | Out-Null
$Table2 = $Sheet2.ListObjects.Add(
[Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
$Sheet2.Range("A1:H$i"), "Assigned_Licenses",
[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
)
$Table2.name = "Assigned_Licenses"
Write-Host -ForegroundColor Green "Ok"

$Myworkbook.Saveas($xlsx_file)
$Myworkbook.Close($true)
$Myexcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Myexcel) | Out-Null

# Close Excel reference file
$XlsxObj | Close-Excel -Save
