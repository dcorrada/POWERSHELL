Param([string]$ExtScript='PSWallet', [string]$ExtKey='NULL',
    [string]$ExtUsr='NULL', [string]$ExtPwd='NULL',
    [string]$ExtAction='NULL')

<#
Name......: PSWallet.ps1
Version...: 24.06.1
Author....: Dario CORRADA

PSWallet aims to be the credential manager tool in order to handle the various 
login attempts required alongside the scripts of this git repository. 
It will store, fetch and update credential onto a SQLite database.

Refs:
* https://github.com/RamblingCookieMonster/PSSQLite
* https://www.powershellgallery.com/packages/PSSQLite
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Safety\\PSWallet\.ps1$" > $null
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
        Import-Module -Name "$workdir\Modules\Gordian.psm1"
        Import-Module PSSQLite
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'PSSQLite')) {
            Install-Module PSSQLite -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [PSSQLite] module: please restart the script",'RESTART','Ok','warning') | Out-Null
            $ThirdParty = 'Ko'
        } elseif (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
            Install-Module ImportExcel -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: please restart the script",'RESTART','Ok','warning') | Out-Null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') | Out-Null
            Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                INITIALIZATION
******************************************************************************* #>
$dbfile = $env:LOCALAPPDATA + '\PSWallet.sqlite'

do {
    if (!(Test-Path $dbfile -PathType Leaf)) {
        # take a view to an Excel demo to import
        $demoxlsx = "$workdir\Safety\PSWallet_DemoImport.xlsx"
        $aansw = [System.Windows.MessageBox]::Show(@"
No database found yet. Would you like 
to view a typical Excel file to import?
"@,'DEMO','YesNo','Warning')
        if ($aansw -eq 'Yes') {
            Invoke-Item $demoxlsx
        }

        # create Key file if no DB file exists
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Title = "Create Key File"
        $SaveFileDialog.initialDirectory = "$env:LOCALAPPDATA"
        $SaveFileDialog.FileName = 'PSWallet.key'
        $SaveFileDialog.filter = 'Key file (*.key)| *.key'
        $SaveFileDialog.ShowDialog() | Out-Null
        $keyfile = $SaveFileDialog.filename
        CreateKeyFile -keyfile "$keyfile" | Out-Null
    } elseif ($ExtKey -eq 'NULL') {
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Title = "Open Key File"
        $OpenFileDialog.initialDirectory = "$env:LOCALAPPDATA"
        $OpenFileDialog.filter = 'Key file (*.key)| *.key'
        $OpenFileDialog.ShowDialog() | Out-Null
        $keyfile = $OpenFileDialog.filename   
    } else {
        $keyfile = $ExtKey
    }
} while ([string]::IsNullOrEmpty($keyfile))


Write-Host -NoNewline 'Accessing DB file... '
$ErrorActionPreference= 'Stop'
try {
    if (Test-Path $dbfile -PathType Leaf) {
        $SQLiteConnection = New-SQLiteConnection -DataSource $dbfile
    } else {
        Write-Host -NoNewline -ForegroundColor Yellow 'No DB found, create it '
        $SQLiteConnection = New-SQLiteConnection -DataSource $dbfile

        # creating DB schema
        Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @'
CREATE TABLE `Logs` (
    `USER` varchar(80),
    `HOST` varchar(80),
    `ACTION` varchar(80),
    `UPTIME` datetime,
    `SCRIPT` text
);
CREATE TABLE `Credits` (
    `USER` varchar(80),
    `PSWD` text,
    `SCRIPT` text
);
CREATE TABLE `Graph` (
    `APPNAME` varchar(80),
    `APPID` varchar(80),
    `OBJID` varchar(80),
    `TENANTID` varchar(80),
    `SECRETNAME` varchar(80),
    `SECRETID` varchar(80),
    `SECRETVALUE` text,
    `SECRETEXPDATE` datetime,
    `UPN` text,
    `SCRIPT` text
);
'@
    }

    $SQLiteConnection.Close()
    Write-Host -ForegroundColor Green 'Ok'
} catch {
    Write-Host -ForegroundColor Red 'Ko'
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Error accessing database",'ABORTING','Ok','Error')
    exit
}
$ErrorActionPreference= 'Inquire'

# Logs clean
$BackInTime = '-90 day'
$SQLiteConnection.Open()
$howmuch = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query "SELECT * FROM Logs WHERE UPTIME > DATETIME('now', '-1 day') AND ACTION LIKE 'history clean%';"
if ($howmuch.Count -eq 0) { # just perform single check a day
    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query "DELETE FROM Logs WHERE UPTIME < DATETIME('now', '$BackInTime');"
    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'history clean $BackInTime',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
}
$SQLiteConnection.Close()


<# *******************************************************************************
                                LOCALES
******************************************************************************* #>
if (($ExtScript -eq 'PSWallet') -or  (!(Test-Path $dbfile -PathType Leaf))) {
    do {
        $adialog = FormBase -w 350 -h 260 -text "MAINTENANCE"
        $importCredits = RadioButton -form $adialog -checked $true -x 20 -y 20 -w 500 -h 30 -text "Import Database from xlsx"
        $exportCredits = RadioButton -form $adialog -checked $false -x 20 -y 50 -w 500 -h 30 -text "Export Database to xlsx"
        $editCredits = RadioButton -form $adialog -checked $false -x 20 -y 80 -w 500 -h 30 -text "Edit single credential"
        $deleteCredits = RadioButton -form $adialog -checked $false -x 20 -y 110 -w 500 -h 30 -text "Delete single credential"
        RETRYButton -form $adialog -x 40 -y 170 -text "Next" | Out-Null
        OKButton -form $adialog -x 180 -y 170 -text "Exit" | Out-Null
        $resultButton = $adialog.ShowDialog()

        if ($resultButton -eq 'RETRY') {
            if ($importCredits.Checked -eq $true) {
                [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
                $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $OpenFileDialog.Title = "Import Table"
                $OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
                $OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
                $OpenFileDialog.ShowDialog() | Out-Null
                $ImportFile = $OpenFileDialog.filename 
                
                $headers = Compare-Object -ReferenceObject ('USER', 'PSWD', 'SCRIPT') -DifferenceObject ((Import-Excel -Path $ImportFile | Get-Member).Name)
                if ($headers.SideIndicator -contains '<=') {
                    [System.Windows.MessageBox]::Show("Required fields doesn't match",'ERROR','Ok','Error')
                } else {
                    Write-Host -NoNewline 'Importing...'
                    $ImportedTable = Import-Excel -Path $ImportFile | ForEach-Object {
                        Write-Host -NoNewline '.'
                        New-Object -TypeName PSObject -Property @{
                            USER    = $_.USER
                            PSWD    = EncryptString -keyfile $keyfile -instring $_.PSWD
                            SCRIPT  = $_.SCRIPT
                        } | Select USER, PSWD, SCRIPT
                    }
                    $DataTable = $ImportedTable | Out-DataTable
                    # purge any existing record
                    $SQLiteConnection.Open()
                    Write-Host -NoNewline '.'
                    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query 'DELETE FROM `Credits`;'
                    # insert new records
                    $stdout = Invoke-SQLiteBulkCopy -DataTable $DataTable -SQLiteConnection $SQLiteConnection -Table Credits -NotifyAfter 0 -ConflictClause Ignore -Force
                    Write-Host -NoNewline '.'
                    Write-Host -ForegroundColor Green ' DONE'
                    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'import table',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
                    $SQLiteConnection.Close()
                }
            } elseif ($exportCredits.Checked -eq $true) {
                Write-Host -NoNewline 'Exporting...'
                $SQLiteConnection.Open()
                $rawdata = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query 'SELECT * FROM `Credits`;'
                Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'export table',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
                $SQLiteConnection.Close()

                $ExportedTable = $rawdata | ForEach-Object {
                    Write-Host -NoNewline '.'
                    New-Object -TypeName PSObject -Property @{
                        USER    = $_.USER
                        PSWD    = DecryptString -keyfile $keyfile -instring $_.PSWD
                        SCRIPT  = $_.SCRIPT
                    } | Select USER, PSWD, SCRIPT
                }
                [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
                $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                $SaveFileDialog.Title = "Export Credits"
                $SaveFileDialog.initialDirectory = "C:$env:HOMEPATH\Downloads"
                $SaveFileDialog.FileName = 'PSWallet.xlsx'
                $SaveFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
                $SaveFileDialog.ShowDialog() | Out-Null
                $ExportFile = $SaveFileDialog.filename

                Write-Host -NoNewline '.'
                $XlsPkg = Open-ExcelPackage -Path $ExportFile -Create
                $XlsPkg = $ExportedTable | Export-Excel -ExcelPackage $XlsPkg -WorksheetName 'Credits' -TableName 'Credits' -TableStyle 'Medium3' -AutoSize -PassThru
                Close-ExcelPackage -ExcelPackage $XlsPkg
                Write-Host -ForegroundColor Green ' DONE'
            } elseif ($editCredits.Checked -eq $true) {
                Write-Host -NoNewline 'Editing...'
                $SQLiteConnection.Open()
                $rawdata = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query 'SELECT * FROM `Credits`;'
                $SQLiteConnection.Close()
                
                $formlist = FormBase -w 250 -h 175 -text 'SELECT'
                Label -form $formlist -x 10 -y 20 -w 40 -text 'Script:' | Out-Null
                $selectedScript = DropDown -form $formlist -x 60 -y 20 -w 120 -opts ($rawdata.SCRIPT | select -Unique | sort)
                Label -form $formlist -x 10 -y 50 -w 40 -text 'User:' | Out-Null
                $selectedUser = DropDown -form $formlist -x 60 -y 50 -w 120 -opts ($rawdata.USER | select -Unique | sort)
                OKButton -form $formlist -x 60 -y 90 -text "Ok" | Out-Null
                $result = $formlist.ShowDialog()
                
                do {
                    $SQLiteConnection.Open()
                    $rowexists = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query "SELECT * FROM Credits WHERE USER = '$($selectedUser.Text)' AND SCRIPT = '$($selectedScript.Text)';"
                    $SQLiteConnection.Close()

                    if ($rowexists -eq $null) {
                        [System.Windows.MessageBox]::Show("No user <$($selectedUser.Text)> related to script <$($selectedScript.Text)>",'ABORTING','Ok','Error') | Out-Null
                        $willabort = 'noabort'
                    } else {
                        $updateform = FormBase -w 300 -h 175 -text 'UPDATE'
                        Label -form $updateform -x 10 -y 20 -w 100 -text 'New password:' | Out-Null
                        $newpwd = TxtBox -form $updateform -x 120 -y 20 -w 150 -masked $true
                        Label -form $updateform -x 10 -y 50 -w 100 -text 'Confirm password:' | Out-Null
                        $confirmpwd = TxtBox -form $updateform -x 120 -y 50 -w 150 -masked $true
                        OKButton -form $updateform -x 60 -y 90 -text "Ok" | Out-Null
                        $result = $updateform.ShowDialog()
                        if ($newpwd.Text -cne $confirmpwd.Text) {
                            $willabort = [System.Windows.MessageBox]::Show("Password doesn't match. Aborting?",'ABORTING','YesNo','Error')
                        } else {
                            $willabort = 'noabort'
                        }
                    }
                } while ($willabort -eq 'No') 

                if ($willabort -eq 'noabort') {
                    $SQLiteConnection.Open()
                    $new_encrypted = EncryptString -keyfile $keyfile -instring $newpwd.Text
                    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query "UPDATE Credits SET PSWD = '$new_encrypted' WHERE USER = '$($selectedUser.Text)' AND SCRIPT = '$($selectedScript.Text)';"
                    $SQLiteConnection.Close()

                    $SQLiteConnection.Open()
                    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'edit row',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
                    $SQLiteConnection.Close()
                    Write-Host -ForegroundColor Green ' DONE'
                }
            } elseif ($deleteCredits.Checked -eq $true) {
                Write-Host -NoNewline 'Deleting...'
                $SQLiteConnection.Open()
                $rawdata = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query 'SELECT * FROM `Credits`;'
                $SQLiteConnection.Close()
                
                $formlist = FormBase -w 250 -h 175 -text 'SELECT'
                Label -form $formlist -x 10 -y 20 -w 40 -text 'Script:' | Out-Null
                $selectedScript = DropDown -form $formlist -x 60 -y 20 -w 120 -opts ($rawdata.SCRIPT | select -Unique | sort)
                Label -form $formlist -x 10 -y 50 -w 40 -text 'User:' | Out-Null
                $selectedUser = DropDown -form $formlist -x 60 -y 50 -w 120 -opts ($rawdata.USER | select -Unique | sort)
                OKButton -form $formlist -x 60 -y 90 -text "Ok" | Out-Null
                $result = $formlist.ShowDialog()
                
                $SQLiteConnection.Open()
                $rowexists = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query "SELECT * FROM Credits WHERE USER = '$($selectedUser.Text)' AND SCRIPT = '$($selectedScript.Text)';"
                $SQLiteConnection.Close()

                if ($rowexists -eq $null) {
                    [System.Windows.MessageBox]::Show("No user <$($selectedUser.Text)> related to script <$($selectedScript.Text)>",'ABORTING','Ok','Error') | Out-Null
                    $willabort = 'noabort'
                } else {
                    $willabort = [System.Windows.MessageBox]::Show("Really delete <$($selectedUser.Text)::$($selectedScript.Text)>?",'CONFIRM','YesNo','Warning')
                    if ($willabort -eq 'Yes') {
                        $SQLiteConnection.Open()
                        Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query "DELETE FROM Credits WHERE USER = '$($selectedUser.Text)' AND SCRIPT = '$($selectedScript.Text)';"
                        $SQLiteConnection.Close()

                        $SQLiteConnection.Open()
                        Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'delete row',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
                        $SQLiteConnection.Close()
                        Write-Host -ForegroundColor Green ' DONE'
                    }
                }
            }
        }
    } until ($resultButton -eq 'OK')
}


<# *******************************************************************************
                                EXTERNAL
******************************************************************************* #>
Write-Host -NoNewline 'Calling PSWallet from '
Write-Host -ForegroundColor Blue "$ExtScript"

if ($ExtAction -eq 'listusr') {
    $SQLiteConnection.Open()
    $TheResult = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
SELECT USER
FROM Credits
WHERE SCRIPT = '$ExtScript'
"@
    $SQLiteConnection.Close()
    if ([string]::IsNullOrEmpty($TheResult)) {
        Write-Host -ForegroundColor Yellow 'PSWallet>>> NO DATA FOUND'
    } else {
        foreach ($item in $TheResult.USER) {
            Write-Host -ForegroundColor Cyan "PSWallet>>> $item"
        }
    }
    $TheAction = 'read table'
} elseif ($ExtAction -eq 'getpwd') {
    $SQLiteConnection.Open()
    $TheResult = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
SELECT PSWD
FROM Credits
WHERE SCRIPT = '$ExtScript' AND USER = '$ExtUsr'
"@
    $SQLiteConnection.Close()
    if ([string]::IsNullOrEmpty($TheResult)) {
        Write-Host -ForegroundColor Yellow 'PSWallet>>> NO DATA FOUND'
    } else {
        foreach ($item in $TheResult.PSWD) {
            $plantxt = DecryptString -keyfile $keyfile -instring $item
            Write-Host -ForegroundColor Cyan "PSWallet>>> $plantxt"
        }
    }
    $TheAction = 'read table'
} elseif ($ExtAction -eq 'add') {
    $SQLiteConnection.Open()
    $ItExists = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
SELECT PSWD
FROM Credits
WHERE SCRIPT = '$ExtScript' AND USER = '$ExtUsr'
"@
    $SQLiteConnection.Close()

    if ([string]::IsNullOrEmpty($ItExists)) {
        $CryptedPwd = EncryptString -keyfile $keyfile -instring $ExtPwd
        $SQLiteConnection.Open()
        Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Credits (USER, PSWD, SCRIPT) 
VALUES ('$ExtUsr', '$CryptedPwd', '$ExtScript');
"@
        $SQLiteConnection.Close()
    } else {
        Write-Host -ForegroundColor Yellow 'PSWallet>>> RECORD ALREDADY EXISTS'
    }

    $TheAction = 'add row'
} elseif ($ExtAction -eq 'addGraph') {
    $toPSWallet = @{ SCRIPT = $ExtScript }
    $newentryform = FormBase -w 410 -h 440 -text 'NEW ENTRY'
    Label -form $newentryform -x 20 -y 10 -w 80 -text 'App name:' | Out-Null
    $toPSWallet.APPNAME = TxtBox -form $newentryform -x 100 -y 10 -w 275
    Label -form $newentryform -x 20 -y 60 -w 80 -text 'Client ID:' | Out-Null
    $toPSWallet.APPID = TxtBox -form $newentryform -x 100 -y 60 -w 275
    Label -form $newentryform -x 20 -y 90 -w 80 -text 'Object ID:' | Out-Null
    $toPSWallet.OBJID = TxtBox -form $newentryform -x 100 -y 90 -w 275
    Label -form $newentryform -x 20 -y 120 -w 80 -text 'Tenant ID:' | Out-Null
    $toPSWallet.TENANTID = TxtBox -form $newentryform -x 100 -y 120 -w 275
    Label -form $newentryform -x 20 -y 170 -w 80 -text 'Secret name:' | Out-Null
    $toPSWallet.SECRETNAME = TxtBox -form $newentryform -x 100 -y 170 -w 275
    Label -form $newentryform -x 20 -y 200 -w 80 -text 'Secret ID:' | Out-Null
    $toPSWallet.SECRETID = TxtBox -form $newentryform -x 100 -y 200 -w 275
    Label -form $newentryform -x 20 -y 230 -w 80 -text 'Secret value:' | Out-Null
    $toPSWallet.SECRETVALUE = TxtBox -form $newentryform -x 100 -y 230 -w 275
    Label -form $newentryform -x 20 -y 260 -w 80 -text 'Expire date:' | Out-Null
    $toPSWallet.SECRETEXPDATE = TxtBox -form $newentryform -x 100 -y 260 -w 275
    Label -form $newentryform -x 20 -y 310 -w 80 -text 'UPN:' | Out-Null
    $toPSWallet.SECRETEXPDATE = TxtBox -form $newentryform -x 100 -y 310 -w 275
    OKButton -form $newentryform -x 140 -y 350 -text "Ok" | Out-Null    
    $resultButton = $newentryform.ShowDialog()

    $SQLiteConnection.Open()
    $ItExists = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
SELECT *
FROM Graph
WHERE APPNAME = '$($toPSWallet.APPNAME.Text)'
"@
    $SQLiteConnection.Close()

    if ([string]::IsNullOrEmpty($ItExists)) {
        $ErrorActionPreference= 'Stop'
        try {
            $CryptedSecret = EncryptString -keyfile $keyfile -instring $toPSWallet.SECRETVALUE.Text
            $SQLiteConnection.Open()
            Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Graph (APPNAME, APPID, OBJID, TENANTID, SECRETNAME, SECRETID, SECRETVALUE, SECRETEXPDATE, UPN, SCRIPT) 
VALUES ('$($toPSWallet.APPNAME.Text)', 
        '$($toPSWallet.APPID.Text)',
        '$($toPSWallet.OBJID.Text)',
        '$($toPSWallet.TENANTID.Text)',
        '$($toPSWallet.SECRETNAME.Text)',
        '$($toPSWallet.SECRETID.Text)',
        '$CryptedSecret',
        '$(($toPSWallet.SECRETEXPDATE.Text | Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())',
        '$($toPSWallet.UPN.Text),
        '$ExtScript');
"@
            $SQLiteConnection.Close()
        } catch {
            Write-Host -ForegroundColor Red 'Ko'
            Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
            [System.Windows.MessageBox]::Show("Error accessing database",'ABORTING','Ok','Error')
        }
        $ErrorActionPreference= 'Inquire'

        foreach ($currentItem in ($toPSWallet.Keys | Sort-Object)) {
            Write-Host -ForegroundColor Cyan "PSWallet>>> [$currentItem] $($toPSWallet[$currentItem].Text)"
        }
    } else {
        Write-Host -ForegroundColor Yellow 'PSWallet>>> RECORD ALREDADY EXISTS'
    }

    $TheAction = 'add row'
} elseif ($ExtAction -eq 'listGraph') {
    $SQLiteConnection.Open()
    $TheResult = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
SELECT UPN, APPNAME, SECRETEXPDATE
FROM Graph
WHERE SCRIPT = '$ExtScript'
"@
    $SQLiteConnection.Close()
    if ([string]::IsNullOrEmpty($TheResult)) {
        Write-Host -ForegroundColor Yellow 'PSWallet>>> NO DATA FOUND'
    } else {
        foreach ($item in $TheResult) {
            Write-Host -ForegroundColor Cyan "PSWallet>>> $($item.UPN) <-> $($item.APPNAME) (secret will expires on $($item.SECRETEXPDATE | Get-Date -format "yyyy-MM-dd"))"
        }
    }
    $TheAction = 'read table'
}

$SQLiteConnection.Open()
Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME, SCRIPT) 
VALUES (
'$($env:USERNAME)',
'$($env:COMPUTERNAME)',
'$TheAction',
'$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())',
'$ExtScript'
);
"@
$SQLiteConnection.Close()
