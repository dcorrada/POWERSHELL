<#
Name......: AssignedLicenses.ps1
Version...: 23.05.1
Author....: Dario CORRADA

This script will connect to Azure AD and query a list of which license(s) are assigned to each user

For more details about AzureAD cmdlets see:
https://docs.microsoft.com/en-us/powershell/module/azuread
#>

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
Import-Module -Name "$workdir\Modules\Forms.psm1"

# the db files
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

Import-Module -Name "$workdir\Modules\Gordian.psm1"
if (!(Test-Path $dbfile -PathType Leaf)) {
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

# get credentials for accessing
foreach ($item in $choices) {
    if ($item.Checked) {
        $usr = $item.Text
        $plain_pwd = $allowed[$usr]
    }
}
$pwd = ConvertTo-SecureString $plain_pwd -AsPlainText -Force
$credits = New-Object System.Management.Automation.PSCredential($usr, $pwd)

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module MSOnline
} catch {
    Install-Module MSOnline -Confirm:$False -Force
    Import-Module MSOnline
}
$ErrorActionPreference= 'Inquire'

# connect to Tenant
$ErrorActionPreference= 'Stop'
Try {
    Connect-MsolService -Credential $credits
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Host -ForegroundColor Red "*** ERROR ACCESSING TENANT ***"
    # Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

# available licenses
$avail_licenses = @{}
$get_licenses = Get-MsolAccountSku
foreach ($license in $get_licenses) {
    $label = $license.AccountSkuId
    $avail = $license.ActiveUnits - $license.ConsumedUnits
    $avail_licenses[$label] = $avail
}
$adialog = FormBase -w 400 -h ((($avail_licenses.Count) * 30) + 120) -text "AVAILABLE LICENSES"
$they = 20
foreach ($item in $avail_licenses.GetEnumerator() | Sort Value) {
    $string = $item.Name + " = " + $item.Value
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,$they)
    $label.Size = New-Object System.Drawing.Size(350,20)
    $label.Text = $string
    $adialog.Controls.Add($label)
    $they += 30
}
OKButton -form $adialog -x 75 -y ($they + 10) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

# retrieve all users that are licensed
$Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" } | Sort-Object DisplayName

# initialize dataframe for collecting data
$parseddata = @{}

$tot = $Users.Count
$usrcount = 0
$parsebar = ProgressBar
Clear-Host
Write-Host -NoNewline "STEP01 - Collecting..."
foreach ($User in $Users) {
    $usrcount ++

    $username = $User.UserPrincipalName
    $fullname = $User.DisplayName

    # for the AccountSku list see https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference 
    $licenses = (Get-MsolUser -UserPrincipalName $username).Licenses.AccountSku | Sort-Object SkuPartNumber
    if ($licenses.Count -ge 1) { # at least one license
        foreach ($license in $licenses) {
            $license = $license.SkuPartNumber
            $splitted = $fullname.Split(' ')
            $parseddata[$username] = @{
                'nome' = $splitted[0]
                'cognome' = $splitted[1]
                'email' = $username
                'licenza' = ''
                'pluslicenza' = ''
                'start' = ''
            }

            if ($license -match "O365_BUSINESS_PREMIUM") {
                $parseddata[$username].licenza += "*Standard"
            } elseif ($license -match "O365_BUSINESS_ESSENTIALS") {
                $parseddata[$username].licenza += "*Basic"
            } elseif ($license -match "EXCHANGESTANDARD") {
                $parseddata[$username].licenza += "*Exchange"   
            } elseif ($license -match "ENTERPRISEPACKPLUS_FACULTY") {
                $parseddata[$username].licenza += "*A3_EnterprisePackPlus"
            } elseif ($license -match "M365EDU_A3_FACULTY") {
                $parseddata[$username].licenza += "*A3_EDU"
            } elseif ($license -match "STANDARDWOFFPACK_FACULTY") {
                $parseddata[$username].licenza += "*A1"
            } else {
                $parseddata[$username].pluslicenza += "*$license"
            }
        }
    }

    # progress
    $percent = ($usrcount / $tot)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($usrcount, $tot, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
Write-Host -ForegroundColor Green " DONE"
$parsebar[0].Close()

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module AzureAD
} catch {
    Install-Module AzureAD -Confirm:$False -Force
    Import-Module AzureAD
}
$ErrorActionPreference= 'Inquire'

# connect to AzureAD
$ErrorActionPreference= 'Stop'
Try {
    Connect-AzureAD -Credential $credits | Out-Null
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Host -ForegroundColor Red "*** ERROR ACCESSING TENANT ***"
    # Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

Start-Sleep 2
$tot = $Users.Count
$usrcount = 0
$parsebar = ProgressBar
Clear-Host
Write-Host -NoNewline "STEP02 - Finalizing..."
foreach ($User in $Users) {
    $usrcount ++

    $username = $User.UserPrincipalName
    $plans = (Get-AzureADUser -SearchString $username).AssignedPlans

    foreach ($record in $plans) {
        if ((($record.Service -eq 'MicrosoftOffice') -or ($record.Service -eq 'exchange')) -and ($record.CapabilityStatus -eq 'Enabled')){
            $started = $record.AssignedTimestamp | Get-Date -format "yyyy/MM/dd"
            if ($parseddata[$username].start -eq '') {
                $parseddata[$username].start = $started
            } elseif ($started -lt $parseddata[$username].start) {
                $parseddata[$username].start = $started
            }
        }
    }
    
    # progress
    $percent = ($usrcount / $tot)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($usrcount, $tot, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
Write-Host -ForegroundColor Green " DONE"
$parsebar[0].Close()

# writing output file
# see https://techexpert.tips/powershell/powershell-creating-excel-file/
Clear-Host
Write-Host -NoNewline "Writing output file... "
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.Title = "Save File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
$OpenFileDialog.filename = 'licenses'
$OpenFileDialog.ShowDialog() | Out-Null
$outfile = $OpenFileDialog.filename
$Myexcel = New-Object -ComObject excel.application
$Myexcel.visible = $false
$Myworkbook = $Myexcel.workbooks.add()
$Sheet1 = $Myworkbook.worksheets.item(1)
$Sheet1.name = "Assigned_Licenses"
$i = 1
foreach ($item in ('NAME','SURNAME','EMAIL','DATE','LICENSE','PLUS')) {
    $Sheet1.cells.item(1,$i) = $item
    $i++
}
$Sheet1.Range("A1:F1").font.size = 12
$Sheet1.Range("A1:F1").font.bold = $true
$Sheet1.Range("A1:F1").font.ColorIndex = 2
$Sheet1.Range("A1:F1").interior.colorindex = 1
$i = 2
$totrec = $parseddata.Keys.Count
$parsebar = ProgressBar
foreach ($item in $parseddata.Keys) {
    $new_record = @(
        $parseddata[$item].nome,
        $parseddata[$item].cognome,
        $parseddata[$item].email,
        $parseddata[$item].start,
        $parseddata[$item].licenza,
        $parseddata[$item].pluslicenza
    )
    $j = 1
    foreach ($value in $new_record) {
        $Sheet1.cells.item($i,$j) = $value
        $j++
    }
    $i++

    # progress
    $percent = (($i-1) / $totrec)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Writing {0} out of {1} records [{2}%]" -f (($i-1), $totrec, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()    
}
$parsebar[0].Close()
$Myworkbook.Activesheet.Cells.EntireColumn.Autofit()
$Myexcel.displayalerts = $false
$Myworkbook.Saveas($outfile)
$Myexcel.displayalerts = $true
$Myexcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Myexcel) | Out-Null
Write-Host -ForegroundColor Green "DONE"
Pause
