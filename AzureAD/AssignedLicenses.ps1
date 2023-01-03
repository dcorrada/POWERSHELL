<#
Name......: AssignedLicenses.ps1
Version...: 22.10.2
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

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# function for killing Outlook instances
function OutlookKiller {
    $ErrorActionPreference= 'SilentlyContinue'
    $outproc = Get-Process outlook
    if ($outproc -ne $null) {
        $ErrorActionPreference= 'Stop'
        Try {
            Stop-Process -ID $outproc.Id -Force
            Start-Sleep 2
        }
        Catch { 
            [System.Windows.MessageBox]::Show("Check out that all Oulook processes have been closed before go ahead",'TASK MANAGER','Ok','Warning') > $null
        }
    }
    $ErrorActionPreference= 'Inquire'
}

# the db files
$dbfile = "C:\Users\$env:USERNAME\AppData\Local\PatrolDB.csv.AES"
$dbfile_unlocked = "C:\Users\$env:USERNAME\AppData\Local\PatrolDB.csv"

if (Test-Path $dbfile -PathType Leaf) {
    # reading current key
    $adialog = FormBase -w 350 -h 200 -text "UNLOCK DB"
    RadioButton -form $adialog -checked $true -x 20 -y 20 -w 500 -h 30 -text "Enter the key for accessing to DB file" | Out-Null
    $currentkey = TxtBox -form $adialog -x 20 -y 50 -w 300 -h 30 -text ''
    $cleanDB = RadioButton -form $adialog -checked $false -x 20 -y 80 -w 500 -h 30 -text "Clean existing DB file"
    OKButton -form $adialog -x 100 -y 120 -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
    if ($cleanDB.Checked -eq $true) {
        $answ = [System.Windows.MessageBox]::Show("Really delete DB file?",'DELETE','YesNo','Warning')
        if ($answ -eq "Yes") {    
            Remove-Item -Path $dbfile
        }
    }
}

Import-Module -Name "$workdir\Modules\FileCryptography.psm1"
if (Test-Path $dbfile -PathType Leaf) {
    # unlocking DB file
    $ErrorActionPreference= 'Stop'
    Try {
        $chiave = ConvertTo-SecureString $currentkey.Text -AsPlainText -Force
        Unprotect-File $dbfile -Algorithm AES -Key $chiave -RemoveSource | Out-Null
        Write-Host -ForegroundColor Green "*** ACCESS GRANTED ***"
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Host -ForegroundColor Red "*** AUTHENTICATION ERROR, aborting ***"
        # Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }
} else {
    # creating DB file
    $adialog = FormBase -w 400 -h 300 -text "DB INIT"
    Label -form $adialog -x 20 -y 20 -w 500 -h 30 -text "Initialize your DB as follows" | Out-Null
    $dbcontent = TxtBox -form $adialog -x 20 -y 50 -w 300 -h 150 -text ''
    $dbcontent.Multiline = $true;
    $dbcontent.Text = @'
user1@foobar.baz;password1
user2@foobar.baz;password2
'@
    $dbcontent.AcceptsReturn = $true
    OKButton -form $adialog -x 100 -y 220 -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
     'USR;PWD' | Out-File $dbfile_unlocked -Encoding ASCII -Append
    $dbcontent.Text | Out-File $dbfile_unlocked -Encoding ASCII -Append
}

# reading DB file
$filecontent = Get-Content -Path $dbfile_unlocked
$allowed = @{}
foreach ($newline in $filecontent) {
    if ($newline -ne 'USR;PWD') {
        ($username, $passwd) = $newline.Split(';')
        $allowed[$username] = $passwd
    }
}

# locking DB file
$newkey = New-CryptographyKey -Algorithm AES -AsPlainText
$securekey = ConvertTo-SecureString $newkey -AsPlainText -Force
$ErrorActionPreference= 'Stop'
Try {
    Protect-File $dbfile_unlocked -Algorithm AES -Key $securekey | Out-Null
    Remove-Item -Path $dbfile_unlocked
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Host -ForegroundColor Red "*** ERROR LOCKING ***"
    Write-Host -ForegroundColor Yellow "[$dbfile_unlocked] was not crypted"
    # Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

# show crypto key
$adialog = FormBase -w 400 -h 250 -text "LOCK DB"
Label -form $adialog -x 20 -y 20 -w 500 -h 30 -text "The key for accessing to DB file will be" | Out-Null
TxtBox -form $adialog -x 20 -y 50 -w 300 -h 30 -text "$newkey" | Out-Null
Label -form $adialog -x 20 -y 80 -w 500 -h 30 -text "Cut 'n' Paste such string somewhere, otherwise..." | Out-Null
$sendme = CheckBox -form $adialog -x 20 -y 110 -checked $false -text "send me an email (Outook)"
OKButton -form $adialog -x 100 -y 160 -text "Ok" | Out-Null
$result = $adialog.ShowDialog()

# send crypto key
if ($sendme.Checked -eq $true) {
    Write-Host -NoNewline "Sending crypto key..."
    [System.Windows.MessageBox]::Show("Click Ok to close Outlook",'CLOSE','Ok','Warning') | Out-Null
    OutlookKiller
    $ErrorActionPreference= 'Stop'
    Try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNameSpace("MAPI")
        $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
        $InboxDef = $namespace.GetDefaultFolder($olFolders::olFolderInBox)
        $InboxDef.FullFolderPath -match "^\\\\(.*@.*)\\(Inbox|Posta)" > $null
        $recipient = $matches[1]
        $email = $outlook.CreateItem(0)
        $email.To = "$recipient"
        $email.Subject = "Your Crypto Key"
        $email.Body = "$newkey"
        $email.Send()
        $Outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        Start-Sleep 3
        OutlookKiller
        Write-Host -ForegroundColor Green " DONE"
    }
    Catch {
        Write-Host -ForegroundColor Red " FAILED"
        # Write-Output "`nError: $($error[0].ToString())"
        $adialog = FormBase -w 400 -h 170 -text "WARNING"
        Label -form $adialog -x 20 -y 20 -w 500 -h 30 -text "Your crypto key was not sent" | Out-Null
        TxtBox -form $adialog -x 20 -y 50 -w 300 -h 30 -text "$newkey" | Out-Null
        OKButton -form $adialog -x 100 -y 80 -text "Ok" | Out-Null
        $result = $adialog.ShowDialog()
    }
    $ErrorActionPreference= 'Inquire'
    Start-Process outlook
} else {
    Write-Host -ForegroundColor Blue "Send crypto key by email disabled"
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
    $choices += RadioButton -form $adialog -x 20 -y $they -checked $isfirst -text $username
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
