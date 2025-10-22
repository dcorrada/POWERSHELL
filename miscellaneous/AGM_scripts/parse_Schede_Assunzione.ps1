<#
Questo script recupera dati dalle schede assunzione e li tabula in un file csv
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\parse_Schede_Assunzione\.ps1$" > $null
$workdir = $matches[1]

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

$apath = 'C:\Users\' + $env:USERNAME + '\Desktop'

# seleziono la cartella di rete (share) dove sono presenti i dati
$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
$OpenFileDialog.InitialDirectory = $apath
$OpenFileDialog.AddExtension = $false
$OpenFileDialog.CheckFileExists = $false
$OpenFileDialog.DereferenceLinks = $true
$OpenFileDialog.Filter = "Folders|`n"
$OpenFileDialog.Multiselect = $false
$OpenFileDialog.Title = "Select folder"
$OpenFileDialogType = $OpenFileDialog.GetType()
$FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
$IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null)
$OpenFileDialogType.GetMethod('OnBeforeVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$IFileDialog)
[uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
$FolderOptions = $OpenFileDialogType.GetMethod('get_Options',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null) -bor $PickFoldersOption
$FileDialogInterfaceType.GetMethod('SetOptions',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$FolderOptions)
$VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName,'System.Windows.Forms.FileDialog+VistaDialogEvents',$false,0,$null,$OpenFileDialog,$null,$null).Unwrap()
[uint32]$AdviceCookie = 0
$AdvisoryParameters = @($VistaDialogEvent,$AdviceCookie)
$AdviseResult = $FileDialogInterfaceType.GetMethod('Advise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdvisoryParameters)
$AdviceCookie = $AdvisoryParameters[1]
$Result = $FileDialogInterfaceType.GetMethod('Show',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,[System.IntPtr]::Zero)
$FileDialogInterfaceType.GetMethod('Unadvise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdviceCookie)
if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
    $FileDialogInterfaceType.GetMethod('GetResult',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$null)
}
$sharepath = $OpenFileDialog.FileName

# recupero la lista delle schede assunzione
$filelist = Get-ChildItem -Path ($sharepath + '\*.xlsx') -Name

# importo le schede
$ErrorActionPreference= 'Stop'
try {
    Import-Module ImportExcel
} catch {
    Install-Module ImportExcel -Confirm:$False -Force
    Import-Module ImportExcel
}
$ErrorActionPreference= 'Inquire'

Clear-Host
Write-Host -ForegroundColor Yellow '*** ANALISI SCHEDE ASSUNZIONI ***'
Start-Sleep 2

$parseddata = @{}
foreach ($item in $filelist) {
    $afile = $sharepath + '\' + $item
    Write-Host -NoNewline "Processing [$afile]... "
    $rawdata = Import-Excel -Path $afile -NoHeader
    foreach ($slot in $rawdata) {
        if ($slot.P1 -eq 'NOME') {
            $nome = $slot.P2
        } elseif ($slot.P1 -eq 'COGNOME') {
            $cognome = $slot.P2
        } elseif ($slot.P1 -eq 'EMAIL AZIENDALE') {
            $email = $slot.P2
        } elseif ($slot.P1 -eq 'AREA DI LAVORO') {
            $ruolo = $slot.P2
        } elseif ($slot.P1 -eq 'UTENTE INTERNO / ESTERNO') {
            $interno = $slot.P2
        } elseif ($slot.P1 -eq 'SEDE AGM DI RIFERIMENTO') {
            $sede = $slot.P2
        }
    }
    if ($email -eq $null) {
        Write-Host -ForegroundColor Red 'FAILED'
        Pause
    } else {
        $parseddata[$email] = @{
            'nome' = $nome
            'cognome' = $cognome
            'email' = $email
            'ruolo' = $ruolo
            'status' = $interno
            'licenza' = ''
            'pluslicenza' = ''
            'start' = ''
            'sede' = $sede
        }
        Write-Host -ForegroundColor Green "DONE"
    }
}

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
Connect-MsolService

Clear-Host
Write-Host -ForegroundColor Yellow '*** RISCONTRO SUL TENANT ***'
Start-Sleep 2

# retrieve all users that are licensed
$Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" } | Sort-Object DisplayName

foreach ($User in $Users) {
    $username = $User.UserPrincipalName
    $fullname = $User.DisplayName

    Write-Host -NoNewline "Looking for $username... "

    $licenses = (Get-MsolUser -UserPrincipalName $username).Licenses.AccountSku | Sort-Object SkuPartNumber
    if ($licenses.Count -ge 1) { # at least one license
        foreach ($license in $licenses) {
            $license = $license.SkuPartNumber

            if (!($parseddata.ContainsKey($username))) {
                ($nome,$cognome) = $fullname.Split(' ')
                $parseddata[$username] = @{
                    'nome' = $nome
                    'cognome' = $cognome
                    'email' = $username
                    'ruolo' = 'undef'
                    'status' = 'undef'
                    'licenza' = ''
                    'pluslicenza' = ''
                    'start' = ''
                    'sede' = 'undef'
                }
            }
            
            if ($license -match "O365_BUSINESS_PREMIUM") {
                $parseddata[$username].licenza += "*Standard"
            } elseif ($license -match "O365_BUSINESS_ESSENTIALS") {
                $parseddata[$username].licenza += "*Basic"
            } else {
                $parseddata[$username].pluslicenza += "*$license"
            }
        }
    }

    Write-Host -ForegroundColor Green 'DONE'
}

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
Connect-AzureAD

Clear-Host
Write-Host -ForegroundColor Yellow '*** RICERCA CREAZIONE ACCOUNT ***'
Start-Sleep 2

foreach ($User in $Users) {
    $username = $User.UserPrincipalName
    $plans = (Get-AzureADUser -SearchString $username).AssignedPlans

    Write-Host -NoNewline "Looking account creation for $username... "

    foreach ($record in $plans) {
        if (($record.Service -eq 'MicrosoftOffice') -and ($record.CapabilityStatus -eq 'Enabled')){
            $started = $record.AssignedTimestamp | Get-Date -format "yyyy/MM/dd"
            $parseddata[$username].start = $started
        }
    }
    
    Write-Host -ForegroundColor Green 'DONE'
}




# controllo cessazioni
$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
$OpenFileDialog.InitialDirectory = $apath
$OpenFileDialog.AddExtension = $false
$OpenFileDialog.CheckFileExists = $false
$OpenFileDialog.DereferenceLinks = $true
$OpenFileDialog.Filter = "Folders|`n"
$OpenFileDialog.Multiselect = $false
$OpenFileDialog.Title = "Select folder"
$OpenFileDialogType = $OpenFileDialog.GetType()
$FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
$IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null)
$OpenFileDialogType.GetMethod('OnBeforeVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$IFileDialog)
[uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
$FolderOptions = $OpenFileDialogType.GetMethod('get_Options',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null) -bor $PickFoldersOption
$FileDialogInterfaceType.GetMethod('SetOptions',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$FolderOptions)
$VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName,'System.Windows.Forms.FileDialog+VistaDialogEvents',$false,0,$null,$OpenFileDialog,$null,$null).Unwrap()
[uint32]$AdviceCookie = 0
$AdvisoryParameters = @($VistaDialogEvent,$AdviceCookie)
$AdviseResult = $FileDialogInterfaceType.GetMethod('Advise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdvisoryParameters)
$AdviceCookie = $AdvisoryParameters[1]
$Result = $FileDialogInterfaceType.GetMethod('Show',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,[System.IntPtr]::Zero)
$FileDialogInterfaceType.GetMethod('Unadvise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdviceCookie)
if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
    $FileDialogInterfaceType.GetMethod('GetResult',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$null)
}
$sharepath = $OpenFileDialog.FileName

# recupero la lista delle schede assunzione
$filelist = Get-ChildItem -Path ($sharepath + '\*.xlsx') -Name

Clear-Host
Write-Host -ForegroundColor Yellow '*** FILTRO SU CESSAZIONI ***'
Start-Sleep 2

foreach ($item in $filelist) {
    $afile = $sharepath + '\' + $item
    Write-Host -NoNewline "Processing [$afile]... "
    $rawdata = Import-Excel -Path $afile -NoHeader
    foreach ($slot in $rawdata) {
        if ($slot.P1 -eq 'EMAIL AZIENDALE') {
            $email = $slot.P2
        }
    }
    if ($email -eq $null) {
        Write-Host -ForegroundColor Red 'FAILED'
        Pause
    } elseif ($parseddata.ContainsKey($email)) {
        $parseddata[$email].status = 'CESSATO'
        Write-Host -ForegroundColor Green 'UPDATED'
    } else {
        Write-Host -ForegroundColor Yellow 'SKIPPED'
    }
}

# output dataframe to a CSV file
$outfile = "C:\Users\$env:USERNAME\Desktop\Licenses.csv"
Write-Host -NoNewline "Writing to $outfile... "

'NOME;COGNOME;EMAIL;SEDE;STATUS;RUOLO;DATA;LICENZA;PLUS' | Out-File $outfile -Encoding ASCII -Append

foreach ($item in $parseddata.Keys) {
    $new_record = @(
        $parseddata[$item].nome,
        $parseddata[$item].cognome,
        $parseddata[$item].email,
        $parseddata[$item].sede,
        $parseddata[$item].status,
        $parseddata[$item].ruolo,
        $parseddata[$item].start,
        $parseddata[$item].licenza,
        $parseddata[$item].pluslicenza
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}

Write-Host -ForegroundColor Green "DONE"
Pause


