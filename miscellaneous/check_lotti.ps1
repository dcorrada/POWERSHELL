<#
Questo script è stato sviluppato per un'esigenza specifica temporanea, non ha uno scopo generico.

Potrebbe essere di interesse perche':
* ci sono esempi di come fare una ricerca ricorsiva dentro una cartella
* ci sono esempi di come importare dati da un file Excel
* c'e' un esempio di dialog box per cercare cartelle di rete
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# roba grafica
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# seleziono il file Excel del lotto
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms')
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Carica Lotto"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
$OpenFileDialog.ShowDialog() | Out-Null
$lottofile = $OpenFileDialog.filename

# importo il file Excel del lotto
$ErrorActionPreference= 'Stop'
try {
    Import-Module ImportExcel
} catch {
    Install-Module ImportExcel -Confirm:$False -Force
    Import-Module ImportExcel
}
$ErrorActionPreference= 'Inquire'
$rawdata = Import-Excel $lottofile # verificare che i nomi dei campi siano univoci
# modificare le seguenti righe in base al nome dei campi
$entries_vodafone = $rawdata.'OFFICE VF'
$entries_tim = $rawdata.'DBR TI'
$entries_inwit = $rawdata.'CODICE INWIT'

# seleziono la cartella di rete (share) dove sono presenti i dati
$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
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

# accesso alla share
$login = Get-Credential
<# decommentare il blocco se ci sono problemi di montaggio
Write-Host -NoNewline 'Attendere il riavvio dei servizi di rete... '
net stop workstation /y | Out-Null
net start workstation | Out-Null
Write-Host 'OK'
#>
New-PSDrive -Name R -PSProvider FileSystem -Root $sharepath -Credential $login > $null

# percorso del file di log
$logfile = "C:\Users\$env:USERNAME\Desktop\check_lotti.log"

# check dei dati esistenti
"*** check dati esistenti ***" | Out-File $logfile -Encoding ASCII -Append
$counter = 0
$tot = $entries_inwit.Count
$searchstring = $sharepath -replace "\\", "\\"
foreach ($entry in $entries_inwit) {
    if ($entry -ne $null) {
        $found = Get-ChildItem -Path R:\ -Filter $entry -Attributes D -Recurse -ErrorAction SilentlyContinue -Force
        if ($found -eq $null) {
            $msg = "[$entry] *** NOT FOUND ***"
        } else {
            $found[0].FullName -match "$searchstring(.+)$" > $null
            $destpath = $matches[1]
            $found = Get-ChildItem -Path "R:\$destpath" -ErrorAction SilentlyContinue -Force
            if ($found -eq $null) {
                $msg = "[$entry] *** NO DATA (empty directory) ***"
            } else {
                $msg = "[$entry] " + $found[0].FullName
            }
        }
        $msg | Out-File $logfile -Encoding ASCII -Append
    }
    $counter++
    Clear-Host
    Write-Host "Check dati esistenti: $counter out of $tot"
}

# check dei dati progetto vodafone
"`n`n*** check dati progetto vodafone ***" | Out-File $logfile -Encoding ASCII -Append
$counter = 0
$tot = $entries_vodafone.Count
foreach ($entry in $entries_vodafone) {
    if ($entry -ne $null) {
        $found = Get-ChildItem -Path R:\ -Filter "$entry.pdf" -Recurse -ErrorAction SilentlyContinue -Force
        if ($found -eq $null) {
            $msg = "[$entry] *** NOT FOUND ***"
        } else {
            $msg = "[$entry] " + $found[0].FullName
        }
        $msg | Out-File $logfile -Encoding ASCII -Append
    }
    $counter++
    Clear-Host
    Write-Host "Check dati progetto Vodafone: $counter out of $tot"
}

# check dei dati progetto TIM
"`n`n*** check dati progetto TIM ***" | Out-File $logfile -Encoding ASCII -Append
$tot = $entries_tim.Count
for ($i = 0; $i -lt $entries_tim.Count; $i++) {
    $entry = $entries_tim[$i]
    if ($entry -eq $null) {
        $entry = $entries_inwit[$i]
    }
    $found = Get-ChildItem -Path R:\ -Filter "$entry.pdf" -Recurse -ErrorAction SilentlyContinue -Force
    if ($found -eq $null) {
        $msg = "[$entry] *** NOT FOUND ***"
    } else {
        $msg = "[$entry] " + $found[0].FullName
    }
    $msg | Out-File $logfile -Encoding ASCII -Append
    Clear-Host
    Write-Host "Check dati progetto TIM: $i out of $tot"
}

# smonto la share
Remove-PSDrive -Name R

# messaggio di chiusura
[System.Windows.MessageBox]::Show("Check lotto terminato",'COMPLETED','Ok','Info')
