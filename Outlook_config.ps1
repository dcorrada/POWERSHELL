<#
Name......: Outlook_config.ps1
Version...: 21.03.1
Author....: Dario CORRADA

This script will import/export Outlook (Office 365) configuration.
In details this script will manage:
* email accounts
* layout
* PST file attachment
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Outlook_config\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# function for killing Outlook instances
function OutlookKiller {
    $ErrorActionPreference= 'Stop'
    Try {
        $outproc = Get-Process outlook
        Stop-Process -ID $outproc.Id -Force
        Start-Sleep 2
    }
    Catch { 
        [System.Windows.MessageBox]::Show("Check out that all Oulook processes have been closed before go ahead",'TASK MANAGER','Ok','Warning') > $null
    }
    $ErrorActionPreference= 'Inquire'
}

# behaviour form
$form_modalita = FormBase -w 300 -h 190 -text "CONFIGURATION"
$import = RadioButton -form $form_modalita -checked $false -x 30 -y 20 -text "Import Outlook config"
$export  = RadioButton -form $form_modalita -checked $true -x 30 -y 50 -text "Export Outlook config"
OKButton -form $form_modalita -x 90 -y 90 -text "Ok"
$result = $form_modalita.ShowDialog()

# select path where configs have to be saved/retrieved
$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
$OpenFileDialog.AddExtension = $false
$OpenFileDialog.CheckFileExists = $false
$OpenFileDialog.DereferenceLinks = $true
$OpenFileDialog.Filter = "Folders|`n"
$OpenFileDialog.Multiselect = $false
$OpenFileDialog.Title = "Select folder where configs have to be saved/retrieved"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
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
$outlook_workdir = $OpenFileDialog.FileName

# all Oulook instances have to be closed before go ahead
OutlookKiller

if ($import.Checked) {
    Write-Host -ForegroundColor Cyan "*** IMPORTING OUTLOOK CONFIGURATION ***"

<#
# attach archivi di outlook
Write-Host "Caricamento archivi di Outlook"
[string[]]$pst_file_list = Get-Content -Path "C:\Optional Software\PST_files.log"
if ($pst_file_list.count -gt 0) {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

    if ($OfficeVer -eq '16') {
         # con Office 365 bisogna lanciare l'istanza di creazione dell'oggetto due volte, ignorando l'errore iniziale
        $ErrorActionPreference = 'SilentlyContinue'
        $outlook = New-Object -comObject Outlook.Application
        $ErrorActionPreference= 'Inquire'
    }
    $outlook = New-Object -ComObject outlook.application
    $namespace = $outlook.GetNameSpace("MAPI")
    foreach ($pst_file in $pst_file_list) {
        $namespace.AddStore($pst_file)
        Write-Host -ForegroundColor Green $pst_file
    }
} else {
    Write-Host -ForegroundColor Cyan "Nessun archivio trovato"
}
Remove-Item "C:\Optional Software\PST_files.log" -Force
Start-Sleep 5
#>

} elseif ($export.Checked) {
    Write-Host -ForegroundColor Cyan "*** EXPORTING OUTLOOK CONFIGURATION ***"

    # retrieving PST files
    $ErrorActionPreference = 'SilentlyContinue'
    $outlook = New-Object -comObject Outlook.Application # for some unexplained reason, this object instance must be invoked twice
    $ErrorActionPreference= 'Stop'  
    Try {
        $outlook = New-Object -ComObject Outlook.Application
        $PST = $outlook.Session.Stores | Where-Object { ($_.FilePath -like '*.PST') }
        $pstlist = @($PST.FilePath)
        OutlookKiller
        Write-Host -ForegroundColor Yellow "`nPST(s) attached found:"
        foreach ($item in $pstlist) {
            Write-Host "Copying $item"
            Copy-Item -Path $item -Destination $outlook_workdir
        }
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        [System.Windows.MessageBox]::Show("Unable to retrieve PST list from Outlook",'ERROR','Ok','Error')
    }
    $ErrorActionPreference= 'Inquire'

    # retrieving Outlook layout
    Write-Host -ForegroundColor Yellow "`nCopying Outlook layout"
    $outlook_aspect = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook\Outlook.xml"
    Copy-Item -Path $outlook_aspect -Destination $outlook_workdir
    
    # retrieving Outlook accounts
    # see https://turbolab.it/windows-10/come-estrarre-profilo-posta-microsoft-outlook-registro-configurazione-1944 
    Write-Host -ForegroundColor Yellow "`nBackup Outlook account(s)"
    
}