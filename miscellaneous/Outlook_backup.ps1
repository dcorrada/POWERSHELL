<#

*** DANGER *** DON'T RUN THIS SCRIPT! IT HAS NOT BEEN COMPLETELY TESTED AND YOUR MAILS COULD BE LOST

Questo script sposta le mail in arrivo, inviate e archiviate verso un file PST di backup
Vedi templato su https://community.idera.com/database-tools/powershell/ask_the_experts/f/powershell_for_exchange-24/22427/move-mail-to-a-pst 
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

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

$answ = [System.Windows.MessageBox]::Show("Posso chiudere Outlook per procedere al backup?",'BACKUP','YesNo','Info')
if ($answ -eq 'Yes') {
    OutlookKiller
} else {
    $answ = [System.Windows.MessageBox]::Show("Ricordati di fare il backup",'BACKUP','Ok','Warning')
    Exit
}

# per qualche motivo sconosciuto devo instanziare l'oggetto due volte
$ErrorActionPreference = 'SilentlyContinue'
$outlook = New-Object -ComObject Outlook.Application
$ErrorActionPreference= 'Inquire'

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNameSpace("MAPI")

<# Qui di seguito un paio di comandi su come esplorare cosa c'e' in Outlook

# lista di account, PST e distribution list collegate
$namespace.Folders | Select Name 

# lista delle cartelle dell'account di posta e del PST
$namespace.Folders["dario.corrada@agmsolutions.net"].Folders | Select Name 
$namespace.Folders["Archivio_Posta"].Folders | Select Name 

#>

# definisco le cartelle sorgente e destinazione su cui spostare i messaggi
$SourceDest = [ordered]@{}
$SourceDest['Inbox'] = 'Archivio'
$SourceDest['Archivio'] = 'Archivio'
$SourceDest['Posta inviata'] = 'Inviata'

# sposto le mail sul PST
foreach ($from_folder in $SourceDest.Keys) {
    $to_folder = $SourceDest[$from_folder]
    Write-Host -ForegroundColor Yellow "`nDA [$from_folder] A [$to_folder]"
    $Source = $namespace.Folders['dario.corrada@agmsolutions.net'].Folders[$from_folder]
    $Dest = $namespace.Folders['Archivio_Posta'].Folders[$to_folder]
    $Messages = $Source.Items
    foreach ($msg in $Messages) {
        Write-Host "Backup di <$($msg.Subject)>"
        [void]$msg.Move($Dest)
    }
}

# chiudo Outlook in modalita' admin e lo riavvio il client
OutlookKiller
Start-Process outlook