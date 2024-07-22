<#
Name......: OneShot.ps1
Version...: 24.07.1
Author....: Dario CORRADA

This script allow to navigate and select single scripts from this repository.
Then it dowload them and launch them locally.

+++ TODO +++
* navigare le cartelle del repo e risalirle
* pinnare gli script nei preferiti
* mostrare una descrizione dello script selezionato via radio button
    (v. https://stackoverflow.com/questions/40715908/invoking-a-scriptblock-when-clicking-a-radio-button)
* creare un launcher che scarica questo script, i moduli e poi lo lancia
* selezionare barnch? (lascerei master di default)
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\OneShot\.ps1$" > $null
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
        Import-Module PowerShellForGitHub
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'PowerShellForGitHub')) {
            Install-Module PowerShellForGitHub -Scope AllUsers -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [PowerShellForGitHub] module: click Ok to restart the script",'RESTART','Ok','warning') > $null
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
                                    BROWSING
******************************************************************************* #>
Set-GitHubConfiguration -DisableTelemetry -SessionOnly
$ParentFolder = 'root'

# filtering only scripts and paths (nor modules folder)
$CurrentItems = (Get-GitHubContent `
    -OwnerName 'dcorrada' `
    -RepositoryName 'POWERSHELL' `
    -BranchName 'master' `
    ).entries | ForEach-Object {
        if ((($_.type -eq 'dir') -or ($_.name -match "\.ps1$")) -and !($_.name -eq 'Modules')) {
            New-Object -TypeName PSObject -Property @{
                NAME    = $_.name
                PATH    = $_.path
                URL     = $_.download_url
            } | Select NAME, PATH, URL   
        }
    }

$isChecked = "none"
$intoBox = ''
do {
    $hmin = ((($CurrentItems.Count) * 30) + 90)
    if ($hmin -lt 300) {
        $hmin = 300
    }
    $adialog = FormBase -w 720 -h $hmin -text "SELECT AN ITEM"
    $they = 20
    $choices = @()
    foreach ($ItemName in ($CurrentItems.Name | Sort-Object)) {
        if (($isChecked -eq 'none') -and ($choices.Count -lt 1)) {
            $gotcha = $true
        } elseif ($ItemName -eq $isChecked) {
            $gotcha = $true
        } else {
            $gotcha = $false
        }
        $choices += RadioButton -form $adialog -x 20 -y $they -checked $gotcha -text $ItemName
        $they += 30 
    }
    TxtBox -form $adialog -x 230 -y 20 -w 450 -h 200 -text $intoBox -multiline $true | Out-Null
    OKButton -form $adialog -x 250 -y 230 -text "Go ahead" | Out-Null
    RETRYButton -form $adialog -x 550 -y 230 -text "Preview" | Out-Null
    $goahead = $adialog.ShowDialog()
    if ($goahead -eq 'RETRY') {
        foreach ($currentOpt in $choices) {
            if ($currentOpt.Checked) {
                $isChecked = "$($currentOpt.Text)"
                <# per il testo in $intobox:
                    * head -n 20 se è uno script
                    * lista elementi se è una cartella
                #>
                $intoBox = "$($currentOpt.Text)"
            }
        }
    }
} while ($goahead -ne 'OK')

