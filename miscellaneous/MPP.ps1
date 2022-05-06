<#
Name......: MPP.ps1
Version...: 22.05.1
Author....: Dario CORRADA

Pipeline per manutenzione programmata
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

function DownloadFilesFromRepo {
Param(
    [string]$Owner,
    [string]$Repository,
    [string]$Path,
    [string]$DestinationPath
    )

    $baseUri = "https://api.github.com/"
    $args = "repos/$Owner/$Repository/contents/$Path"
    $wr = Invoke-WebRequest -UseBasicParsing -Uri $($baseuri+$args)
    $objects = $wr.Content | ConvertFrom-Json
    $files = $objects | where {$_.type -eq "file"} | Select -exp download_url
    $directories = $objects | where {$_.type -eq "dir"}
    
    $directories | ForEach-Object { 
        DownloadFilesFromRepo -Owner $Owner -Repository $Repository -Path $_.path -DestinationPath $($DestinationPath+$_.name)
    }
    
    if (-not (Test-Path $DestinationPath)) {
        # Destination path does not exist, let's create it
        try {
            New-Item -Path $DestinationPath -ItemType Directory -ErrorAction Stop
        } catch {
            throw "Could not create path '$DestinationPath'!"
        }
    }

    foreach ($file in $files) {
        $fileDestination = Join-Path $DestinationPath (Split-Path $file -Leaf)
        try {
            Invoke-WebRequest -UseBasicParsing -Uri $file -OutFile $fileDestination -ErrorAction Stop | out-null
            "Fetching '$($file)'..."
        } catch {
            throw "Unable to download '$($file.path)'"
        }
    }
}

# creo una cartella temporanea e scarico gli script
$tmppath = 'C:\MPPtemp'
if (Test-Path $tmppath) {
    Remove-Item -Path $tmppath -Recurse -Force > $null
}
New-Item -ItemType directory -Path $tmppath > $null
New-Item -ItemType directory -Path "$tmppath\Modules" > $null
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Modules' -DestinationPath "$tmppath\Modules"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Avira_wrapper.ps1' -DestinationPath $tmppath
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Ccleaner_wrapper.ps1' -DestinationPath $tmppath
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Malwarebytes_wrapper.ps1' -DestinationPath $tmppath
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Update_Win10.ps1' -DestinationPath $tmppath
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'CleanOptimize.ps1' -DestinationPath $tmppath
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'SafetyScan.ps1' -DestinationPath $tmppath

# Informazioni
$usrname = $env:USERNAME
$hostname = $env:COMPUTERNAME
$usrdir = $env:USERPROFILE
if ($env:USERDNSDOMAIN) {
    $dom = $env:USERDNSDOMAIN
} else {
    $dom = '.'
}
$wifi = 'disconnected'
$ether = 'disconnected'
$netlist = Get-NetAdapter
foreach ($item in $netlist) {
    if (($item.Name -eq 'Ethernet') -and ($item.Status -eq 'Up')) {
        $ether = (Get-NetIPConfiguration | Where-Object { $_.InterfaceAlias -eq 'Ethernet' }).IPv4Address.IPAddress
    } elseif (($item.Name -eq 'Wi-Fi') -and ($item.Status -eq 'Up')) {
        $ether = (Get-NetIPConfiguration | Where-Object { $_.InterfaceAlias -eq 'Wi-Fi' }).IPv4Address.IPAddress
    }
}
[System.Windows.MessageBox]::Show("HOSTNAME:`t$hostname`nUSERNANE:`t$dom\$usrname`nUSERDIR:`t`t$usrdir`nWI-FI:`t`t$wifi`nLAN:`t`t$ether",'INFO','Ok','Info') > $null

# importo le mia libreria grafica
Import-Module -Name "$tmppath\Modules\Forms.psm1"

# pannello di controllo
$swlist = @{}
$form_panel = FormBase -w 350 -h 330 -text "CONTROL PANEL"
$swlist['01-Avira'] = CheckBox -form $form_panel -checked $true -x 20 -y 20 -text "1. Avira software updater"
$swlist['02-Malwarebytes'] = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "2. Malwarebytes launcher"
$swlist['03-MSERT'] = CheckBox -form $form_panel -checked $false -x 20 -y 80 -text "3. MS Safety Scanner"
$swlist['04-Ccleaner'] = CheckBox -form $form_panel -checked $true -x 20 -y 110 -text "4. Ccleaner launcher"
$swlist['06-Winupdate'] = CheckBox -form $form_panel -checked $true -x 20 -y 140 -text "5. Windows 10 updates"
$swlist['05-Defrag'] = CheckBox -form $form_panel -checked $false -x 20 -y 170 -text "6. Storage cleaner"
OKButton -form $form_panel -x 100 -y 220 -text "Ok"
$result = $form_panel.ShowDialog()

foreach ($item in ($swlist.Keys | Sort-Object)) {
    if ($swlist[$item].Checked -eq $true) {
        Write-Host -ForegroundColor Blue "[$item]"
        if ($item -eq '01-Avira') {
            PowerShell.exe "& ""$tmppath\Avira_wrapper.ps1"
            [System.Windows.MessageBox]::Show("Click Ok to next step...",'WAITING','Ok','Info') > $null
        } elseif ($item -eq '04-Ccleaner') {
            PowerShell.exe "& ""$tmppath\Ccleaner_wrapper.ps1"
            [System.Windows.MessageBox]::Show("Click Ok to next step...",'WAITING','Ok','Info') > $null
        } elseif ($item -eq '02-Malwarebytes') {
            PowerShell.exe "& ""$tmppath\Malwarebytes_wrapper.ps1"
            [System.Windows.MessageBox]::Show("Click Ok to next step...",'WAITING','Ok','Info') > $null
        } elseif ($item -eq '03-MSERT') {
            PowerShell.exe "& ""$tmppath\SafetyScan.ps1"
            [System.Windows.MessageBox]::Show("Click Ok to next step...",'WAITING','Ok','Info') > $null
        } elseif ($item -eq '05-Defrag') {
            # creo un file batch, che verra' lanciato al prossimo reboot
            New-Item -ItemType file -Path "$tmppath\STEP01.cmd" > $null
@"
PowerShell.exe "& "'$tmppath\CleanOptimize.ps1'
pause
rmdir /s /q "C:\MPPtemp"
del "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\STEP01.cmd"
"@ | Out-File "$tmppath\STEP01.cmd" -Encoding ASCII -Append
            Copy-Item -Path "$tmppath\STEP01.cmd" -Destination "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
            [System.Windows.MessageBox]::Show("Storage cleaner is planned for the next boot",'DEFRAG','Ok','Info') > $null
        } elseif ($item -eq '06-Winupdate') {
            PowerShell.exe "& ""$tmppath\Update_Win10.ps1"
        }
    }    
}

# rimuovo la cartella temporanea
if ($swlist['05-Defrag'].Checked -eq $false) {
    Remove-Item 'C:\MPPtemp' -Recurse -Force
}
