<#
Name......: PPPC.ps1
Version...: 25.6.1
Author....: Dario CORRADA

Pipeline per la preparazione di nuovi PC

[241114] commentato il download dei seguenti script, rimossi dalla sequenza di lancio:
    * 3rd_Parties\Wazuh.ps1
    * AD\JoinUser.ps1
    * AzureAD\CreateMSAccount.ps1
#>

# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script. "
        Pause
        Exit
    }
}

# elevated script execution with admin privileges
$ErrorActionPreference= 'Stop'
try {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($testadmin -eq $false) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
        exit $LASTEXITCODE
    }
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

# header 
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
        DownloadFilesFromRepo -Owner $Owner -Repository $Repository -Path $_.path -DestinationPath $($DestinationPath+'\'+$_.name)
    }
    
    if (-not (Test-Path $DestinationPath)) {
        # Destination path does not exist, let's create it
        try {
            New-Item -Path $DestinationPath -ItemType Directory -ErrorAction Stop | out-null
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

# check NuGet
foreach ($pp in (Get-PackageProvider)) {
    if ($pp.Name -eq 'NuGet') {
        $foundit = $pp.Name
    }
}
if ($foundit -ne 'NuGet') {
    $ErrorActionPreference= 'Stop'
    Try {
        Install-PackageProvider -Name "NuGet" -MinimumVersion "2.8.5.208" -Force
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }
    $ErrorActionPreference= 'Inquire'
}

# verifico se winget funziona e se SecureBoot sia su
$info = systeminfo
if ($info[2] -match 'Windows 11') {
    if (!(Confirm-SecureBootUEFI)) {
        [System.Windows.MessageBox]::Show("It seems Secure Boot is disabled",'SECURE BOOT','Ok','Warning') | Out-Null
    }

    winget.exe source update
}

# creo una cartella temporanea e scarico gli script
$tmppath = 'C:\PPPCtemp'
if (Test-Path $tmppath) {
    Remove-Item -Path $tmppath -Recurse -Force > $null
}
New-Item -ItemType directory -Path $tmppath > $null
New-Item -ItemType directory -Path "$tmppath\Modules" > $null
#New-Item -ItemType directory -Path "$tmppath\3rd_Parties" > $null
New-Item -ItemType directory -Path "$tmppath\AD" > $null
#New-Item -ItemType directory -Path "$tmppath\AzureAD" > $null
New-Item -ItemType directory -Path "$tmppath\Updates" > $null
New-Item -ItemType directory -Path "$tmppath\Upkeep" > $null
New-Item -ItemType directory -Path "$tmppath\Safety" > $null
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Modules' -DestinationPath "$tmppath\Modules"
#DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path '3rd_Parties\Wazuh.ps1' -DestinationPath "$tmppath\3rd_Parties"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'AD\Join2Domain.ps1' -DestinationPath "$tmppath\AD"
#DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'AD\JoinUser.ps1' -DestinationPath "$tmppath\AD"
#DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'AzureAD\CreateMSAccount.ps1' -DestinationPath "$tmppath\AzureAD"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Updates\drvUpdate_Win10.ps1' -DestinationPath "$tmppath\Updates"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Updates\Update_Win10.ps1' -DestinationPath "$tmppath\Updates"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Upkeep\Powerize.ps1' -DestinationPath "$tmppath\Upkeep"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Safety\Stargate.ps1' -DestinationPath "$tmppath\Safety"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Safety\PedoMellon.ps1' -DestinationPath "$tmppath\Safety"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Safety\Disable_Bitlocker.ps1' -DestinationPath "$tmppath\Safety"
DownloadFilesFromRepo -Owner 'dcorrada' -Repository 'POWERSHELL' -Path 'Init_PC.ps1' -DestinationPath $tmppath

# download di materiale custom dal branch 'tempus'
New-Item -ItemType directory -Path "$tmppath\miscellaneous\AGM_scripts" > $null
$downbin = $tmppath + '\miscellaneous\AGM_scripts\AGMConfMan_init.ps1'
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/dcorrada/POWERSHELL/tempus/miscellaneous/AGM_scripts/AGMConfMan_init.ps1' -OutFile $downbin -ErrorAction Stop | out-null
$downbin = $tmppath + '\miscellaneous\Quotes.ps1'
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/dcorrada/POWERSHELL/tempus/miscellaneous/Quotes.ps1' -OutFile $downbin -ErrorAction Stop | out-null

# creo i file batch per gli step da eseguire
New-Item -ItemType file -Path "$tmppath\STEP01.cmd" > $null
@"
copy "$tmppath\STEP02.cmd" "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
PowerShell.exe "& "'$tmppath\Init_PC.ps1'
del "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\STEP01.cmd"
"@ | Out-File "$tmppath\STEP01.cmd" -Encoding ASCII -Append

New-Item -ItemType file -Path "$tmppath\STEP02.cmd" > $null
@"
copy "$tmppath\STEP03.cmd" "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
PowerShell.exe "& "'$tmppath\AD\Join2Domain.ps1'
del "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\STEP02.cmd"
"@ | Out-File "$tmppath\STEP02.cmd" -Encoding ASCII -Append

New-Item -ItemType file -Path "$tmppath\STEP03.cmd" > $null
@"
PowerShell.exe "& "'$tmppath\Upkeep\Powerize.ps1'
Pause
PowerShell.exe "& "'$tmppath\miscellaneous\AGM_Scripts\AGMConfMan_init.ps1'
Pause
PowerShell.exe "& "'$tmppath\Safety\Disable_Bitlocker.ps1'
Pause
PowerShell.exe "& "'$tmppath\Updates\drvUpdate_Win10.ps1'
Pause
PowerShell.exe "& "'$tmppath\miscellaneous\Quotes.ps1'
rd /s /q "$tmppath"
del "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\STEP03.cmd"
"@ | Out-File "$tmppath\STEP03.cmd" -Encoding ASCII -Append

# copio il primo batch file per il riavvio successivo
Copy-Item -Path "$tmppath\STEP01.cmd" -Destination "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

# lancio lo STEP00
PowerShell.exe "& ""$tmppath\Updates\Update_Win10.ps1"
