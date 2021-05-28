<#
Name......: AcrobatFix.ps1
Version...: 21.05.1
Author....: Dario CORRADA

This script aims to restore AcrobatDC Enterprise licence .
The script could be run also if you would like to reserialize Acrobat with a 
different serial number. 

See also https://helpx.adobe.com/acrobat/kb/how-to-Re-Serialize-Acrobat-using-the-APTEE-tool.html
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AcrobatFix\.ps1$" > $null
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

# copy APTEE tool
New-Item -ItemType directory -Path "C:\TEMPUS" > $null
Copy-Item "$workdir\adobe_prtk.exe" -Destination "C:\TEMPUS" -Force > $null

# refresh Acrobat licence
$anyprov = [System.Windows.MessageBox]::Show("Could you provide an existing prov.xml file?",'RESERIALIZE','YesNo','Info')
if ($anyprov -eq "Yes") {

    # select prov.xml file
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = "C:\"
    $OpenFileDialog.filter = 'XML file (*.xml)| *.xml'
    $OpenFileDialog.ShowDialog() | Out-Null
    $provfile = $OpenFileDialog.filename

    # restore licence
    $ErrorActionPreference= 'SilentlyContinue'
    Set-Location "C:\TEMPUS"
    $opts = ("--tool=VolumeSerialize", "--provfile=$provfile", "–stream")
    $returncode = .\adobe_prtk.exe @opts
    $ErrorActionPreference= 'Inquire'
    Write-Host -NoNewline "Checking licence..."
    Start-Sleep -s 3
    if ($returncode -eq "Return Code = 0") {
        Write-Host -ForegroundColor Green " PASSED"
    } else {
        Write-Host -ForegroundColor Red " FAILED"
        $reserial = [System.Windows.MessageBox]::Show("Would you use another serial number?",'RESERIALIZE','YesNo','Info')
    }
}

# alternative procedure
if (($anyprov -eq "No") -or ($reserial -eq "Yes")) {
    $ErrorActionPreference= 'SilentlyContinue'
    Set-Location "C:\TEMPUS"
    $serialnew = Read-Host "Please enter the serial number"

    # create a new prov.xml file
    $opts = ("--tool=VolumeSerialize", "--generate", "--serial=$serialnew", '--leid=V7{}AcrobatCont-12-Win-GM', "--regsuppress=ss", "--eulasuppress", "--provfile=C:\TEMPUS\prov.xml")
    $returncode = .\adobe_prtk.exe @opts
    Start-Sleep -s 3
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
    $OpenFileDialog.filter = 'Executable (*.xml)| *.xml'
    $OpenFileDialog.filename = 'prov'
    $OpenFileDialog.ShowDialog() | Out-Null
    $destfile = $OpenFileDialog.filename
    Copy-Item "C:\TEMPUS\prov.xml" -Destination $destfile -Force > $null

    # unserialize the existing serial key
    $opts = ("--tool=UnSerialize", '--leid=V7{}AcrobatETLA-12-Win-GM', "--deactivate", "--force", "–removeSWTag")
    $returncode = .\adobe_prtk.exe @opts
    Start-Sleep -s 3

    # reserialize Acrobat using the newly created prov.xml file
    $opts = ("--tool=VolumeSerialize",  "--provfile=C:\TEMPUS\prov.xml", "--stream")
    $returncode = .\adobe_prtk.exe @opts
    Start-Sleep -s 3

    $ErrorActionPreference= 'Inquire'
}

# clean tempfiles
Set-Location $workdir
Remove-Item "C:\TEMPUS" -Recurse -Force
