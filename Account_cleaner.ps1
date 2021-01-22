<#
Name......: Account_cleaner.ps1
Version...: 19.12.1
Author....: Dario CORRADA

This script removes account from local computer
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Account_cleaner\.ps1$" > $null
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

# create temporary directory
$tmppath = 'C:\TEMPSOFTWARE'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
}

# getting users list
$userlist = Get-ChildItem C:\Users

# control panel
$hsize = 150 + (30 * $userlist.Count)
$form_panel = FormBase -w 300 -h $hsize -text "USER FOLDERS"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(200,30)
$label.Text = "Select users to be deleted:"
$form_panel.Controls.Add($label)
$vpos = 50
$boxes = @()
foreach ($elem in $userlist) {
    $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $elem
    $vpos += 30
}
$vpos += 20
OKButton -form $form_panel -x 90 -y $vpos -text "Ok"
$result = $form_panel.ShowDialog()

foreach ($box in $boxes) {
    if ($box.Checked -eq $true) {
        $theuser = $box.Text

        Write-Host "Removing $theuser..."

        # search and remove keys
        $record = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
            Get-ItemProperty | Where-Object {$_.ProfileimagePath -match "C:\\Users\\$theuser" } | Select-Object -Property ProfileimagePath, PSChildName
        $keypath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $record.PSChildName
        Remove-Item -Path $keypath -Recurse
        
        # removing user folder
        $thepath = 'C:\Users\' + $theuser
        Move-Item -Path $thepath -Destination 'C:\TEMPSOFTWARE' -Force
    }
}

$ErrorActionPreference = 'Stop'
Try {
    # Cambio i permessi su file e cartelle
    # $string = '/F C:\TEMPSOFTWARE /R /D Y'
    # Start-Process -Wait takeown.exe $string

    # elevate Explorer
    taskkill /F /IM explorer.exe
    Start-Sleep 3
    Start-Process C:\Windows\explorer.exe

    $path = "C:\TEMPSOFTWARE"
    $shell = new-object -comobject "Shell.Application"
    $item = $shell.Namespace(0).ParseName("$path")
    $item.InvokeVerb("delete")
    Clear-RecycleBin -Force -Confirm:$false
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    pause
}
$ErrorActionPreference = 'Inquire'
