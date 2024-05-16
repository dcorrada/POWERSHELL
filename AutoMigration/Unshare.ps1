<#
Name......: Unshare.ps1
Version...: 20.1.2
Author....: Dario CORRADA

This script remove Volume C: for sharing and delete temporary directory
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AutoMigration\\Unshare\.ps1$" > $null
$repopath = $matches[1]

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# unsharing C: volume
$grants = Get-SmbShare
foreach ($item in $grants.Name) {
    if ($item -eq 'C') {
        Write-Host -NoNewline -ForegroundColor Yellow "`nUnsharing C: volume..."
        $ErrorActionPreference = 'Stop'
        try {
            Remove-SmbShare -Name "C" -Force
            Write-Host -ForegroundColor Green " DONE"
        }
        catch {
            Write-Host -ForegroundColor Red " FAILED"
            Write-Host -ForegroundColor Red "$($error[0].ToString())"
            Pause
        }
        $ErrorActionPreference = 'Inquire'
    }
}

# remove tmpdir
$tmppath = 'C:\AUTOMIGRATION'
if (Test-Path $tmppath) {
    $answ = [System.Windows.MessageBox]::Show("Remove temp data?",'DELETE','YesNo','Info')
    if ($answ -eq "Yes") {
        Write-Host -NoNewline -ForegroundColor Yellow "`nRemoving temporary data..."
        Remove-Item "$tmppath" -Recurse -Force
        Write-Host -ForegroundColor Green " DONE"
    }
}