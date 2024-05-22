$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

$downbin = 'C:\Users\' + $env:USERNAME + '\Downloads\MPP.ps1'
if (Test-Path $downbin) {
    Remove-Item $downbin -Force
}
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/dcorrada/POWERSHELL/tempus/miscellaneous/MPP.ps1' -OutFile $downbin
PowerShell.exe "& ""$downbin"
Remove-Item $downbin -Force
