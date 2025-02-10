These scripts require Powershell ActiveDirectory module

Add the following chunck to your codes:

# check Active Directory module
if ((Get-Module -Name ActiveDirectory -ListAvailable) -eq $null) {
    $ErrorActionPreference= 'Stop'
    try {
        Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability –Online
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to install RSAT"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'
}

If not already installed, download RSAT package (Remote Server Administration Tools)
available on https://www.microsoft.com/en-us/download/details.aspx?id=45520