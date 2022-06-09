These scripts require Powershell ActiveDirectory module

Add the following chunck to your codes:

# import Active Directory module
$ErrorActionPreference= 'Stop'
try {
    Import-Module ActiveDirectory
} catch {
    Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
    Import-Module ActiveDirectory
}
$ErrorActionPreference= 'Inquire'

If not already installed, download RSAT package (Remote Server Administration Tools)
available on https://www.microsoft.com/en-us/download/details.aspx?id=45520