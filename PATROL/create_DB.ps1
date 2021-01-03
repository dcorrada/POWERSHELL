# This script creates a crypto key and related encrypted DB

$workdir = Get-Location
Import-Module -Name "$workdir\Moduli\FileCryptography.psm1"

# crypto key
$key = New-CryptographyKey -Algorithm AES -AsPlainText
$key | Out-File "$workdir\crypto.key" -Encoding ASCII -Append

# encrypted DB
$whoami = $env:UserName
"SCRIPT;USER" | Out-File "$workdir\PatrolDB.csv" -Encoding ASCII -Append
"UpdateDB;$whoami" | Out-File "$workdir\PatrolDB.csv" -Encoding ASCII -Append
$securekey = ConvertTo-SecureString $key -AsPlainText -Force
Protect-File "$workdir\PatrolDB.csv" -Algorithm AES -Key $securekey -RemoveSource
