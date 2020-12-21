# Script per la creazione di una chiave di criptaggio e relativo DB

$workdir = Get-Location
Import-Module -Name "$workdir\Moduli\FileCryptography.psm1"

# creo la chiave
$key = New-CryptographyKey -Algorithm AES -AsPlainText
$key | Out-File "$workdir\crypto.key" -Encoding ASCII -Append

# creo il DB
$whoami = $env:UserName
"SCRIPT;USER" | Out-File "$workdir\PatrolDB.csv" -Encoding ASCII -Append
"UpdateDB;$whoami" | Out-File "$workdir\PatrolDB.csv" -Encoding ASCII -Append
$securekey = ConvertTo-SecureString $key -AsPlainText -Force
Protect-File "$workdir\PatrolDB.csv" -Algorithm AES -Key $securekey -RemoveSource
