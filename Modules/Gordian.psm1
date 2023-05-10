# see also https://medium.com/@sumindaniro/encrypt-decrypt-data-with-powershell-4a1316a0834b

# create an encryption key file
function CreateKeyFile {
    param ($keyfile)

    $ErrorActionPreference= 'Stop'
    Try {
        $EncryptionKeyBytes = New-Object Byte[] 32
        [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($EncryptionKeyBytes)
        $EncryptionKeyBytes | Out-File -FilePath "$keyfile"
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
    }

    return $EncryptionKeyBytes
}
Export-ModuleMember -Function CreateKeyFile

# encrypt data from a plain text file
function EncryptFile {
    param ($keyfile, $infile, $outfile, $purge = $true)

    $ErrorActionPreference= 'Stop'
    Try {
        $plain_content = Get-Content -Path "$infile"
        $EncryptionKeyData = Get-Content -Path "$keyfile"
        $secured_content = ConvertTo-SecureString "$plain_content" -AsPlainText -Force
        $encrypted_content = ConvertFrom-SecureString $secured_content -Key $EncryptionKeyData
        $encrypted_content | Out-File -FilePath "$outfile"
        if ($purge) {
            Remove-Item -Path "$infile" -Force
        }
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
    }

    return $encrypted_content
}
Export-ModuleMember -Function EncryptFile

# decrypt data from an ecnrypted file
function DecryptFile {
    param ($keyfile, $infile)

    $ErrorActionPreference= 'Stop'
    Try {
        $EncryptionKeyData = Get-Content -Path "$keyfile"
        $encrypted_content = Get-Content -Path "$infile"
        $secured_content = ConvertTo-SecureString $encrypted_content -Key $EncryptionKeyData
        $plain_content = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secured_content))
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
    }

    return $plain_content
}
Export-ModuleMember -Function DecryptFile
