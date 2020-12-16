function Patrol {
    param ($scriptname)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName PresentationFramework
    Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\Forms.psm1'

    # Recupero le credenziali AD
    $credit = LoginWindow

    if ((new-object directoryservices.directoryentry "",$usr,$MaskedTextBox.Text).psbase.name -ne $null) {
        # carico il modulo per criptare/decriptare il DB
        Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\FileCryptography.psm1'

        Copy-Item -Path "\\192.168.2.251\Dario\SCRIPT\PATROL\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
        $stringa = Get-Content '\\192.168.2.251\Dario\SCRIPT\PATROL\crypto.key'
        $chiave = ConvertTo-SecureString $stringa -AsPlainText -Force
        Unprotect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Algorithm AES -Key $chiave -RemoveSource > $null

        $filecontent = Get-Content -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
        $allowed = @()
        foreach ($newline in $filecontent) {
            ($script, $utente) = $newline.Split(';')
            if ($script -eq $scriptname) {
                $allowed += $utente
            }
        }
        
        Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"

        if ($allowed -contains 'everyone') { 
            Write-Host -ForegroundColor Cyan "Patrol disabled"
            $status = 'granted'
        } elseif ($allowed -contains $usr) {
            Write-Host -ForegroundColor Green "Patrol unlocked"
            $status = 'granted'
        } else {
            Write-Host -ForegroundColor Red "Patrol locked"
            $status = 'blocked'
        }        
       
    } else {
        [System.Windows.MessageBox]::Show("Password o username errati",'ATTENZIONE','Ok','Error')
        Exit
    }

    # Scrivo il record di accesso
    net stop workstation /y > $null
    net start workstation > $null
    Start-Sleep 3
    New-PSDrive -Name Z -PSProvider FileSystem -Root '\\192.168.2.251\Dario\SCRIPT\PATROL' -Credential $credit > $null

    $rec_data = Get-Date -UFormat "%d/%m/%Y-%H:%M"
    $new_record = @(
        $rec_data,
        $usr,
        $scriptname,
        $status
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File "Z:\ACCESSI_PATROL.csv" -Encoding ASCII -Append
    

    if ($status -eq 'blocked') {
        [System.Windows.MessageBox]::Show("L'utente non e' abilitato all'utilizzo di questo script",'ATTENZIONE','Ok','Error')
        Exit
    }

    Remove-PSDrive -Name Z

    return $credit
}
Export-ModuleMember -Function Patrol