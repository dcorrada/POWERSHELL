/*
Questa query dovrebbe incrociare dati per recuperare le versioni dei sistemi
operativi installati
*/

/*
SELECT EstrazioneUtenti.FULLNAME, EstrazioneUtenti.EMAIL,
       EstrazioneAsset.HOSTNAME,
       ADcomputers.OS AS 'AD_OS',
       CONCAT(AzureDevices.OSTYPE, ' ', AzureDevices.OSVER) AS 'AZURE_OS',
       GFIparsed.OS AS 'GFI_OS'
FROM EstrazioneAsset
LEFT JOIN GFIparsed ON EstrazioneAsset.HOSTNAME = GFIparsed.HOSTNAME
LEFT JOIN ADcomputers ON EstrazioneAsset.HOSTNAME = ADcomputers.HOSTNAME
LEFT JOIN AzureDevices ON EstrazioneAsset.HOSTNAME = AzureDevices.HOSTNAME
LEFT JOIN EstrazioneUtenti ON EstrazioneAsset.USRNAME = EstrazioneUtenti.USRNAME
WHERE EstrazioneAsset.STATUS = 'Assegnato'
*/

-- query usando tabella ponte Xhosts
SELECT DISTINCT EstrazioneUtenti.FULLNAME, EstrazioneUtenti.EMAIL,
       EstrazioneAsset.HOSTNAME,
       ADcomputers.OS AS 'AD_OS',
       CONCAT(AzureDevices.OSTYPE, ' ', AzureDevices.OSVER) AS 'AZURE_OS',
       GFIparsed.OS AS 'GFI_OS'
FROM Xhosts
LEFT JOIN EstrazioneAsset ON Xhosts.ESTRAZIONEASSET = EstrazioneAsset.ID
LEFT JOIN GFIparsed ON Xhosts.GFIPARSED = GFIparsed.ID
LEFT JOIN ADcomputers ON Xhosts.ADCOMPUTERS = ADcomputers.ID
LEFT JOIN AzureDevices ON Xhosts.AZUREDEVICES = AzureDevices.ID
LEFT JOIN EstrazioneUtenti ON EstrazioneAsset.USRNAME = EstrazioneUtenti.USRNAME
WHERE EstrazioneAsset.STATUS = 'Assegnato'
