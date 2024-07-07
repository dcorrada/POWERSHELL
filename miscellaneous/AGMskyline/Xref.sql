/*
Tabelle di crossref per fare le join

L'impiego di LEFT JOIN e' dato dal fatto che le tabelle a sinistra (da Snipe, 
da schede assunzione, ecc.) dovrebbero essere la risorsa primaria di riferimento
*/

-- crossrefs sugli asset (laptop)
DROP TABLE IF EXISTS Xhosts;
CREATE TABLE Xhosts (ID INT NOT NULL AUTO_INCREMENT, PRIMARY KEY (ID)) ENGINE=InnoDB
SELECT EstrazioneAsset.HOSTNAME AS 'HOSTNAME', EstrazioneAsset.ID AS 'ESTRAZIONEASSET',
       ADcomputers.ID AS 'ADCOMPUTERS', 
       AzureDevices.ID AS 'AZUREDEVICES', 
       GFIparsed.ID AS 'GFIPARSED',
       TrendMicroparsed.ID AS 'TRENDMICROPARSED',
       CheckinFrom.ID AS 'CHECKINFROM'
FROM EstrazioneAsset
LEFT JOIN ADcomputers ON EstrazioneAsset.HOSTNAME = ADcomputers.HOSTNAME
LEFT JOIN AzureDevices ON EstrazioneAsset.HOSTNAME = AzureDevices.HOSTNAME
LEFT JOIN GFIparsed ON EstrazioneAsset.HOSTNAME = GFIparsed.HOSTNAME
LEFT JOIN TrendMicroparsed ON EstrazioneAsset.HOSTNAME = TrendMicroparsed.HOSTNAME
LEFT JOIN CheckinFrom ON EstrazioneAsset.HOSTNAME = CheckinFrom.HOSTNAME;

-- crossrefs sugli utenti
DROP TABLE IF EXISTS Xusers;
CREATE TABLE Xusers (ID INT NOT NULL AUTO_INCREMENT, PRIMARY KEY (ID)) ENGINE=InnoDB
SELECT ADusers.FULLNAME AS 'FULLNAME', o365licenses.ID AS 'O365LICENSES',
       AzureDevices.ID AS 'AZUREDEVICES', DLmembers.ID AS 'DLMEMBERS',
       SchedeAssunzione.ID AS 'SCHEDEASSUNZIONE',
       SchedeTelefoni.ID AS 'SCHEDETELEFONI',
       SchedeSIM.ID AS 'SCHEDESIM',
       ThirdPartiesLicenses.ID AS 'THIRDPARTIESLICENSES',
       EstrazioneUtenti.ID as 'ESTRAZIONEUTENTI',
       CheckinFrom.ID AS 'CHECKINFROM',
       ADusers.ID AS 'ADUSERS',
       PwdExpire.ID AS 'PWDEXPIRE'
FROM ADusers
LEFT JOIN o365licenses ON ADusers.UPN = o365licenses.MAIL
LEFT JOIN AzureDevices ON ADusers.UPN = AzureDevices.MAIL
LEFT JOIN DLmembers ON ADusers.UPN = DLmembers.EMAIL
LEFT JOIN SchedeAssunzione ON ADusers.UPN = SchedeAssunzione.MAIL
LEFT JOIN SchedeTelefoni ON LOWER(ADusers.FULLNAME) LIKE CONCAT('%', LOWER(SchedeTelefoni.FULLNAME), '%')
LEFT JOIN SchedeSIM ON LOWER(ADusers.FULLNAME) LIKE CONCAT('%', LOWER(SchedeSIM.FULLNAME), '%')
LEFT JOIN ThirdPartiesLicenses ON ADusers.UPN = ThirdPartiesLicenses.MAIL
LEFT JOIN EstrazioneUtenti ON ADusers.UPN = EstrazioneUtenti.EMAIL
LEFT JOIN CheckinFrom ON ADusers.UPN = CheckinFrom.MAIL
LEFT JOIN PwdExpire ON ADusers.USRNAME = PwdExpire.USRNAME;
