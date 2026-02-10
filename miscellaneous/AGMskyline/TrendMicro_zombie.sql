/*
Questa query incrocia gli asset estratti da SnipeIT con gli hostname indicati 
su TrendMicro allo scopo di individuare se ci fosse qualche liceenza da 
dismettere (ad es. se quell'hostname punti ad un asset non piu' assegnato).
*/
pager less -SFX;

USE AGMskyline;

SELECT TrendMicroparsed.HOSTNAME, EstrazioneAsset.STATUS
FROM TrendMicroparsed
LEFT JOIN EstrazioneAsset 
ON TrendMicroparsed.HOSTNAME = EstrazioneAsset.HOSTNAME
WHERE EstrazioneAsset.STATUS != 'Assegnato';
