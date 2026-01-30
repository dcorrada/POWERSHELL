/*
Questa query incrocia gli asset estratti da SnipeIT con gli hostname indicati 
su TrendMicro allo scopo di individuare se ci fosse qualche liceenza da 
dismettere (ad es. se quell'hostname punti ad un asset non piu' assegnato).
*/
pager less -SFX;

USE AGMskyline;

SELECT ConfigManager.HOSTNAME, EstrazioneAsset.STATUS
FROM ConfigManager
LEFT JOIN EstrazioneAsset 
ON ConfigManager.HOSTNAME = EstrazioneAsset.HOSTNAME
