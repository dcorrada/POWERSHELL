/*
Query per cercare su AD utenze cessate da spostare nella OU appropriata.
Questa query è INDICATIVA: Verificare SEMPRE i singoli record che non si 
tratti di falsi positivi (ad esempio: utenti agganciati sul tenant 
SUCCESSIVAMENTE alla data di aggiornamento delle tabelle; utenti che hanno
cambiato mail nel frattempo; utenze di servizio; ecc.)
*/
pager less -SFX;

USE AGMskyline;

SELECT ADusers.*
FROM ADusers
INNER JOIN EstrazioneUtenti ON  ADusers.USRNAME = EstrazioneUtenti.USRNAME
WHERE ADusers.OU IS NULL OR (ADusers.OU NOT LIKE '%Esterni' AND ADusers.OU NOT LIKE '%Stagisti AGM' AND ADusers.OU NOT LIKE '%Utenti Milano' AND ADusers.OU NOT LIKE '%Utenti Torino')
UNION
SELECT ADusers.*
FROM ADusers
LEFT JOIN o365licenses ON o365licenses.MAIL LIKE CONCAT(ADusers.USRNAME, '%')
WHERE o365licenses.MAIL IS NULL AND ADusers.OU NOT LIKE '%Disable' AND ADusers.USRNAME LIKE '%.%'