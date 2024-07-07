/*
Ricerca incongruenze tra fonti

Per fare delle FULL OUTER JOIN in MySQL vedi pure
https://stackoverflow.com/questions/4796872/how-can-i-do-a-full-outer-join-in-mysql 
*/

-- check di indirizzi mail tra SnipeIT e il tenant di Microsoft365
SELECT EstrazioneUtenti.USRNAME AS 'SNIPE_USR', EstrazioneUtenti.EMAIL AS 'SNIPE_MAIL',
       o365licenses.MAIL AS 'O365MAIL'
FROM EstrazioneUtenti
FULL OUTER JOIN o365licenses ON EstrazioneUtenti.EMAIL = o365licenses.MAIL


SELECT EstrazioneUtenti.USRNAME AS 'SNIPE_USR', EstrazioneUtenti.EMAIL AS 'SNIPE_MAIL',
       o365licenses.MAIL AS 'O365MAIL'
FROM EstrazioneUtenti
LEFT JOIN o365licenses ON EstrazioneUtenti.EMAIL = o365licenses.MAIL
UNION ALL
SELECT EstrazioneUtenti.USRNAME AS 'SNIPE_USR', EstrazioneUtenti.EMAIL AS 'SNIPE_MAIL',
       o365licenses.MAIL AS 'O365MAIL'
FROM EstrazioneUtenti
RIGHT JOIN o365licenses ON EstrazioneUtenti.EMAIL = o365licenses.MAIL
WHERE EstrazioneUtenti.EMAIL IS NULL

-- check di username tra SnipeIT e AD
-- (non ho ancora creatio una tabella con le utenze di dominio su AD ;-p)