/*
Query per individuare gli utenti, in forze, che hanno la password scaduta.
La tabella PwdExpire fa riferimento all'utenza di dominio e non all'account 
Microsoft 365 (la gestione credenziali su quest'ultimo viene mediata dalla MFA).

I record presentati il campo STATUS come "ASSUNTO" sono utenze per cui la 
password docrà essere rinnovata, quelli presentati il campo STATUS come null 
occorre verificare prima il motivo per cui non ci sia una scheda assunzione 
(utenze precedenti all'introduzione delle schede, piva, utenze senza asset 
assegnato, utenze di servizio, ecc.)
*/

SELECT DISTINCT PwdExpire.*, SchedeAssunzione.STATUS
FROM PwdExpire
LEFT JOIN Xusers ON PwdExpire.ID = Xusers.PWDEXPIRE
LEFT JOIN SchedeAssunzione ON Xusers.SCHEDEASSUNZIONE = SchedeAssunzione.ID
WHERE ((SchedeAssunzione.STATUS = 'ASSUNTO') 
  OR (SchedeAssunzione.STATUS IS NULL))
  AND PwdExpire.PWD_EXPIRED = 'True';
