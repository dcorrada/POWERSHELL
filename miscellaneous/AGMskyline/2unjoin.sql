/*
Query per cercare hostname non piu' assegnati su SnipeIT 
ma ancora presenti su AD (e quindi da rimuovere da dominio)
*/
pager less -SFX;

USE AGMskyline;

SELECT EstrazioneAsset.HOSTNAME, EstrazioneAsset.STATUS, ADcomputers.OU
FROM EstrazioneAsset
INNER JOIN ADcomputers ON  ADcomputers.HOSTNAME = EstrazioneAsset.HOSTNAME
WHERE EstrazioneAsset.STATUS NOT LIKE 'Assegnato';
