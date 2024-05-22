# script di prova per la gestione degli array

$oldarray = (1,2,3,4,5)

$newarray = @()

do { # itero un ciclo per svuotare completamente il vecchio array
    $elem,$oldarray = $oldarray # faccio uno shift dell'elemento di testa sul vecchio array
    $newarray += $elem # faccio un push in coda dell'elemento sul nuovo array
} while ($oldarray.Count -gt 0)