<#
Name......: MonthlyFilter.ps1
Version...: 25.02.1
Author....: Dario CORRADA

Questo script e' un proof of concept per ottenere una reportistica mensile del 
bilancio di licenze Microsoft 365 assegnate su recuperate. Lo script leggera' 
dal file Excel prodotto con lo script AssignedLicensesSDK.ps1. 

Comincera' leggendo il worksheet "Assigned_Licenses" per ottenere una lista 
preliminare di licenze assegnate in uno specifico mese.

Successivamente andra' a parsare le voci presenti nel worksheet "Orphaned": 
* cerchera' le voci annotate come "user no longer exists on tenant", o come
  "license dismissed for this user", per valutare se sia stata 
  effettivamente recuperata una licenza;
* cerchera' le voci annotate come "assigned license(s) to this user": in 
  questo caso occorrera' valutare sia stata assegnata una licenza, nel 
  periodo temporale di riferimento, e se poi sia stata successivamente 
  recuperata a causa di uno switch (ie passaggio da Basic a Standard) 
  oppure se l'utente sia cessato nello stesso periodo
#>