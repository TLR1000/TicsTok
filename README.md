# TicsTok
automated tics mail sender
Gebruikt een google account voor email, een google sheet en google script.

Bestaat uit 2 functionele delen:

ProcessIncoming - handelt alle mails af voor sub, unsub en wijzigingen.

Vrijdagrun - het proces dat iedere vrijdag draait om de files te maken en te versturen

Het proces maakt gebruik van een sheet waar de users en mailadressen worden bijgehouden, en na verzending de timestamp gezet wordt.
