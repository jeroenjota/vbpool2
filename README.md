# VBPool 2.0
VBPool2.0 of VBP2 is een klein Windows programmaatje, geschreven in Visual Basic 6.0, voor een uitgebreide voetbalpool tijdens een groot toernooi zoals de WK of de EK.
Op dit moment is het nog niet geschikt voor de clubtoernooien zoals de nationale competitie of de Championsleague, wellicht iets voor versie 3.0.

# TODO
Het systeem is nog lang niet klaar
Wat er nog moet:
- Deelnemers formulieren
- Deelnemers invoeren (webinterface ?)
- Wedstrijduitslagen verwerken
- Rapportages:
  * Poolstand op volgorde/alfabetisch
  * Grafiek opbouw van punten per deelnemer
  * Punten per deelnemers per onderdeel
  * Toernooi - stand van zaken, komende wedstrijden
  * En nog zo wat ...
  
## Systeem
Het systeem werkt als volgt:

- Deelnemers vullen van te voren een formulier in met al hun voorspellingen. 
- Elke juist ingevulde voorspelling kan punten opleveren.
- De deelnemer met de meeste punten aan het eind van het toernooi wint.
- Per wedstrijddag kunnen dagprijzen worden gewonnen

Verder kan worden ingesteld:
- De hoogte van de inleg
- De prijzen
  * Dagprijzen voor het hoogste puntenaantal van de dag, de hoogste en laagste positie in de pool na afloop van de dag
  * Eindprijzen: Maximaal 4. Per prijs kan een percentage worden ingesteld van de inleg (minus de uitgekleerde dagprijzen) 
  * De Rode Lantaarn; de prijs voor de laatste in de pool (bijvoorbeeld de inleg terug)

Voor alle prijzen geldt dat als er meerdere winnaars zijn de prijs wordt verdeeld.

### Mogelijke voorspellingen
De organisator bepaald voor welke voorspelling er punten kunnen worden verdiend en hoeveel punten een juiste voorspelling oplevert:
- Groepsstand
- Finaleteams (met of zonder juiste plaats in het schema)
- Eindstand
- Topscorer en het aantal doelpunten
- Diverse soorten aantallen zoals 
  * Gele/Rode kaarten
  * Doelpunten
  * Penalties
  * Gelijke spelen
- Wedstrijduitslagen
  * Ruststand
  * Eindstand
  * Toto
- Het totaal aantal doelpunten op één dag

Onderling hoeven de ingevulde voorspellingen niet te kloppen.
Het principe is dat elke juist ingevulde voorspelling telt. 
> Dus je kunt een wedstijduitslag invullen die onderling niet klopt, bijvoorbeeld rust 0-1, eindstand 2-0 en toto een 3. Als de einduitslag 2-0 is heb je daarvoor dan toch punten verdiend.
Er zijn voorspellingen waarvan nog niet bekend is welke ploegen daar spelen, maar daarvoor kun je al bijvoorbeeld wel de juiste uitslag invullen,. ook al zijn er hele andere ploegen ingevuld.

## Installatie
Kopieer uit de map `compiled` de bestanden `vbp2.exe` en `vbpSetup.mdb` naar een lege map/directory.
Start **vbp2** door erop te dubbelklikken.
Je kunt het programma ook opstarten vanuit een commando-venster.

Bij de eerste opstart wordt er een kopie gemaakt van de vbpoolSetup.mdb database. Daarna wordt de gebruiker gevrqagd de gegevens van de organisator in te vullen. 
Vervolgens wordt vanaf het internet de gegevens van het laatst bekende voetbaltoernooi ingelezen. Alle teams, pool-indelingen en wedstrijden worden naar de lokale database gekopieerd.

## Vereisten
Een computer met daarop het Windows besturingesprogramm. Het programma is niet getest op Windows versies eerder dan Windows 7.

## Admin
Start vbp2 vanuit een terminal (commando ventser) met `vbp2 admin` om speciale rechten te kunnen verkeijgen voor het veranderen van de toernooi gegevens.
In geval van een totale chaos kunnen te allen tijde de data van de server worden teruggezet.

