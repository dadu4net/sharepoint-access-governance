# Referentie
## Prompt
Maak een HTML-pagina voor overzicht van externe of gasttoegang tot SharePoint-sites en documenten. De pagina bevat twee doorzoekbare tabbladen, een vernieuwfunctie via gedelegeerde Microsoft Graph-autorisatie en CSV/JSON-export.

## Vragen vanuit Copilot
Geen aanvullende vragen gesteld. De HTML is configureerbaar gemaakt voor Tenant ID, Client ID, redirect-URI en scanniveau.

## Vervolg en gemaakte keuzes
- Tabblad 1: sitenaam, URL en aantal externen.
- Tabblad 2: naam externe, e-mail, sitenaam en site-URL.
- Interactieve aanmelding met MSAL en gedelegeerde Graph-scopes.
- Optionele diepe scan van bibliotheken en items.
- Lokale opslag van configuratie en laatste scanresultaat in de browser.
- CSV- en JSON-export.
- Noppa-vormgeving met accentkleur #F2B82C.

## Belangrijke beperking
De uitkomst is beperkt tot gegevens die Microsoft Graph retourneert binnen de effectieve rechten en scopes van de ingelogde gebruiker. Een browserdashboard is geen garantie op een volledig tenantbreed overzicht. Groepslidmaatschappen, overervende machtigingen en anonieme links kunnen aanvullende verwerking of een beheerde backend vereisen.


## Vervolgvraag
Maak `README.md` van de instructies.

## Resultaat
Een uitgebreide README is gemaakt met instructies voor app-registratie, gedelegeerde Graph-rechten, configuratie, gebruik, beveiliging, beperkingen, productiearchitectuur en probleemoplossing.

## Vervolgvraag
Maak `README.md` van de instructies.

## Resultaat
Een uitgebreide README is gemaakt met instructies voor app-registratie, gedelegeerde Graph-rechten, configuratie, gebruik, beveiliging, beperkingen, productiearchitectuur en probleemoplossing.


## Vervolgvraag
Pas het HTML-bestand daadwerkelijk aan volgens de Noppa Design Guide. Gebruik Noppa-componenten, design tokens en typography in de volledige pagina. Maak een nieuwe Noppa-versie met hero, KPI-tegels en gestylede tabellen.

## Resultaat
Nieuwe Noppa-versie gemaakt op basis van de interne Noppa Brand Guide en de Noppa Website Master Skill. Toegepast: Kentledge met webfallback, officiële kleurentokens, signature gradient, 60-30-10-opbouw, hero, KPI-componenten, Noppa-knoppen, drawer, light/dark mode, responsieve tabellen, merkfooter en directe Nederlandse tone of voice.
