# SharePoint-dashboard Externe Toegang

Een HTML-dashboard voor het inzichtelijk maken van externe toegang tot SharePoint-sites en documenten. Het dashboard gebruikt gedelegeerde Microsoft Graph-autorisatie en toont alleen gegevens die de aangemelde gebruiker op basis van diens rechten en de verleende API-scopes mag inzien.

## Bestanden

- `sharepoint-externe-toegang-dashboard.html`  
  Het interactieve dashboard.
- `README.md`  
  Installatie-, configuratie- en gebruiksinstructies.
- `referentie.md`  
  Vastlegging van de oorspronkelijke prompt, keuzes en vervolgstappen.

## Functionaliteit

### Tabblad Sites

Toont per SharePoint-site:

- sitenaam;
- site-URL;
- aantal unieke externen.

### Tabblad Externen

Toont per gevonden externe toegang:

- naam van de externe;
- e-mailadres van de externe;
- sitenaam;
- site-URL.

### Overige mogelijkheden

- interactieve aanmelding via Microsoft Authentication Library, MSAL;
- vernieuwen van de gegevens via Microsoft Graph;
- zoeken in beide tabellen;
- scanlog voor voortgang, fouten en overgeslagen onderdelen;
- export naar CSV;
- export naar JSON;
- lokale opslag van configuratie en het laatste scanresultaat in de browser;
- keuze tussen een sitescan en een uitgebreidere scan van documentbibliotheken en items;
- Noppa-vormgeving met `#F2B82C` als accentkleur.

## Vereisten

Voor gebruik zijn de volgende onderdelen nodig:

1. Een Microsoft Entra ID-tenant.
2. Rechten om een app-registratie te maken of te laten maken.
3. Een locatie waarop het HTML-bestand via HTTPS wordt gepubliceerd.
4. Een account dat toegang heeft tot de SharePoint-sites die moeten worden onderzocht.
5. Toestemming voor de benodigde gedelegeerde Microsoft Graph-rechten.

> Open het HTML-bestand bij voorkeur niet rechtstreeks via `file://`. Publiceer het via een geldige HTTPS-locatie die als redirect-URI in de app-registratie kan worden vastgelegd.

## Entra ID-app registreren

1. Open het Microsoft Entra-beheercentrum.
2. Open **App-registraties**.
3. Kies **Nieuwe registratie**.
4. Geef de app bijvoorbeeld de naam `SharePoint Externe Toegang Dashboard`.
5. Kies als ondersteund accounttype **Alleen accounts in deze organisatiemap**.
6. Registreer de toepassing.
7. Noteer:
   - de **Toepassings-id, client-id**;
   - de **Map-id, tenant-id**.
8. Open **Verificatie**.
9. Voeg een platform van het type **Single-page application** toe.
10. Voeg de exacte HTTPS-URL van het dashboard toe als redirect-URI.

Voorbeeld:

```text
https://contoso.sharepoint.com/sites/governance/SiteAssets/sharepoint-externe-toegang-dashboard.html
```

De redirect-URI in Entra ID en de redirect-URI in het dashboard moeten exact overeenkomen.

## Benodigde Microsoft Graph-rechten

Voeg onder **API-machtigingen** de volgende gedelegeerde Microsoft Graph-rechten toe.

### Minimale scopes

```text
User.Read
Sites.Read.All
```

### Aanvullende scope voor Entra-gasten

```text
User.Read.All
```

`User.Read.All` wordt gebruikt wanneer de optie voor het ophalen en verrijken van Entra-gastaccounts is ingeschakeld.

Afhankelijk van het tenantbeleid kan beheerderstoestemming nodig zijn. Laat de machtigingen beoordelen en goedkeuren volgens het beveiligings- en governancebeleid van de organisatie.

## Geen client secret gebruiken

Het dashboard is een browsertoepassing en gebruikt interactieve, gedelegeerde autorisatie. Plaats daarom geen client secret, certificaatwachtwoord of ander geheim in het HTML-bestand.

De gegevens worden namens de aangemelde gebruiker opgehaald. Daardoor wordt het uiteindelijke resultaat begrensd door:

- de SharePoint-rechten van de gebruiker;
- de scopes in het toegangstoken;
- het tenantbeleid;
- wat Microsoft Graph via de gebruikte eindpunten retourneert.

## Dashboard configureren

1. Publiceer `sharepoint-externe-toegang-dashboard.html` op de geregistreerde HTTPS-locatie.
2. Open het dashboard in de browser.
3. Kies **Instellingen**.
4. Vul de volgende velden in:
   - **Tenant ID**;
   - **Client ID**;
   - **Redirect-URI**.
5. Kies het gewenste scanniveau.
6. Stel eventueel het maximale aantal items per bibliotheek in.
7. Bepaal of Entra-gasten moeten worden opgehaald.
8. Kies **Opslaan**.

De configuratie wordt lokaal in de browser opgeslagen. Sla hierin geen geheime waarden op.

## Scanniveaus

### Site-machtigingen

Deze optie onderzoekt de door Microsoft Graph gevonden sites en vraagt de sitegerelateerde machtigingen op.

Gebruik deze stand voor:

- een snellere eerste inventarisatie;
- het controleren van direct gevonden externe toegang op siteniveau;
- tenants met veel documenten waarbij een itemscan nog niet gewenst is.

### Site, bibliotheken en items

Deze optie onderzoekt daarnaast documentbibliotheken en items, tot het ingestelde maximale aantal items per bibliotheek.

Gebruik deze stand wanneer ook losse document- of itemmachtigingen relevant zijn.

> Deze browservariant doorloopt niet gegarandeerd elke geneste map en ieder bestand in een grote tenant. Zie ook de sectie **Beperkingen**.

## Dashboard gebruiken

### Aanmelden

1. Kies **Aanmelden**.
2. Meld je aan met het account waarmee de scan moet worden uitgevoerd.
3. Accepteer alleen de machtigingen die binnen het organisatiebeleid zijn toegestaan.

### Gegevens vernieuwen

1. Kies **Vernieuwen**.
2. Volg de voortgang via de voortgangsbalk.
3. Open **Instellingen** om de scanlog te bekijken.
4. Controleer na afloop de tabbladen **Sites** en **Externen**.

### Tabellen doorzoeken

Gebruik het zoekveld boven een tabel om direct te filteren op onder andere:

- sitenaam;
- site-URL;
- naam van de externe;
- e-mailadres.

### Exporteren

Kies:

- **CSV export** voor een tabelbestand dat eenvoudig in Excel of Power BI kan worden verwerkt;
- **JSON export** voor het volledige actuele resultaat, inclusief aanvullende technische velden uit de scan.

## Betekenis van de tellingen

Bovenaan het dashboard staan vier kengetallen:

- **Sites met externen**: sites waarvoor in de scan externe toegang is gevonden;
- **Unieke externen**: unieke externe identiteiten op basis van het gevonden e-mailadres of identificatienummer;
- **Gevonden toegangstoewijzingen**: het aantal gevonden en ontdubbelde combinaties van toegang, site en resource;
- **Laatst vernieuwd**: datum en tijd van de laatste geslaagde scan.

Deze cijfers geven het resultaat van de uitgevoerde scan weer. Ze vormen niet automatisch bewijs dat de gehele tenant volledig is onderzocht.

## Volledig tenantoverzicht

Voor een tenantbreed overzicht moet de gebruiker of de gebruikte oplossing voldoende rechten hebben om:

- alle relevante SharePoint-sites te lezen;
- site- en documentmachtigingen uit te lezen;
- gastaccounts in Microsoft Entra ID te herkennen.

Een gewone site-eigenaar ziet alleen de sites en gegevens die binnen diens effectieve rechten beschikbaar zijn. Voor externe toegang via losse documenten moeten bibliotheek- en itemmachtigingen worden meegenomen. Alleen siteleden uitlezen is daarvoor niet voldoende.

## Beperkingen

Houd rekening met de volgende functionele beperkingen:

- het dashboard werkt binnen de rechten van de aangemelde gebruiker;
- een sitescan is geen volledige documentmachtigingenscan;
- de uitgebreide scan gebruikt een maximumaantal items per bibliotheek;
- geneste SharePoint- en Entra-groepen worden niet volledig uitgevouwen;
- groepslidmaatschappen kunnen indirecte toegang veroorzaken die niet als individuele externe wordt weergegeven;
- anonieme koppelingen hebben mogelijk geen herkenbare gastidentiteit;
- toegang via deelkoppelingen en overervende machtigingen kan aanvullende verwerking vereisen;
- grote tenants kunnen te maken krijgen met throttling, time-outs en lange scantijden;
- de laatste resultaten en configuratie worden lokaal in de browser opgeslagen;
- de HTML-versie bevat geen centrale historie, planning of automatische periodieke scan.

## Aanbevolen productiearchitectuur

Voor een structurele governance- of auditoplossing is een beheerde backend geschikter. Denk aan:

- een Azure Function of andere beveiligde API-laag;
- application permissions met expliciete beheerderstoestemming;
- certificaatgebaseerde authenticatie;
- centrale opslag van scans en historische snapshots;
- gecontroleerde recursieve verwerking van bibliotheken, mappen en bestanden;
- verwerking van throttling en hervatten vanaf controlepunten;
- uitvouwen van groepen en indirecte toegang;
- classificatie van directe toegang, gasttoegang, organisatiekoppelingen en anonieme koppelingen;
- rapportage in Power BI of een beveiligde beheerapplicatie.

Dit dashboard is vooral geschikt als interactieve eerste versie en als basis voor verdere ontwikkeling.

## Beveiligingsadvies

- Gebruik alleen gedelegeerde rechten die nodig zijn voor het gekozen scenario.
- Plaats nooit een client secret in HTML of JavaScript.
- Publiceer het dashboard alleen op een vertrouwde HTTPS-locatie.
- Beperk toegang tot de dashboardpagina tot bevoegde medewerkers.
- Laat aanvullende API-rechten vooraf beoordelen.
- Exporteerbestanden kunnen persoonsgegevens en toegangsgegevens bevatten. Behandel deze als vertrouwelijke informatie.
- Verwijder exports wanneer deze niet meer nodig zijn.
- Controleer periodiek de verleende app-toestemmingen en redirect-URI's.

## Problemen oplossen

### Aanmelden lukt niet

Controleer:

- of Tenant ID en Client ID correct zijn ingevuld;
- of de app als Single-page application is geregistreerd;
- of de redirect-URI exact overeenkomt;
- of pop-ups voor de dashboardlocatie zijn toegestaan;
- of de gebruiker toestemming mag geven voor de gevraagde scopes.

### Fout 401 of 403

Dit wijst meestal op ontbrekende toestemming, onvoldoende gebruikersrechten of een toegangstoken zonder de benodigde scope. Bekijk de scanlog en controleer daarna de API-machtigingen en effectieve toegangsrechten.

### Er worden geen sites gevonden

Controleer:

- of de aangemelde gebruiker toegang heeft tot SharePoint-sites;
- of `Sites.Read.All` is verleend;
- of beheerderstoestemming nodig is;
- of de sites via Microsoft Graph zichtbaar zijn voor dit account.

### Er worden geen externen gevonden

Mogelijke oorzaken:

- er is binnen de onderzochte machtigingen geen directe gastidentiteit gevonden;
- externe toegang loopt via een groep;
- alleen losse documenten zijn gedeeld terwijl uitsluitend de sitescan is gekozen;
- de machtiging is een anonieme of organisatiebrede koppeling;
- `User.Read.All` ontbreekt terwijl Entra-verrijking is ingeschakeld.

### De scan stopt bij een bibliotheek of item

Open de scanlog. Het dashboard registreert overgeslagen machtigingsaanvragen en gaat waar mogelijk verder. Verlaag bij zeer grote bibliotheken het maximale aantal items en voer opnieuw een scan uit.

## Privacy en gegevensopslag

Het dashboard bewaart de volgende informatie lokaal in de browser:

- Tenant ID;
- Client ID;
- redirect-URI;
- gekozen scanniveau;
- itemlimiet;
- het laatste scanresultaat.

Het toegangstoken wordt door MSAL in de sessieopslag beheerd. Sluit de browsersessie na gebruik op een gedeeld apparaat en wis browsergegevens wanneer dat volgens het organisatiebeleid nodig is.

## Versiebeheer

Aanbevolen wijzigingen om vast te leggen:

- versienummer van het dashboard;
- datum van wijziging;
- gebruikte scopes;
- gewijzigde Graph-eindpunten;
- aanpassingen in herkenning van gasten;
- wijzigingen in exportvelden;
- beveiligingsreview en goedkeuring.

## Samenvatting

De oplossing biedt een praktisch en doorzoekbaar overzicht van gevonden externe SharePoint-toegang. De knop **Vernieuwen** voert een actuele scan uit namens de ingelogde gebruiker. Voor een aantoonbaar volledig tenantoverzicht is een centraal beheerde scanoplossing met passende rechten, recursieve verwerking en historische opslag de aanbevolen vervolgstap.
