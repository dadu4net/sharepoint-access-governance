# Noppa SharePoint & Teams Access Governance

Dit is een Client-Side Single Page Application (SPA), gebouwd met Next.js en React, ontworpen om direct vanuit een (SharePoint) omgeving of lokaal te draaien zonder afhankelijkheid van een Node.js server. 

De applicatie verzorgt een 4-niveau diepe autorisatie-audit (Site -> Bibliotheek -> Map -> Item) en biedt een tenant-brede scanner om alle externe gasten en actieve deellinks op te sporen in je Microsoft 365 omgeving.

## Vereisten
- [Node.js](https://nodejs.org/) (minimaal versie 18.x)
- Een Microsoft Entra ID (voorheen Azure AD) tenant met admin-rechten om een **App Registratie** aan te maken.

## 1. Lokale Installatie & Ontwikkeling

Kloon of download de repository en installeer de afhankelijkheden:

```bash
# Ga naar de map
cd sharepoint-access-governance

# Installeer dependencies
npm install

# Start de lokale ontwikkelserver
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in je browser om de applicatie te bekijken. 

## 2. Microsoft Entra ID App Registratie (MSAL Configuratie)

Omdat de app draait als een 100% Client-Side SPA, heb je een App Registratie nodig in Entra ID (Azure AD) die geconfigureerd is met de **Implicit Grant / PKCE** flow.

1. Ga naar het [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Ga naar **App-registraties** en kies **Nieuwe registratie**.
3. Geef de app een naam (bijv. "Noppa Governance Dashboard").
4. Kies voor **Accounts in een organisatiemap (Elke Microsoft Entra-directory - Multitenant)** of een enkele tenant.
5. Bij **Omleidings-URI (Redirect URI)** kies je **Single-page application (SPA)** en vul je de URL in waar de app gehost wordt (voor lokale ontwikkeling is dit `http://localhost:3000`).
6. Klik op **Registreren**.
7. Zodra aangemaakt, noteer je:
   - Het **Toepassings-ID (Client ID)**
   - Het **Map-ID (Tenant ID)**
8. Ga naar het tabblad **API-machtigingen (API Permissions)** en voeg de volgende *Gedelegeerde* (Delegated) Graph API permissies toe (Admin Consent is sterk aanbevolen):
   - `User.Read`
   - `Sites.Read.All`
   - `Directory.Read.All` (Voor het scannen van gastaccounts in de tenant)

Wanneer je de webapplicatie opent, kun je via het "Tenant Instellingen" modal of direct in het inlogvenster je Client ID en Tenant ID invullen om te authentiseren. Deze worden veilig lokaal opgeslagen in de browser (`localStorage`).

## 3. Bouwen voor Productie (Static Export)

Om de applicatie te exporteren naar een set statische HTML, CSS en JavaScript bestanden die in elke willekeurige map (bijv. in SharePoint SiteAssets) kunnen draaien:

```bash
npm run build
```

Door de configuratie (`output: 'export'` en `assetPrefix: '.'`) genereert Next.js een `out/` map waarbij de scripts en CSS-bestanden relatief worden benaderd.

## 4. Implementatie (Deployment naar SharePoint)

Omdat de applicatie louter bestaat uit front-end code (HTML, CSS, JS) en direct vanuit de browser met Microsoft Graph communiceert, kan deze eenvoudig in SharePoint gehost worden.

1. Genereer de statische bestanden met `npm run build`.
2. Open je SharePoint site (bijv. het admin of IT portaal).
3. Navigeer naar **Inhoud van site (Site Contents)** -> **Site-activa (SiteAssets)**.
4. Maak een nieuwe map aan (bijv. "GovernanceApp").
5. Upload de **volledige inhoud** van de `out/` map (inclusief `index.html` en de `_next/` map) in deze nieuwe map.
6. Voeg de absolute URL van deze SharePoint map toe als **Redirect URI** (SPA) in je Entra ID App Registratie.
7. Gebruikers met de link kunnen nu het `index.html` bestand openen vanuit SharePoint, waarna de app volledig zal werken. 
   *(Let op: Je kunt de app ook insluiten als een SharePoint pagina web-onderdeel door gebruik te maken van een iframe / code-snippet die naar de HTML wijst)*.

## Design Systeem (Noppa UX/UI)

Dit project is gebouwd onder de Noppa merkrichtlijnen, met sterke integratie van:
* Kentledge Typografie
* 60-30-10 Kleur theorie (`cyan`, `royal`, `card` achtergronden)
* Apple Human Interface (frosted glass blur effecten en verfijnde box-shadows)
* Tabulaire cijfers in KPI-statistieken voor de overzichtelijkheid. 
* Tailwind CSS als basis (geconfigureerd in `tailwind.config.js` en `src/styles/noppa.css`).
