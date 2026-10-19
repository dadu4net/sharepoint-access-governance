# Noppa SharePoint & Teams Access Governance

Dit project biedt een gespecialiseerde 4-niveau diepe autorisatie-audit (Site -> Bibliotheek -> Map -> Item) en een tenant-brede externe toegangsverkenner voor Microsoft SharePoint & Teams.

De applicatie is beschikbaar als:
1. **Windows Desktop Executable (`.exe`)**: Zelfstandige desktop app voor Windows 10/11 met ingebouwde micro-server en native MSAL authenticatie.
2. **Statische Web SPA (HTML/JS/CSS)**: Direct te hosten in een SharePoint `SiteAssets` bibliotheek of op een webserver.

---

## ⚡ Snelstart: Windows Executable (.exe)

Voor gedetailleerde instructies, zie de complete gids: [INSTALL_README.md](file:///c:/DEV/sharepoint-access-governance/INSTALL_README.md).

### Direct starten of installeren:
In de map `dist/` staan twee kant-en-klare Windows executables:
* **`dist\Noppa-SharePoint-Governance.exe`** (~2.1 MB) — Lichtgewicht, vereist .NET 8 op het systeem.
* **`dist\standalone\Noppa-SharePoint-Governance.exe`** (~164 MB) — 100% self-contained, werkt op elke Windows PC zonder voorwaarden.

### Automatische Installatie (Startmenu & Bureaublad):
Voer het meegeleverde PowerShell-script uit (geen admin-rechten nodig):
```powershell
.\install.ps1
```
Dit kopieert de applicatie naar `%LOCALAPPDATA%\Programs\NoppaGovernance` en maakt snelkoppelingen aan in het Startmenu en op je Bureaublad met het Noppa logo.

---

## 1. Lokale Ontwikkeling (Web SPA)

### Vereisten
- [Node.js](https://nodejs.org/) (minimaal v18.x)
- [.NET 8 SDK](https://dotnet.microsoft.com/) (alleen nodig bij het bouwen van de `.exe`)

```bash
# Installeer web dependencies
npm install

# Start de lokale ontwikkelserver
npm run dev
```

Open [http://localhost:3334](http://localhost:3334) in je browser om de applicatie te bekijken.

---

## 2. Microsoft Entra ID (Azure AD) App Registratie

De applicatie maakt gebruik van de Microsoft Authentication Library (MSAL) met de PKCE/SPA-flow:

1. Ga naar het [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Ga naar **App-registraties** en kies **Nieuwe registratie**.
3. Kies bij **Omleidings-URI (Redirect URI)** voor **Single-page application (SPA)**:
   * Voor de Windows Desktop App: `http://localhost:3334/` en `http://localhost:3334`
   * Voor SharePoint hosting: bijv. `https://<tenant>.sharepoint.com/sites/<site>/SiteAssets/GovernanceApp/`
4. Voeg onder **API-machtigingen (API Permissions)** de volgende *Gedelegeerde* (Delegated) Graph API-rechten toe:
   * `User.Read`
   * `Sites.Read.All`
   * `Directory.Read.All` (voor de tenant gasten audit)
5. Klik op het tandwiel-icoon (**Instellingen**) in de applicatie om je **Tenant ID** en **Client ID** op te slaan.

---

## 3. Bouwen voor Productie

```bash
# 1. Alleen de statische web export bouwen (genereert out/):
npm run build

# 2. Complete Windows Desktop Executable bouwen (.exe in dist/):
npm run build:exe

# 3. Complete Standalone Executable bouwen (.exe in dist/standalone/):
npm run build:exe:standalone
```

---

## 4. Implementatie (Deployment naar SharePoint)

1. Genereer de statische bestanden met `npm run build`.
2. Open je SharePoint site (bijv. het admin of IT portaal).
3. Navigeer naar **Inhoud van site (Site Contents)** -> **Site-activa (SiteAssets)**.
4. Maak een nieuwe map aan (bijv. "GovernanceApp").
5. Upload de **volledige inhoud** van de `out/` map in deze nieuwe map.
6. Voeg de URL toe als **Redirect URI** (SPA) in je Entra ID App Registratie.

---

## 5. Design Systeem (Noppa UX/UI)

Dit project is gebouwd onder de officiële Noppa merkrichtlijnen:
* **Kentledge Typografie**
* **60-30-10 Kleurbalans** (`cyan`, `royal`, `card` achtergronden)
* **Apple Human Interface** (frosted glass blur effecten en verfijnde box-shadows)
* **Tabulaire cijfers** in KPI-statistieken voor maximale scanbaarheid
* **Tailwind CSS** als fundament (`tailwind.config.js` en `src/styles/noppa.css`).
