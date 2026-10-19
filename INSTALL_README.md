# Noppa SharePoint & Teams Access Governance — Windows Executable (.exe) Installatie & Gebruikershandleiding

Deze handleiding beschrijft de installatie, configuratie en het gebruik van de Windows Desktop Executable (`Noppa-SharePoint-Governance.exe`) voor het Noppa SharePoint & Teams Access Governance portaal.

---

## 1. Wat is de Desktop Executable?

De desktop executable is een zelfstandige Windows-applicatie (.exe) gebaseerd op .NET 8 en Microsoft Edge WebView2:

* **Geen Node.js of lokale server handmatig starten:** De applicatie bevat een ingebouwde micro-webserver (`http://localhost:3334/`) die de Noppa SPA razendsnel in het geheugen laadt.
* **Native Microsoft 365 Authenticatie (Entra ID):** Wanneer je op "Inloggen M365" klikt, opent er een afgeschermd native pop-up venster waarin de gebruiker veilig met MFA kan inloggen. Zodra het token is verkregen, sluit het pop-up venster zichzelf volautomatisch.
* **Blijvende Sessies & Instellingen:** De browserdata, Microsoft login-sessie en Tenant-configuratie worden veilig opgeslagen in `%LOCALAPPDATA%\NoppaGovernance\WebView2Data`, zodat je niet bij elke herstart opnieuw hoeft in te loggen.
* **Volledig Noppa Design System:** Bevat alle Noppa huisstijlcomponenten (60-30-10 kleurbalans, Kentledge typografie, Apple frosted-glass blur en responsieve tabellen).

---

## 2. Beschikbare Executables in `dist/`

In de map `dist/` zijn twee kant-en-klare varianten gegenereerd:

| Bestand | Grootte | Beschrijving | Wanneer te gebruiken? |
| :--- | :--- | :--- | :--- |
| **`dist\Noppa-SharePoint-Governance.exe`** | **~2.1 MB** | **Aanbevolen (Lichtgewicht)**. Maakt gebruik van de op het systeem aanwezige .NET 8 runtime. | Ideaal voor IT-beheerders en werkplekken waar .NET 8 al geïnstalleerd is. |
| **`dist\standalone\Noppa-SharePoint-Governance.exe`** | **~164 MB** | **100% Self-Contained**. Bevat de complete .NET runtime ingebouwd. | Draait gegarandeerd op élke Windows 10/11 PC, zónder dat er voorwaarden of runtimes vereist zijn. |

---

## 3. Installatie Opties

### Optie A: Direct Draaien (Portable / Zonder installatie)
Je kunt `Noppa-SharePoint-Governance.exe` simpelweg dubbelklikken vanaf elke locatie (bijvoorbeeld je Bureaublad of een netwerkschijf). De applicatie start onmiddellijk op.

---

### Optie B: Automatische Installatie naar Startmenu & Bureaublad (Aanbevolen)
Er is een geautomatiseerd PowerShell installatiescript aanwezig (`install.ps1`). Dit script vereist **geen** administratorrechten:

1. Open PowerShell in de projectmap.
2. Voer het installatiescript uit:
   ```powershell
   .\install.ps1
   ```
   *(Wil je direct de 100% self-contained variant installeren en meteen starten? Gebruik dan: `.\install.ps1 -Standalone -LaunchAfterInstall`)*

**Wat doet het installatiescript?**
* Installeert de bestanden netjes naar `%LOCALAPPDATA%\Programs\NoppaGovernance`.
* Maakt een snelkoppeling aan in het **Windows Startmenu** onder `Noppa Solutions`.
* Plaatst een snelkoppeling op je **Bureaublad** met het officiële Noppa logo-icoon.
* Genereert een lokaal de-installatiescript.

---

### Optie C: De-installeren
Wil je de applicatie en alle bijbehorende snelkoppelingen verwijderen? Voer simpelweg uit:
```powershell
.\uninstall.ps1
```

---

## 4. Microsoft Entra ID (Azure AD) App Registratie Instellen

Omdat de desktop-app intern communiceert via een lokale beveiligde lus (`http://localhost:3334/`), moet de Microsoft Entra App Registratie als volgt worden ingesteld:

1. Ga naar het [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Navigeer naar **Identity (Identiteit)** > **Applications (Toepassingen)** > **App registrations (App-registraties)**.
3. Selecteer je bestaande registratie of klik op **New registration (Nieuwe registratie)**.
4. Ga naar het tabblad **Authentication (Verificatie)**:
   * Klik op **Add a platform (Platform toevoegen)** > kies **Single-page application (SPA)**.
   * Voeg de volgende **Redirect URI's** toe:
     * `http://localhost:3334/`
     * `http://localhost:3334`
5. Ga naar het tabblad **API permissions (API-machtigingen)** en voeg de volgende **Gedelegeerde (Delegated)** Microsoft Graph machtigingen toe:
   * `User.Read` (Profiel van de ingelogde gebruiker lezen)
   * `Sites.Read.All` (SharePoint & Teams sites, bibliotheken en unieke rechten analyseren)
   * `Directory.Read.All` (Lijst van externe gasten en contactpersonen in de tenant scannen)
   * *Tip: Laat een globale beheerder op **Grant admin consent (Beheerdersgoedkeuring verlenen)** klikken.*
6. Open de Noppa Governance App op je PC, klik rechtsboven op het **tandwiel-icoon (Instellingen)** en vul in:
   * **Tenant ID (Map-ID)**
   * **Client ID (Toepassings-ID)**
   * Klik op **Opslaan**. Deze instellingen blijven automatisch bewaard.

---

## 5. Ontwikkelaars: Zelf een Nieuwe .exe Genereren

Heb je wijzigingen aangebracht in de React/Next.js componenten of de CSS? Je kunt met één enkel commando een gloednieuwe .exe genereren:

```bash
# Compileert de web frontend en genereert de lichtgewicht .exe in dist/
npm run build:exe

# OF: genereer de 100% self-contained variant
npm run build:exe:standalone
```

Het commando voert automatisch de volgende stappen uit:
1. `next build` (genereert de statische SPA export in `out/`).
2. Comprimeert `out/` naar `desktop/wwwroot.zip`.
3. Compileert het .NET 8 project met embedded resources en WebView2 naar `dist/Noppa-SharePoint-Governance.exe`.
