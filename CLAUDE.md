# Outlook Add-in — Mail naar SharePoint

## Organisatie
JCVKESSEL Groep (~5 bedrijven, elk eigen SharePoint sites):
- **Entropal**: aluminium werk
- **Solar**: accu's en zonnepanelen
- **Bouw**: bouwprojecten (andere werkwijze, fase 2)

Alle bedrijven onderscheiden **Projecten** (offertes) en **Werken** (opdrachten). Mail begint bij een Project en verhuist naar een Werk na akkoord. Dit zijn verschillende bibliotheken in SharePoint.

## Strategie
1 add-in, meerdere flows. Detecteer organisatie van gebruiker → toon juiste flow.

- **Fase 1 (nu)**: Solar/Entropal — klikken reduceren, auto-detectie organisatie, toggle Projecten/Werken
- **Fase 2 (later)**: Bouw — coderingen in onderwerpregel (projectnr, STABI-code, leverancier), auto-detectie en opbouw via dropdowns

## Deployment
- **Code**: Push naar `main` → GitHub Actions → Azure Static Web Apps (`jolly-flower-0f8beda03.2.azurestaticapps.net`)
- **Manifest**: Admin uploadt `manifest.xml` handmatig in M365 Admin Center. Propagatie kan tot 24 uur duren.
- `dist/` staat in `.gitignore`, wordt gebouwd in CI

## Technisch
- Klassiek (legacy) desktop Outlook + Outlook web
- Outlook versie 2603, build 19822.20114 (M365, ondersteunt Mailbox 1.8+)
- Azure AD App ID: `06e23f21-f875-4425-aca3-ccd0b06bb24f`
- Tenant: `eba9b46b-0bb0-493e-8724-854a60012ad4`
- Auth: SSO met PKCE fallback (SSO-token werkt niet direct op Graph → 401 → PKCE → refresh token in localStorage)
- Taal: Nederlands (UI en communicatie)

## Auth-architectuur
- `Office.auth.getAccessToken()` geeft een bootstrap-token dat 401 geeft op Graph (geen OBO-backend)
- Bij 401: `markSsoFailedForGraph()` → volgende calls gaan direct naar PKCE
- PKCE slaat refresh token op in localStorage → daarna silent re-auth, geen popup meer
- Gedeelde mailbox: `getSharedPropertiesAsync` detecteert `targetMailbox`, doorgegeven aan alle Graph-calls als `/users/{mailboxUser}/` i.p.v. `/me/`

## Manifest
- Huidige versie: **1.0.0.4**
- `VersionOverridesV1_1` bevat `<SupportsSharedFolders>true</SupportsSharedFolders>` (vereist voor knop in gedeelde mailboxen)
- V1_1 requirements op `MinVersion="1.8"` (correct voor SupportsSharedFolders)
- **Openstaand actiepunt**: admin moet manifest 1.0.0.4 heropladen in M365 Admin Center zodat de knop ook verschijnt bij gedeelde mailboxen

## Te controleren in Azure AD
- Application ID URI moet exact zijn: `api://jolly-flower-0f8beda03.2.azurestaticapps.net/06e23f21-f875-4425-aca3-ccd0b06bb24f`
- Te vinden via: Azure AD → App Registraties → app → "Expose an API"
- Pre-autorisatie van Office-clients is niet blokkerend (PKCE fallback vangt het op), maar wel netjes om in te stellen
