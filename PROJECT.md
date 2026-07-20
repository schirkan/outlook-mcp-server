# Outlook MCP Server

MCP-Server (Model Context Protocol) fuer den Zugriff auf eine lokale **klassische** Outlook-Installation (Microsoft Outlook fuer Windows, klassischer Desktop-Client - kein Microsoft 365 Cloud, kein Graph API, kein "New Outlook").

## Ziele

- **E-Mails lesen** (Ordner auflisten, einzelne Mails, Suche, Anhaenge auflisten)
- **E-Mails schreiben** (neue Mail, Antworten, Allen antworten, Weiterleiten, Entwurf speichern)
- **Kalender lesen** (Kalender auflisten, Termine abfragen, freie Zeitfenster suchen)
- **Kalender schreiben** (Termin erstellen, aktualisieren, loeschen, Teilnehmer einladen)
- Zugriff vollstaendig **lokal** via `Microsoft.Office.Interop.Outlook` (COM-Interop)
- Bereitstellung als MCP-Server (`modelcontextprotocol/csharp-sdk`)
- API-Design semantisch nahe an **Microsoft Graph Mail/Calendar API**, technisch unabhaengig

## Nicht-Ziele (v1)

- Cloud-Variante (Microsoft 365, Outlook.com, Exchange Online via Graph)
- Microsoft Graph API als Backend
- Moderner "New Outlook" / One Outlook / Outlook (Preview)
- **Kontakte, Aufgaben, Notizen, Regeln** (Follow-up-Phase, nicht v1)
- Mobile- oder Web-Frontend
- Multi-User / Multi-Profil (Outlook laeuft im aktiven Windows-Benutzerprofil)
- Linux/macOS (Windows-only wegen COM-Interop)

## Constraints

- Outlook-Version: Klassischer Desktop-Client (Outlook 2016 / 2019 / 2021 LTSC/Retail, alle On-Prem- oder Hybrid-Setups mit klassischem MAPI-Profil)
- Anbindung: `Microsoft.Office.Interop.Outlook` (COM-Addin-frei, late-bound wo noetig)
- Authentifizierung: kein eigener Credential-Store - interop nutzt das aktive Outlook-Profil des angemeldeten Windows-Benutzers
- Sprache: **C#** / **.NET 8+**
- SDK: `modelcontextprotocol/csharp-sdk`
- Plattform: **Windows** (x64)

## Architektur (Uebersicht)

```
+-----------------------------+
|   MCP-Client (Claude, IDE,  |
|   eigenes Tool, etc.)       |
+-------------||--------------+
              ||  MCP / JSON-RPC / stdio (oder HTTP/SSE)
              ||
+-------------vv--------------+
|   OutlookMcpServer (C#)     |
|   - MailTools               |
|   - CalendarTools           |
+-------------||--------------+
              || C# / COM-Interop
+-------------vv--------------+
|   Microsoft.Office.         |
|   Interop.Outlook           |
+-------------||--------------+
              || MAPI / RPC
+-------------vv--------------+
|   Outlook (klassisch,       |
|   laufendes Profil)         |
+-----------------------------+
```

Detail-Architektur: siehe [`specs/ARCHITECTURE.md`](specs/ARCHITECTURE.md).
API-Design: siehe [`specs/API-DESIGN.md`](specs/API-DESIGN.md).
Vision & Scope: siehe [`specs/VISION.md`](specs/VISION.md).

## Current Status

- [x] Projekt-Ordner angelegt
- [x] GitHub-Repo eingerichtet (`schirkan/outlook-mcp-server`, public)
- [x] `.gitignore` angelegt
- [x] Initial `PROJECT.md` + `README.md`
- [x] Workboard eingerichtet (Board-ID `outlook-mcp-server`, 10 Karten)
- [x] Research abgeschlossen (Graph API Mail+Calendar, Interop Mail+Calendar, MCP C# SDK)
- [x] Specs geschrieben (`specs/VISION.md`, `specs/ARCHITECTURE.md`, `specs/API-DESIGN.md`)
- [x] Solution-Scaffold (.NET 8 + MCP SDK + Outlook Interop, `OutlookMcpServer.sln`)
- [x] Domain-Schicht (DTOs, `IOutlookService`, `OutlookService` mit Validierung)
- [x] Interop-Adapter Grundgerüst (COM-Boundary Mail+Calendar, `EnumMappings`, 24 Stubs)
- [x] Konfiguration + DI + Transport (stdio, `appsettings.json`)
- [x] MCP-Tools (MailTools + CalendarTools + ActiveSelectionTools = 27 Tools)
- [x] Unit-Tests xUnit (**47/47 grün**, `OutlookMcpServer.Domain.Tests`) — 35 bestehend + 12 Phase-3h-Tests (Polymorphie / Scope-Filter / Top-Cap / Empty-Selection / Validation / Exception-Propagation)
- [x] **Echte COM-Interop-Impl** (Karte 3.5) — alle 26 Methoden implementiert (24 Mail/Calendar + 2 Active-Selection), 0 verbleibende `NotImplementedException`
- [x] **Publish-Profil `minimal.pubxml`** — Self-Contained + Single-File + Trim(partial) + Compression + Embedded-PDB + InvariantGlobalization → 17,7 MB exe
- [x] **Integration-Tests Projekt-Skelett** (Karte 7) — `tests/OutlookMcpServer.IntegrationTests/` mit SkippableFact + Outlook-DetectOutlook (5s-Task-Wait-Timeout gegen COM-Haenger) + 6 Beispiel-Tests (3 MailFolder + 3 Calendar). Vollständige Test-Suite (send/create/respond/delete mit Cleanup) + lokale Verifikation durch Martin stehen aus.
- [x] **README erweitert** (Karte 8) — Features v1 / Quick Start (Build/Test/Publish) / Configuration (appsettings.json + Allow*-Flags-Tabelle) / Transport (stdio + HTTP/SSE-Loopback) / MCP-Client-Setup (Claude Desktop, Cline, Continue.dev) / Architecture (3-Layer + ASCII-Diagramm) / Development (Project-Structure + Add-a-new-MCP-Tool-Workflow) / Roadmap / License
- [x] **Beispiel-Configs `examples/`** (Karte 9) — separate JSON-Files für Copy-Paste (`claude-desktop-config.json`, `cline-mcp-settings.json`, `appsettings.http.json`, `appsettings.readonly.json` + `examples/README.md`)
- [x] **CI/CD-Pipeline** (GitHub Actions) — `ci.yml` (Build + Test bei push/PR auf main) + `release.yml` (Tag-Triggered Publish + GitHub Release mit auto-generierten Notes). Windows-Runner wegen COM-Interop.
- [ ] Manuelle COM-Verifikation gegen echtes Outlook-Profil (Martin) — Karte 7 Acceptance `lokal gruen`

## Git

- **Repo-Typ:** GitHub (public)
- **Pfad / URL:** https://github.com/schirkan/outlook-mcp-server
- **Remote(s):** `origin` -> https://github.com/schirkan/outlook-mcp-server.git
- **Eingerichtet am:** 2026-07-15
- **Standard-Branch:** `main`
- **`.gitignore`-Status:** vorhanden

## CI / CD (GitHub Actions)

Zwei separate Workflows in `.github/workflows/`:

- **CI** (`.github/workflows/ci.yml`, Commit `e727a69`)
  - Trigger: `push` + `pull_request` auf `main`
  - Runner: `windows-latest` (COM-Interop = Windows-only)
  - .NET SDK: aus `global.json` (8.0.x, rollForward latestFeature)
  - Steps: checkout → setup-dotnet (NuGet-Cache) → restore → build (Release) → test Domain.Tests → test IntegrationTests
  - IntegrationTests skippen sauber via `OutlookIntegrationTestBase.DetectOutlook` (5s-Task-Wait-Timeout) wenn kein Outlook-Profil vorhanden
  - Timeout 20 Min, `concurrency.cancel-in-progress` (spart CI-Minuten)
  - Publish laeuft NICHT in CI (zu teuer, separate Datei)

- **Release** (`.github/workflows/release.yml`)
  - Trigger: `push` auf Tag-Match `v[0-9]+.[0-9]+.[0-9]+` (z. B. `v1.0.0`, `v1.0.1`, `v2.0.0`)
  - Pre-Release-Tags (z. B. `v1.0.0-rc.1`, `v1.0.0-beta.2`) werden via `prerelease`-Expression auto-markiert
  - Runner: `windows-latest`, .NET SDK aus `global.json`
  - Steps: checkout (full history) → setup-dotnet → restore → publish mit `minimal.pubxml` (Self-Contained + Single-File + Trim partial + Compression) → Zip → Upload-Artifact → GitHub-Release (mit auto-generierten Notes aus Commits seit letztem Tag)
  - Permissions: `contents: write` (fuer `gh release create`)
  - Artifact: `outlook-mcp-server-${{ github.ref_name }}.zip` (OutlookMcpServer.exe ~17.7 MB + pdb + aspnetcorev2_inprocess.dll)

- **Workflow manuell triggern** (zum Testen ohne Push):
  - GitHub UI: `Actions`-Tab → Workflow auswaehlen → `Run workflow`
  - CLI: `gh workflow run ci.yml` (mit `-F key=value` fuer Inputs, falls vorhanden)

- **Release-Prozess** (lokal):
  ```bash
  # Tag erstellen (SemVer + aussagekraeftige Message)
  git tag -a v1.0.0 -m "v1.0.0 — erste stabile Version"

  # Tag pushen — loest release.yml aus
  git push origin v1.0.0

  # Status checken
  gh release list
  ```

## Workboard

- **Board-ID:** `outlook-mcp-server`
- **Stats:** 10 Karten total · 10 done · 0 in_progress · 0 backlog
- **Done:**
  - `d8753677-91bc-4181-9e39-4c5139d12990` — Doku: Beispiel-Config + MCP-Client-Setup (Claude Desktop, Cline) (low) — README enthält inline-Configs + `examples/` mit separaten JSON-Files für Copy-Paste
  - `f78b75ed-6f77-439c-abd2-7b03a1d9f371` — Impl: Echte COM-Interop (26 Methoden + Active-Selection) (high)
  - `022e0b4e-f07a-499a-904d-4c4a49443871` — Tests: Integration (Outlook-Profil, xUnit, COM-Adapter) — Skeleton + 6 Beispiel-Tests done, vollständige Test-Suite (send/create/respond/delete mit Cleanup) + manuelle Verifikation durch Martin stehen aus
  - `737dbaa1-f169-4094-af81-a6204ece9052` — Doku: README erweitert (Build, Konfiguration, Verwendung)
  - Solution-Scaffold · Domain-Schicht · Interop-Adapter Grundgerüst · Konfiguration + Transport · MCP-Tools · Unit-Tests (**47/47 grün**)

## Project Files

- `README.md` - Landing Page (kurzer Ueberblick, Links)
- `PROJECT.md` - diese Datei (Status, Constraints, Architektur-Uebersicht)
- `specs/VISION.md` - Vision, Scope, Ziele/Nicht-Ziele (Detail)
- `specs/ARCHITECTURE.md` - Schichten, Komponenten, Datenfluesse
- `specs/API-DESIGN.md` - MCP-Tools & Resources fuer Mail + Kalender
- `DECISIONS.md` - Architektur- und Designentscheidungen mit Datum + Begruendung
