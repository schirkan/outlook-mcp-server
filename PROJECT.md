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
- [x] MCP-Tools (MailTools 16 + CalendarTools 9 + ActiveSelectionTools 2 = **27 Tools** — inkl. Martins `list_mails_recursive` (Karte 4149bc2) + Bulk-Get `get_mails`)
- [x] Unit-Tests xUnit (**47/47 grün**, `OutlookMcpServer.Domain.Tests`) — 35 bestehend + 12 Phase-3h-Tests (Polymorphie / Scope-Filter / Top-Cap / Empty-Selection / Validation / Exception-Propagation)
- [x] **Echte COM-Interop-Impl** (Karte 3.5) — alle 26 Methoden implementiert (24 Mail/Calendar + 2 Active-Selection), 0 verbleibende `NotImplementedException`
- [x] **Bulk-Get `get_mails`** (Bulk-Variante von `get_mail`) — neues MCP-Tool liest 1-50 EntryIDs in einem Aufruf, liefert `{ value: [...], notFoundIds: [...] }` (Bulk-Semantik, einzelne ungültige IDs landen in `notFoundIds`). 11 neue Tests (5 Tool + 6 Service-Validation: null/empty/too-many/empty-id/dedup/happy-path). Code-Schicht: Model + Interop-Adapter (COM-Loop mit per-ID-try/catch) + Service (Validation 1-50 + Dedup) + Tool (CSV-frei, JSON-Array-IDs) + Fakes. Dokumentiert in `specs/API-DESIGN.md` (`### getMails`) + Mapping-Tabelle. Production-Default `includeBody=false` weil Body-Inhalt oft groß.
- [x] **CI/CD-Pipeline Hardening + Push-Force-Pattern** — Martins Commits 6b053ed (CallSite-Pattern), 4149bc2 (list_mails_recursive), d8ea89e (PublishTrimmed=false), 22479dc (Logging) waren zwischen meinem letzten Push und dieser Session auf origin/main gelandet. Mein Bulk-Get-Commit wurde via `git rebase origin/main` integriert (9/10 Files auto-merged, einziger Konflikt `tests/.../OutlookServiceTests.cs` in der Test-Klasse gelöst via `git checkout --theirs` + manuellem Anhängen von Martins 5 ListMailsRecursive-Tests vor meinen 6 GetMails-Tests). Rebase-Commit `872ebfb` mit korrigierter `--force-with-lease`-Push. Backup-Branch `backup-get-mails` (Original `feb387b`) bleibt bis zur Bestätigung als Rollback-Safety erhalten.
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

- **CI** (`.github/workflows/ci.yml`, Commits `e727a69` + `5bafc3e` fuer Cache-Fix)
  - Trigger: `push` + `pull_request` auf `main`
  - Runner: `windows-latest` (COM-Interop = Windows-only)
  - .NET SDK: aus `global.json` (8.0.x, rollForward latestFeature)
  - Steps: checkout → setup-dotnet (**kein** NuGet-Cache, siehe "Cache-Status" unten) → restore → build (Release) → test Domain.Tests → test IntegrationTests
  - IntegrationTests skippen sauber via `OutlookIntegrationTestBase.DetectOutlook` (5s-Task-Wait-Timeout) wenn kein Outlook-Profil vorhanden
  - Timeout 20 Min, `concurrency.cancel-in-progress` (spart CI-Minuten)
  - Publish laeuft NICHT in CI (zu teuer, separate Datei)

- **Release** (`.github/workflows/release.yml`, Commits `6e69e56` + `5bafc3e` fuer Cache-Fix)
  - Trigger: `push` auf Tag-Match `v[0-9]+.[0-9]+.[0-9]+` (z. B. `v1.0.0`, `v1.0.1`, `v2.0.0`)
  - Pre-Release-Tags (z. B. `v1.0.0-rc.1`, `v1.0.0-beta.2`) werden via `prerelease`-Expression auto-markiert
  - Runner: `windows-latest`, .NET SDK aus `global.json`
  - Steps: checkout (full history) → setup-dotnet (**kein** NuGet-Cache, siehe "Cache-Status" unten) → restore → publish mit `minimal.pubxml` (Self-Contained + Single-File + Trim partial + Compression) → Zip → Upload-Artifact → GitHub-Release (mit auto-generierten Notes aus Commits seit letztem Tag)
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

- **Pipeline-Status (Live, nach erstem v0.1.0 Release):**
  - Erste gepublishte GitHub-Release: **[v0.1.0](https://github.com/schirkan/outlook-mcp-server/releases/tag/v0.1.0)** — ausgeloest durch `git push origin v0.1.0` (Tag-Trigger fuer `release.yml`)
  - Asset: `outlook-mcp-server.zip` (OutlookMcpServer.exe ~17,7 MB Self-Contained Single-File + pdb + aspnetcorev2_inprocess.dll)
  - Release-Run-Dauer (Run [29724649418](https://github.com/schirkan/outlook-mcp-server/actions/runs/29724649418)): **~116s (1,9 min)** auf windows-latest
  - CI-Laufzeit (nach Cache-Fix 5bafc3e): **~68s** auf windows-latest
  - Tag-zu-Release-Zeit gesamt: ~2 min (Tag-Push → GitHub Release published + Asset verfuegbar)
  - Auto-Update fuer MCP-Clients (Claude Desktop, Cline, Continue.dev, Cursor): aktuell **manuell** — Download pro Release aus GitHub, exe ersetzen. Auto-Update-Mechanik (GitHub-Releases-API pollen, SemVer-Vergleich, Auto-Replace) ist kein v1-Scope, Folge-Schritt.

- **Cache-Status (Bekannte Limitierung):**
  - Beide Workflows (`ci.yml` + `release.yml`) nutzen **kein** `cache: true` auf `actions/setup-dotnet@v4`. Grund: die Action verlangt eine `packages.lock.json`, die unser Repo (noch) nicht hat — mit `cache: true` wirft die Action "Dependencies lock file is not found", Restore/Build/Tests werden gar nicht erreicht (Fruehausstieg nach ~45s).
  - **Workaround (Commit 5bafc3e):** `cache: true` entfernt, Inline-Kommentar in beiden Workflow-Files erklaert das Trade-off. NuGet-Restore ohne Cache ist ~10-20s langsamer pro CI-/Release-Lauf, fuer unsere Projektgroesse vernachlaessigbar.
  - **Saubere Loesung (Roadmap):** Central Package Management (CPM) via `Directory.Packages.props` aktivieren + `<RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>` in jedem csproj setzen + `dotnet restore --use-lock-file` einmal lokal ausfuehren → generiert `packages.lock.json` pro Projekt. Danach kann `cache: true` in beiden Workflows wieder aktiviert werden, ohne die Cache-Key-Findungs-Failure. Folge-Schritt, kein Dringlichkeits-Bedarf.

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
