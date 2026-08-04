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
- [x] MCP-Tools (MailTools 16 + CalendarTools 9 + ActiveSelectionTools 3 = **28 Tools** — inkl. Martins `list_mails_recursive` (Karte 4149bc2) + Bulk-Get `get_mails` + `open_item_by_id` (Inspector-Öffnung, 2026-07-28))
- [x] Unit-Tests xUnit (**47/47 grün**, `OutlookMcpServer.Domain.Tests`) — 35 bestehend + 12 Phase-3h-Tests (Polymorphie / Scope-Filter / Top-Cap / Empty-Selection / Validation / Exception-Propagation)
- [x] **Echte COM-Interop-Impl** (Karte 3.5) — alle 26 Methoden implementiert (24 Mail/Calendar + 2 Active-Selection), 0 verbleibende `NotImplementedException`
- [x] **Bulk-Get `get_mails`** (Bulk-Variante von `get_mail`) — neues MCP-Tool liest 1-50 EntryIDs in einem Aufruf, liefert `{ value: [...], notFoundIds: [...] }` (Bulk-Semantik, einzelne ungültige IDs landen in `notFoundIds`). 11 neue Tests (5 Tool + 6 Service-Validation: null/empty/too-many/empty-id/dedup/happy-path). Code-Schicht: Model + Interop-Adapter (COM-Loop mit per-ID-try/catch) + Service (Validation 1-50 + Dedup) + Tool (CSV-frei, JSON-Array-IDs) + Fakes. Dokumentiert in `specs/API-DESIGN.md` (`### getMails`) + Mapping-Tabelle. Production-Default `includeBody=false` weil Body-Inhalt oft groß.
- [x] **CI/CD-Pipeline Hardening + Push-Force-Pattern** — Martins Commits 6b053ed (CallSite-Pattern), 4149bc2 (list_mails_recursive), d8ea89e (PublishTrimmed=false), 22479dc (Logging) waren zwischen meinem letzten Push und dieser Session auf origin/main gelandet. Mein Bulk-Get-Commit wurde via `git rebase origin/main` integriert (9/10 Files auto-merged, einziger Konflikt `tests/.../OutlookServiceTests.cs` in der Test-Klasse gelöst via `git checkout --theirs` + manuellem Anhängen von Martins 5 ListMailsRecursive-Tests vor meinen 6 GetMails-Tests). Rebase-Commit `872ebfb` mit korrigierter `--force-with-lease`-Push. Backup-Branch `backup-get-mails` (Original `feb387b`) bleibt bis zur Bestätigung als Rollback-Safety erhalten.
- [x] **Publish-Profil `minimal.pubxml`** — Self-Contained + Single-File + Compression + Embedded-PDB + InvariantGlobalization → **44,94 MB exe** (PublishTrimmed=false per `d8ea89e` — `BuiltInComInterop.IsSupported=false` unter Trimmed blockiert sonst COM-Aktivierung)
- [x] **Integration-Tests Projekt-Skelett** (Karte 7) — `tests/OutlookMcpServer.IntegrationTests/` mit SkippableFact + Outlook-DetectOutlook (5s-Task-Wait-Timeout gegen COM-Haenger) + 6 Beispiel-Tests (3 MailFolder + 3 Calendar). Vollständige Test-Suite (send/create/respond/delete mit Cleanup) + lokale Verifikation durch Martin stehen aus.
- [x] **README erweitert** (Karte 8) — Features v1 / Quick Start (Build/Test/Publish) / Configuration (appsettings.json + Allow*-Flags-Tabelle) / Transport (stdio + HTTP/SSE-Loopback) / MCP-Client-Setup (Claude Desktop, Cline, Continue.dev) / Architecture (3-Layer + ASCII-Diagramm) / Development (Project-Structure + Add-a-new-MCP-Tool-Workflow) / Roadmap / License
- [x] **Beispiel-Configs `examples/`** (Karte 9) — separate JSON-Files für Copy-Paste (`claude-desktop-config.json`, `cline-mcp-settings.json`, `appsettings.http.json`, `appsettings.readonly.json` + `examples/README.md`)
- [x] **CI/CD-Pipeline** (GitHub Actions) — `ci.yml` (Build + Test bei push/PR auf main) + `release.yml` (Tag-Triggered Publish + GitHub Release mit auto-generierten Notes). Windows-Runner wegen COM-Interop.
- [x] **AGENTS.md-Konvention angenommen** (c544c30, 2026-08-04) — Workspace-weite Migration: Umbenennung `PROJECT.md` → `AGENTS.md` per `git mv` (History via Rename-Detection erhalten). `AGENTS.md` ist Auto-Discovery-Standard für Cursor, OpenAI Codex CLI, Cognition Devin, Cline, Aider, Continue, Windsurf, Cody. Historische `PROJECT.md`-Referenzen in `## Aktueller Status`-Sektionen bleiben dokumentarisch erhalten (siehe Workspace-MEMORY-Direktive 2026-08-02).
- [x] **`.codegraph/` aus VCS ausgeschlossen** (d1e75ca, 2026-08-04) — CodeGraph-Cache (`codegraph.db` ~3,6 MB) gehört nicht ins Repo. Per IDEA-008 vorgesehen, Commit war seit der initialen CodeGraph-Indexierung am 2026-07-31 ausstehend und jetzt nachgeholt.
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
  - Steps: checkout (full history) → setup-dotnet (**kein** NuGet-Cache, siehe "Cache-Status" unten) → restore → publish mit `minimal.pubxml` (Self-Contained + Single-File + Compression) → Zip → Upload-Artifact → GitHub-Release (mit auto-generierten Notes aus Commits seit letztem Tag)
  - Permissions: `contents: write` (fuer `gh release create`)
  - Artifact: `outlook-mcp-server-${{ github.ref_name }}.zip` (OutlookMcpServer.exe ~44.94 MB + pdb + aspnetcorev2_inprocess.dll)

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
  - Asset: `outlook-mcp-server.zip` (OutlookMcpServer.exe ~44,94 MB Self-Contained Single-File + pdb + aspnetcorev2_inprocess.dll)
  - Release-Run-Dauer (Run [29724649418](https://github.com/schirkan/outlook-mcp-server/actions/runs/29724649418)): **~116s (1,9 min)** auf windows-latest
  - CI-Laufzeit (nach Cache-Fix 5bafc3e): **~68s** auf windows-latest
  - Tag-zu-Release-Zeit gesamt: ~2 min (Tag-Push → GitHub Release published + Asset verfuegbar)
  - Auto-Update fuer MCP-Clients (Claude Desktop, Cline, Continue.dev, Cursor): aktuell **manuell** — Download pro Release aus GitHub, exe ersetzen. Auto-Update-Mechanik (GitHub-Releases-API pollen, SemVer-Vergleich, Auto-Replace) ist kein v1-Scope, Folge-Schritt.

- **Cache-Status (Bekannte Limitierung):**
  - Beide Workflows (`ci.yml` + `release.yml`) nutzen **kein** `cache: true` auf `actions/setup-dotnet@v4`. Grund: die Action verlangt eine `packages.lock.json`, die unser Repo (noch) nicht hat — mit `cache: true` wirft die Action "Dependencies lock file is not found", Restore/Build/Tests werden gar nicht erreicht (Fruehausstieg nach ~45s).
  - **Workaround (Commit 5bafc3e):** `cache: true` entfernt, Inline-Kommentar in beiden Workflow-Files erklaert das Trade-off. NuGet-Restore ohne Cache ist ~10-20s langsamer pro CI-/Release-Lauf, fuer unsere Projektgroesse vernachlaessigbar.
  - **Saubere Loesung (Roadmap):** Central Package Management (CPM) via `Directory.Packages.props` aktivieren + `<RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>` in jedem csproj setzen + `dotnet restore --use-lock-file` einmal lokal ausfuehren → generiert `packages.lock.json` pro Projekt. Danach kann `cache: true` in beiden Workflows wieder aktiviert werden, ohne die Cache-Key-Findungs-Failure. Folge-Schritt, kein Dringlichkeits-Bedarf.

## Letzte Commits (2026-07-21 / 2026-07-22 / 2026-07-23 / 2026-07-24 / 2026-07-28 / 2026-07-31 / 2026-08-04)

Dreizehn Commits von Martin zwischen den letzten Pushes aus älteren Sessions auf `origin/main` gelandet - hiermit in der Doku nachgeholt. Vier davon am 2026-07-21 (Phase 1 + Karte 3.5 + CI/CD-Hardening), drei weitere am 2026-07-22 (BodyFormat + Multi-Store + Recipients-Type-Filter), einer am 2026-07-23 (list_mails-Stabilization), drei weitere am 2026-07-24 (BodyFormat-Nullability-API-Change + Active-Mail-Reload + Interop-Partial-Split), einer am 2026-07-28 (`open_item_by_id` UI-Tool für Inspector-Öffnung), einer am 2026-07-31 (OlDefaultFolder-IDs-Korrektur + WellKnownName-Population + Integration-Tests grün). Zwei weitere am 2026-08-04 (AGENTS.md-Konvention + `.codegraph/`-Exclude) im neuen Abschnitt unten:

- **`22479dc`** (2026-07-21 18:46) — **Logging**: `FileLoggerProvider` (optional zuschaltbar via `appsettings.json:Logging:FilePath`), `OutlookJsonContext` (Source-Generation-JSON-Kontext für polymorphe `ActiveItem`-Serialisierung nach PublishTrimmed-Änderung), `OutlookJsonOptionsFactory`, Logger-Konventionen in `Program.cs` + `MailTools.cs` (try/catch um Tool-Bodies mit `OutlookServiceException`-Logging). +590/−39 in 7 Files.
- **`d8ea89e`** (2026-07-21 23:50) — **fix(publish): PublishTrimmed=false**: `.PublishTrimmed` zwingend false weil .NET unter Trimmed `BuiltInComInterop.IsSupported=false` setzt und damit alle COM-Aktivierungen blockiert (`Marshal.GetTypeFromProgID`, `GetObjectForIUnknown`, `[ComImport]`-CoClass). EXE-Größe: 17,5 MB → **44,94 MB**. DECISIONS-Eintrag 2026-07-21 vorhanden.
- **`4149bc2`** (2026-07-21 23:50) — **feat(mail): `list_mails_recursive(scope, top, filter)`**: Neues MCP-Tool für Property-Filter (z. B. `isRead eq false`) über Ordner-Hierarchie. Class-43-Filter (nur `olMail`), EntryID-basierte Deduplication, Hard-Cap bei `top`. Tool-Count: 25 → 26.
- **`6b053ed`** (2026-07-21 23:54) — **fix(com): CallSite dynamic binder**: `Microsoft.CSharp.RuntimeBinder.CallSite` statt `GetType().GetProperty()`-Reflection. Reflection lieferte nur RCW-Wrapper-Properties (nicht die echten IDispatch-Properties), dadurch waren `To`/`CC`/`BCC`-Recipients-Collections leer. Neue Helper-Funktion `GetDynamicProperty` in `OutlookInteropAdapter`.

### 2026-07-22 — BodyFormat + Multi-Store Folder-Resolve + Recipients-Type-Filter

- **`f90aa05`** (2026-07-22 18:50) — **fix recursive**: kleiner Fix in `OutlookInteropAdapter` (Rekursions-Safety) und `OutlookMcpJsonContext` (Source-Generation-Polymorphie-Hook). +26/−1 in 2 Files.
- **`24e4e46`** (2026-07-22 20:43) — **feat(mail): BodyFormat (markdown/text/html) + Multi-Store Folder-Resolve + UnRead-Alias**: `BodyFormat`-Enum (`markdown`|`text`|`html`) + `BodyFormatExtensions.ParseBodyFormat`-Helper. Neuer interner Komponent `HtmlBodyConverter` (`src/OutlookMcpServer.Interop/HtmlBodyConverter.cs`) für HTML→Markdown via ReverseMarkdown 3.7.0. `ResolveDefaultFolderSmart`-Helper für Multi-Store-Profile (Cached-Mode/Exchange + PST). `[UnRead]` → `[Unread]`-Alias-Normalisierung in `list_mails_recursive`-Filter (DASL-Filter ist case-sensitive — Caller können beide Schreibweisen verwenden). `JsonTypeInfo`: `BodyFormat` + `string[]` für Source-Gen registriert. +374/−67 in 12 Files.
- **`df11443`** (2026-07-22 20:56) — **feat(mail,cal): bodyFormat + recipients-by-type**: `bodyFormat`-Parameter an acht Tools exponiert: `list_mails`, `get_mail`, `get_mails`, `list_mails_recursive` (Mail); `get_event`, `list_events` (Calendar); `get_active_item`, `get_selected_items` (ActiveMail-Teil — wirkt nur auf Mails, nicht auf Termine). Outlook-MAPI-MailItems haben keine separaten To/CC/BCC-Properties — `TryMapMailRecipients` liest einmal `mail.Recipients` und filtert nach `.Type` (2=To, 3=CC, 4=BCC). Calendar verwendet jetzt `HtmlBodyConverter` wie Mail (vorher: HTML 1:1 durchgereicht, gleiche Word-Styling-Bloat-Problem). +174/−70 in 9 Files.

### 2026-07-23 — list_mails-Stabilization

- **`45de587`** (2026-07-23 14:28) - **fix(mail): stabilize folder resolve and list_mails path**: `ResolveDefaultFolderSmart` umstrukturiert - zuerst `ns.GetDefaultFolder(olId)`, dann Store-Iteration als Fallback (vorher umgekehrt). Neuer `ResolveDefaultFolderByOlIdWithSchemaCheck` mit `IsMailFolder`-Heuristik (`DefaultItemType == 0 = olMailItem`) für Mail-Schema-Validierung in Multi-Store-Profilen. Neuer `NormalizeDaslFilter` (adapter-side) für `[UnRead]` → `[Unread]`-DASL-Normalisierung (redundant zu Service-Layer in `OutlookService.ListMailsRecursiveAsync`, jetzt auch für direkte `list_mails`-Calls verfügbar). `ListMailsAsync` komplett umstrukturiert mit konsistentem `try/finally`, defensiven `Marshal.ReleaseComObject`-Aufrufen, Error-Handling bei Sort-Failure (`[ReceivedTime]`-Sort kann auf manchen Foldern `0x80020009` werfen). `ListMailsRecursiveAsync` nutzt jetzt `ResolveDefaultFolderByOlIdWithSchemaCheck` statt `ResolveDefaultFolderSmart` - Mail-Schema-Check statt nur Folder-Smart-Resolve. +210/-102 in 1 File (`src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs`).

### 2026-07-24 - BodyFormat-Nullability-API-Change + Active-Mail-Reload + Interop-Partial-Split

- **`f40a4a8`** (2026-07-24 14:44) - **refactor: BodyFormat zu BodyFormat? (null = kein Body); includeBody-Parameter entfernt**: Breaking API-Change für BodyFormat. **Vier** Tools (`list_mails`, `list_mails_recursive`, `search_mails`, `list_events`) haben den `bodyFormat`-Parameter komplett verloren (liefern nie einen Body, Header-only Listen-Calls). **Fünf** Tools (`get_mail`, `get_mails`, `get_event`, `get_active_item`, `get_selected_items`) haben `includeBody` verloren und der `bodyFormat`-Parameter ist jetzt `BodyFormat?` (nullable): Default = `null` = kein Body (vorher: Default = `markdown`). `BodyFormatExtensions.ParseBodyFormat` gibt jetzt `BodyFormat?` zurück (null bei null/empty statt Default `markdown`). Performance- und Token-Ersparnis bei Listen-Calls (kein Body mehr bei `list_mails`). 14 Files, +182/−149.
- **`e4b2591`** (2026-07-24 19:36) - **fix(active-mail): reload inspector item by entry id**: WordEditor-basierte Body-Extraktion für `get_active_item` entfernt (experimentell, lieferte unzuverlässige Ergebnisse). Active `MailItem` wird via `EntryID` + `Namespace.GetItemFromID` neu geladen vor dem Mapping - garantiert konsistentes Body-Verhalten mit `get_mail`/`get_mails`. `get_selected_items` unverändert (kein Inspector-Kontext dort, EntryID-Reload nicht nötig). 1 File, +35/−1 (`src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs` - vor dem Partial-Split, danach wandert die Logik in `OutlookInteropAdapter.ActiveSelection.cs`).
- **`87e5896`** (2026-07-24 20:03) - **refactor(interop): split outlook adapter into partials**: `OutlookInteropAdapter.cs` (2971 Zeilen "Gott-Klasse") aufgeteilt in 8 thematische Partials (`OutlookInteropAdapter.{ActiveSelection,Calendar,CalendarMapping,CalendarMutations,Common,Mail,MailMapping,MailMutations}.cs`). Bessere Wartbarkeit + leichtere Navigation; **keine API-Änderungen** - alle öffentlichen Methoden via Partial-Mechanik erhalten, gleiche Signaturen, gleiche Semantik. Reviewer können thematische Änderungen jetzt isoliert betrachten. 9 Files, +2767/−2961.

### 2026-07-28 — `open_item_by_id` UI-Tool (Inspector-Öffnung)

- **`6edbf78`** (2026-07-28 16:39) - **feat(ui): add open_item_by_id tool**: Neues MCP-Tool `open_item_by_id(id)` öffnet ein Outlook-Item per EntryID in einem Inspector-Fenster (`Item.Display(false)` via `Namespace.GetItemFromID`). Geeignet für `MailItem` und `AppointmentItem` — die Anzeige erfolgt über Outlook selbst. Liefert `OpenItemResult { Opened: bool, Id: string }`. **Erste UI-Mutation** in der bisher read-only Active-Inspector/Selection-Sektion — `get_active_item`/`get_selected_items` bleiben read-only, `open_item_by_id` darf explizit ein Inspector-Fenster öffnen. Validation in `OutlookService.OpenItemByIdAsync` per `ValidateStringNotEmpty(id)`. Tool-Count: 27 → **28** (ActiveSelectionTools 2 → 3). 10 Files, +108/−1. Martin hat README.md + specs/API-DESIGN.md direkt im Commit mit-aktualisiert.

### 2026-07-31 — OlDefaultFolder-IDs-Korrektur + WellKnownName-Population + Integration-Tests grün

- **`5254f19`** (2026-07-30 17:10, Author Copilot) - **fix(outlook): correct OlDefaultFolder IDs and populate WellKnownName**: **Kritischer Bugfix** — die `OlEnumMappings.ToOlDefaultFolder`-Mapping-Werte waren komplett falsch (Copy-Paste-Error, nicht Version-Drift — Werte sind seit Outlook 2010 stabil). Korrekturen gegen Microsoft Learn / Outlook-PIA-Doku verifiziert: `Inbox 12 → 6` (`olFolderInbox`), `SentItems 18 → 5` (`olFolderSentMail`), `Outbox 6 → 4` (`olFolderOutbox`), `DeletedItems 5 → 3` (`olFolderDeletedItems`), `JunkEmail 16 → 23` (`olFolderJunk`), `Drafts 11 → 16` (`olFolderDrafts`). **`Archive → null`** — kein offizielles `OlDefaultFolder`-Konstante; Outlook stellt keinen globalen Archive-Default bereit, Caller müssen expliziten Folder-Pfad verwenden. **Drei weitere Fixes**: (a) `ListMailFoldersAsync` iteriert jetzt `storeRoot.Folders` (sichtbare Mail-Ordner aller Postfächer) statt `ns.Folders` (Postfach-Liste — falsche Semantik für Caller); (b) `MapMailFolder` + `ListMailFoldersAsync` populaten `WellKnownName` per EntryID-Abgleich gegen `Namespace.GetDefaultFolder(olId)` (locale-unabhängig); (c) `GetMailFolderAsync` setzt `WellKnownName` wenn Caller einen Well-Known-Namen statt EntryID übergibt. **Integration-Tests jetzt grün**: `ListMailFolders_ReturnsAtLeastInbox`, `GetMailFolder_Inbox_ReturnsValidFolder` (Outlook 2016 + Exchange-Mailbox). Full Test-Suite **69/69 passing** (63 domain + 6 integration). 3 Files, +146/−30. Author = GitHub Copilot (von Martin approviert).

### 2026-08-04 — AGENTS.md-Konvention + `.codegraph/`-Exclude

- **`d1e75ca`** (2026-08-04 16:29, Author OpenClaw) - **chore(gitignore): exclude .codegraph/ directory**: CodeGraph generiert lokal einen Code-Graph-Cache in `.codegraph/` (`codegraph.db` ~3,6 MB für outlook-mcp-server), der nicht ins VCS gehört. Per IDEA-008 vorgesehen, Commit war seit der initialen CodeGraph-Indexierung am 2026-07-31 ausstehend und jetzt nachgeholt. 1 File, +3.
- **`c544c30`** (2026-08-04 16:29, Author OpenClaw) - **chore: rename PROJECT.md to AGENTS.md**: Workspace-weite Migration auf die `AGENTS.md`-Konvention (https://agents.md). `PROJECT.md` hatte keine Auto-Discovery — AI-Tools (Cursor, OpenAI Codex CLI, Cognition Devin, Cline, Aider, Continue, Windsurf, Cody) sahen die Datei nur bei manuellem Öffnen. `git mv` (Rename-Detection, History bleibt erhalten). Historische `PROJECT.md`-Referenzen in `## Aktueller Status`-Sektionen bleiben dokumentarisch erhalten (siehe Workspace-MEMORY-Direktive 2026-08-02). `README.md` (menschen-lesbar) und `DECISIONS.md` (Architektur-Entscheidungen) bleiben unverändert; `projects/PROJECT-RULES.md` (Konzept-Beschreibung) wurde auf AGENTS.md-Verweise aktualisiert. 1 File (Rename), 2 ++/--.

### Architektur-Implikationen

- **COM-Pattern vereinheitlicht**: Alle Property-Zugriffe laufen jetzt über `GetDynamicProperty` (CallSite-Pfad). `TryGetString`, `TryGetDynamic`, `MapMailItem`, `MapAppointmentItem` profitieren automatisch. Neue Mapping-Code-Pfade müssen nicht extra angepasst werden.
- **Publish-Größe verdoppelt (17,5 MB → 44,94 MB)**: Distribution-Trade-off (Download-Zeit, CI-Laufzeit, User-Disk-Footprint) wurde größer. Akzeptiert weil COM-Funktionalität unter Trim unmöglich ist.
- **`list_mails_recursive` und `get_mails` sind orthogonal**: Ersteres findet Mails anhand eines Property-Filters über die Ordner-Hierarchie, letzteres bulk-loaded Mails anhand einer EntryID-Liste (vermeidet N round-trips nach `list_mails`/`search_mails`). Beide ergänzen sich, ersetzen sich nicht.
- **`bodyFormat`-Parameter für Body-Konvertierung (geändert 2026-07-24)**: **Fünf** Tools haben den `bodyFormat`-Parameter (nullable, Default = `null` = kein Body): `get_mail`, `get_mails`, `get_event`, `get_active_item`, `get_selected_items`. **Vier** Tools liefern nie einen Body und haben den Parameter daher nicht: `list_mails`, `list_mails_recursive`, `search_mails`, `list_events`. Erlaubte Werte: `markdown` (kompakt + LLM-freundlich), `text` (Plain Text, maximal kompakt), `html` (Outlook-Original 1:1, mit Tabellen/Inline-Bildern). Outlook speichert intern immer HTML - `HtmlBodyConverter` (ReverseMarkdown-basiert) konvertiert on-the-fly wenn ein Format angefordert wird. Listen-Calls sparen jetzt Performance + Tokens weil kein Body mehr mitgeliefert wird. Tool-Count bleibt **27** (keine neuen Tools, nur Parameter-Umstellung).
- **Mail-Schema-Validierung im Adapter (2026-07-23)**: `ResolveDefaultFolderByOlIdWithSchemaCheck` + `IsMailFolder`-Heuristik prüfen `DefaultItemType == 0` (`olMailItem`) auf jeden via `GetDefaultFolder` gefundenen Folder. Verhindert, dass Notes/Journal/Tasks-Default-Folder fälschlich als Mail-Folder interpretiert werden und `list_mails_recursive` ungewollte Items liefert. Adapter-seitige `[UnRead]` → `[Unread]`-DASL-Normalisierung (`NormalizeDaslFilter`) macht die Case-Insensitivity zum Default - Caller müssen nicht mehr zwischen `[UnRead]` und `[Unread]` raten.
- **`OutlookInteropAdapter`-Partial-Split (2026-07-24)**: Die 2971-Zeilen-Gott-Klasse wurde in 8 thematische Partials aufgeteilt (`OutlookInteropAdapter.{ActiveSelection,Calendar,CalendarMapping,CalendarMutations,Common,Mail,MailMapping,MailMutations}.cs`). Bessere Wartbarkeit + leichtere Navigation; keine API-Änderung. Compile-Zeiten leicht reduziert durch Parallel-Kompilation der Partials. Reviewer können thematische Änderungen isoliert betrachten (z. B. nur `MailMutations` diff beim Review eines Mutation-Patches).
- **Active-Mail EntryID-Reload (2026-07-24)**: `get_active_item` lädt das aktive `MailItem` jetzt via `Namespace.GetItemFromID(EntryID)` neu, statt die Inspector-Referenz direkt zu mappen. Eliminiert die experimentelle WordEditor-basierte Body-Extraktion (unzuverlässig, Inspector-Kontext konnte während des Reads kippen) und garantiert konsistentes Body-Verhalten mit `get_mail`/`get_mails`. `get_selected_items` unverändert (kein Inspector-Kontext, EntryID-Reload nicht nötig).
- **`open_item_by_id` UI-Mutation (2026-07-28)**: Erste UI-Mutation in der Active-Inspector/Selection-Sektion. `get_active_item` und `get_selected_items` sind weiterhin read-only auf UI-State (COM-Properties `Application.ActiveInspector()`/`ActiveExplorer().Selection`). `open_item_by_id` darf explizit ein Inspector-Fenster öffnen — typischer Use-Case: „öffne die Mail mit EntryID X" nach `list_mails` oder `search_mails`, oder „öffne den Termin mit EntryID Y" nach `list_events`. Architektur-Muster `GetItemFromID + Item.Display(false) + ReleaseComObject im finally` ist Standard für COM-UI-Aktionen. Validation: leere/null EntryID → `InvalidInput`.
- **OlDefaultFolder-IDs-Korrektur (2026-07-31)**: Die in `OlEnumMappings.ToOlDefaultFolder` hinterlegten Outlook-Default-Folder-IDs waren komplett falsch (Copy-Paste-Error — Werte sind seit Outlook 2010 stabil, kein Version-Drift). Korrigiert gegen Microsoft Learn / Outlook-PIA. **Architektur-Lesson**: bei COM/Outlook-Enum-Mappings immer gegen die offizielle Dokumentation verifizieren, nicht aus dem Gedächtnis oder anderen Quellen abschreiben. `Archive` ist KEIN offizielles `OlDefaultFolder` — Outlook stellt keinen globalen Archive-Default bereit; Caller müssen einen expliziten Folder-Pfad verwenden. Neue Hilfs-Liste `KnownFolderMappings` in `OlEnumMappings` listet alle gemappten Well-Known-Folder für Service-/Adapter-Iteration.
- **`wellKnownName`-Population (2026-07-31)**: `MapMailFolder` bekommt optionalen `wellKnownName`-Parameter; `ListMailFoldersAsync` sammelt vorab alle WellKnown-COM-Referenzen via `Namespace.GetDefaultFolder(olId)` und matched Folder-EntryIDs gegen dieses Dictionary (locale-unabhängig, COM-robust gegen Outlook-Lokalisierungs-Variationen). `GetMailFolderAsync` setzt `wellKnownName` wenn Caller einen Well-Known-Namen statt EntryID übergibt. Vorteil für Caller: `folder.wellKnownName` ist jetzt zuverlässig befüllt — vorher war es immer `null` (Folder konnte nur per DisplayName/EntryID identifiziert werden, was sprachabhängig und fragil ist).
- **`ListMailFoldersAsync` Outlook-Semantik-Korrektur (2026-07-31)**: Iteriert jetzt `storeRoot.Folders` (sichtbare Top-Level-Mailordner pro Postfach/Store), nicht `ns.Folders` (Postfächer/Stores selbst — falsche Semantik für einen Caller der „alle Mail-Ordner" will). Vorher lieferte der Call eine Liste von Postfächern wie `"Martin.Pietschmann@itergo.com"` ohne sichtbare Mail-Ordner.
- **AGENTS.md-Konvention (2026-08-04)**: Workspace-weite Migration von `PROJECT.md` zu `AGENTS.md` für Auto-Discovery durch AI-Tools. Hierarchie: `~/.openclaw/workspace/AGENTS.md` (Workspace-Verhaltensregeln, projekt-übergreifend) + `projects/<name>/AGENTS.md` (projektspezifisch: Status, Workboard-Referenzen, offene Punkte) — werden von AI-Tools automatisch gemerged (workspace > project). Saubere Trennung: Verhalten bleibt in der Root-AGENTS.md, Projekt-Spezifika in der Projekt-AGENTS.md. `README.md` bleibt menschen-lesbar, `DECISIONS.md` für Architektur-Entscheidungen, `projects/PROJECT-RULES.md` beschreibt weiterhin das Konzept (mit aktualisierten AGENTS.md-Verweisen).

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
- `AGENTS.md` - diese Datei (Status, Constraints, Architektur-Uebersicht)
- `specs/VISION.md` - Vision, Scope, Ziele/Nicht-Ziele (Detail)
- `specs/ARCHITECTURE.md` - Schichten, Komponenten, Datenfluesse
- `specs/API-DESIGN.md` - MCP-Tools & Resources fuer Mail + Kalender
- `DECISIONS.md` - Architektur- und Designentscheidungen mit Datum + Begruendung
