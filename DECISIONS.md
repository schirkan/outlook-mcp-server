# DECISIONS — Outlook MCP Server

Architektur- und Designentscheidungen mit Datum + Begründung. Wird iterativ erweitert.

---

## 2026-07-15 — Stack: C# / .NET 8+ / modelcontextprotocol/csharp-sdk

**Status**: accepted

**Kontext**: Martin hat C# + modelcontextprotocol/csharp-sdk vorgegeben.

**Entscheidung**: C# / .NET 8 LTS, offizielles `ModelContextProtocol` NuGet-Paket.

**Alternativen** (verworfen):
- TypeScript + offizielles MCP-TypeScript-SDK: keine native COM-Interop-Bridge, würde Node-COMAddin oder Win32-COM-Bridge erfordern → Mehraufwand
- Rust + mcp-rust: keine ausgereiften Outlook-Interop-Crates → Mehraufwand
- Python + pywin32: Martin-Konvention ist C#/.NET

**Konsequenzen**:
- PIA-Verweis (`Microsoft.Office.Interop.Outlook`) aus NuGet (>= 16.0.0)
- `ModelContextProtocol` (Hauptpaket, mit Hosting/DI)
- `.csproj` zielt auf `net8.0-windows` (wegen COM + PIA)
- TFM ist explizit `net8.0-windows` (nicht `net8.0`), damit COM-PIAs geladen werden

---

## 2026-07-15 — Outlook-Anbindung: COM-Interop, kein Graph API

**Status**: accepted

**Kontext**: Martin hat explizit lokale klassische Outlook-Installation + Interop-Anbindung gewünscht. Kein Graph API, kein Cloud, kein Graph-Endpoint.

**Entscheidung**: `Microsoft.Office.Interop.Outlook` (COM-PIA) aus NuGet. Direkter Out-of-Process-Aufruf (kein Exchange-RPC).

**Konsequenzen**:
- Outlook muss lokal laufen (gleicher Windows-User-Account wie MCP-Server)
- Kein OAuth, kein Token, keine HTTP-Calls
- Kein Multi-User / kein headless Backend (Use-Case „persönlicher Assistent")
- Performance: COM-Aufrufe sind schnell (<50ms pro typische Operation)
- Outlook-Versionen: 2016 / 2019 / 2021 / 2024 (klassischer Desktop) — alle nutzen dasselbe COM-Objektmodell (geringe Variationen dokumentiert)

---

## 2026-07-15 — API-Design: Graph-Mail/Calendar-Semantik, technisch unabhängig

**Status**: accepted

**Kontext**: Martin wünschte „vorhandene APIs (zb. GraphAPI) analysieren um eine ziel-api zu designen". Agent-Frameworks, die bereits mit Graph arbeiten, sollen mit möglichst wenig Anpassung auch gegen diesen Server laufen.

**Entscheidung**: Tool- und Property-Namen folgen Microsoft Graph (`messages`, `mailFolders`, `events`, `calendars`, `subject`, `from`, `toRecipients`, `bodyPreview`, `isAllDay`, `importance`, `sensitivity`, `showAs`, …). Technische Implementation ist COM-Interop, kein HTTP/Graph-Endpoint.

**Konsequenzen**:
- DTOs und JSON-Schema 1:1 zu Microsoft Graph (CamelCase, ISO-8601-Datumswerte, etc.)
- MAPI-spezifische Felder (EntryID, MAPI-Properties) sind zusätzliche optionale Erweiterungen
- Dokumentation verweist auf Graph-Docs für Property-Semantik
- Mapping-Tabelle in `API-DESIGN.md` zeigt Graph-Endpoint → MCP-Tool → COM-Aufruf

---

## 2026-07-15 — Transport: stdio (default) + optional HTTP/SSE (loopback only)

**Status**: accepted

**Kontext**: Lokales Outlook + lokaler MCP-Server. Default-MCP-Clients (Claude Desktop, Cline, Continue) starten den Server als Subprozess über stdio.

**Entscheidung**:
- Default: stdio-Transport (kein Netzwerk-Endpoint)
- Optional HTTP/SSE für Remote-Use-Cases, dann Bind `127.0.0.1` only (kein `0.0.0.0`)

**Konsequenzen**:
- stdio erfordert keine Auth (Subprozess = implizite User-Kontext-Sicherheit)
- HTTP/SSE loopback erfordert keine Auth, aber Client-Whitelist über PID oder Token denkbar (v2)
- Niemals `0.0.0.0` binden (sonst Credential-Drift-Risiko)

---

## 2026-07-15 — Architektur: Schichten MCP-Tools / Domain / COM-Adapter

**Status**: accepted

**Kontext**: Martin-Direktive „Dokumentiere alles". Schichtenarchitektur ermöglicht isolierte Tests (FakeOutlookService) und saubere COM-Boundary.

**Entscheidung**:
1. `MailTools` / `CalendarTools` (MCP-Tool-Layer, stateless)
2. `OutlookService` + `IOutlookService` (Domain, Graph-DTOs, Validierung)
3. `InteropOutlookAdapter` (COM-Boundary, einzige Stelle mit `Marshal.ReleaseComObject`)
4. `Microsoft.Office.Interop.Outlook` (NuGet-PIA)

**Konsequenzen**:
- Unit-Tests ohne Outlook möglich (FakeOutlookService in `OutlookMcpServer.Domain.Tests`)
- Integration-Tests mit Outlook benötigen laufendes Outlook-Profil
- COM-Lifecycle ist an einer Stelle zentralisiert → kein Leak
- Klare Verantwortlichkeiten → einfaches Refactoring

---

## 2026-07-15 — Repo: GitHub, public

**Status**: accepted

**Kontext**: Martin-Direktive „Repo kann public sein" nach initialem Default auf private.

**Entscheidung**: GitHub-Repo `schirkan/outlook-mcp-server` ist **public**. Vor dem ersten Code-Push mit sensiblen Daten schützen (z. B. keine Beispiel-Config mit echten Mail-Adressen).

**Konsequenzen**:
- Issue-Tracker öffentlich → Privacy-freundliche Beispiel-Daten in Doku
- Lizenz muss vor Pull-Requests final sein (offen: MIT vs. Apache-2.0)
- CI/CD erst nach erster stabiler Version öffentlich

---

## 2026-07-15 — Lizenz: Apache-2.0

**Status**: accepted

**Kontext**: Offene Lizenz-Frage aus Initial-Decisions, Martin hat nach Setup-Phase um Festlegung gebeten („Geht auch mit Lizenz?").

**Entscheidung**: **Apache License 2.0**. Gleiche Lizenz wie `modelcontextprotocol/csharp-sdk` → konsistent, kompatibel, inkl. explizitem Patent-Schutz.

**Alternativen** (verworfen):
- **MIT**: permissiver, aber kein expliziter Patent-Schutz
- **BSL / Quelle-zur-Pflicht**: kommerziell restriktiver, schließt Contributions aus

**Konsequenzen**:
- `LICENSE`-Datei im Repo mit Apache-2.0-Standardtext + Copyright 2026 Martin
- README zeigt auf LICENSE
- Contributions unter Apache-2.0 (kein explizites CLA erforderlich)
- Kompatibel mit modelcontextprotocol/csharp-sdk (Apache-2.0) und allen anderen Apache-2.0-Komponenten

> **Korrektur 2026-07-15 (zweiter Eintrag unten):** Lizenz wurde auf **MIT** umgestellt — Martins urspruengliche Frage war „Geht auch MIT?", nicht „leg irgendeine Lizenz fest".

---

## 2026-07-15 — Lizenz-Korrektur: Apache-2.0 → MIT

**Status**: accepted

**Kontext**: Martins Rueckfrage „Geht auch mit Lizenz?" habe ich zunaechst als „ja, Apache-2.0 festlegen" interpretiert und einen Apache-2.0-Commit (82e17ad) gepusht. Martin hat kurz darauf korrigiert: die Frage war „Geht auch die MIT-Lizenz?" — er wollte von Anfang an MIT.

**Entscheidung**: Wechsel von Apache-2.0 zu **MIT License**. Apache-2.0-Commit (82e17ad) bleibt in der History (kein Force-Push auf public Repo), LICENSE-Datei wird durch MIT-Standardtext ueberschrieben.

**Alternativen** (verworfen):
- Apache-2.0 weiterfuehren: ignoriert Martins Wunsch
- Force-Push + Reset auf vor-Apache-Commit: sauberere History, aber Force-Push auf public Repo heikel

**Konsequenzen**:
- LICENSE-Datei: MIT-Standardtext + Copyright 2026 Martin
- README zeigt weiterhin auf LICENSE (jetzt MIT-Inhalt)
- Contributions unter MIT (sehr permissiv, kein expliziter Patent-Schutz)
- Inkompatibilitaet mit Apache-2.0-Projekten bzgl. Patent-Schutz: hier kein Issue, da keine Apache-2.0-Deps mit Patent-Klausel direkt genutzt werden (`modelcontextprotocol/csharp-sdk` waere sonst ein Diskussionspunkt — Martin hat sich bewusst fuer MIT entschieden)

**Lesson**: Bei Rueckfragen lieber 30s nachfragen statt interpretieren. Pattern wiederholt sich (vgl. Lessons „nie ohne Empirie behaupten" und „AI-Recall-Browser-Reader-Bug"): ich neige dazu, schnelle Annahmen zu treffen, wenn der User eine offene Frage stellt. Kuenftig: bei mehrdeutigen Fragen explizit nachfragen, statt aus dem Kontext zu extrapolieren.

---

## 2026-07-15 — Out of Scope v1 (zur Klarstellung)

- **Kontakte / Aufgaben / Notizen**: andere Objekt-Hierarchie in MAPI, separates Design; erst nach Stabilisierung der Mail+Calendar-Pfade
- **Regeln / Suchordner / QuickSteps**: MAPI-spezifische Erweiterungen jenseits von Graph → würde API-Semantik sprengen
- **Multi-User / Multi-Profil**: würde Server-as-a-Service erfordern (Auth, Isolation); widerspricht „lokal-only, im User-Kontext"
- **Modern/New Outlook**: nutzt andere APIs (Store-App, WebView2), wäre ein paralleler Server
- **Free/Busy über mehrere Postfächer / Delegierte Kalender** (v1.1)
- **MIME-Rekonstruktion aus `PR_TRANSPORT_MESSAGE_HEADERS`** (`getMailMime`, v1.1)
- **Search-Folders / Inbox-Rules** (v1.1)

---

## 2026-07-15 — Active-Inspector / Selection als v1-Scope

**Status**: accepted

**Kontext**: Martins Rückfrage „Gibt es eine Möglichkeit, auf die aktuelle geöffnete Mail / den geöffneten Termin / selektierten Eintrag in einer Liste zuzugreifen?". Microsoft Graph kann das nicht (Server-seitig, kein UI-State). COM-Interop bietet `Application.ActiveInspector()`, `Inspector.CurrentItem`, `Application.ActiveExplorer()`, `Explorer.Selection` als native API. Das Feature ist ein **Alleinstellungsmerkmal** dieses Servers gegenüber einem reinen Graph-Wrapper.

**Entscheidung**: Feature wird in **Karte 3.5** (echte COM-Implementierung) aufgenommen — nicht als separate Karte. Spec zuerst (`specs/VISION.md`, `specs/API-DESIGN.md`, `specs/ARCHITECTURE.md`), dann Implementation in einem Schwung mit den 24 COM-Adapter-Methoden.

**Tool-Surface (v1)**:
- `getActiveItem()` → polymorphe Rückgabe (`{kind:"mail",item}` / `{kind:"event",item}` / `null`) per STJ `JsonDerivedType`
- `getSelectedItems(scope, top)` → Liste markierter Items mit Filter und Cap, `OutlookNotActive` wenn kein Explorer aktiv

**Scope v1 für Active-Inspector**: `mail` (gelesen oder im Entwurfs-/Edit-Modus), `event` (gelesen oder in Bearbeitung). Tasks/Contacts bleiben v1.1.
**Scope v1 für Selection**: `ActiveExplorer()` muss aktiv sein, sonst Fehler. `Selection.Count == 0` ist valide (kein Fehler, leere Liste). `scope`-Filter ist „mail" / „calendar" / „any" (default „any").

**Alternativen** (verworfen):
- **Eigene Karte nach 3.5** (z. B. Karte 3.6): bricht Scope, COM wird eh neu geschrieben → ein Schwung günstiger
- **Reines Spec-Dokument ohne Impl**: macht das Feature nicht nutzbar
- **Auch für Tasks/Contacts**: sprengt V1-Scope laut VISION.md
- **Write auf den UI-State (z. B. Selection manipulieren)**: würde API-Semantik sprengen und ist nicht angefragt — bleibt explizit out-of-scope

**Konsequenzen**:
- 2 neue Methoden auf `IOutlookService` und `IInteropOutlookAdapter`: `GetActiveItemAsync`, `GetSelectedItemsAsync`
- 2 neue DTOs im Domain-Layer: `ActiveItem` (abstract record + `ActiveMail`/`ActiveEvent` Sub-Types), `SelectedItems` (record mit `value` + `count`)
- 2 neue MCP-Tools: `getActiveItem`, `getSelectedItems` — leben in neuer `ActiveSelectionTools`-Klasse in `src/OutlookMcpServer/Tools/`
- `OutlookInteropAdapter` bekommt 2 zusätzliche COM-Mappings (Inspector + Explorer)
- `FakeOutlookService` + `FakeInteropAdapter` erweitert; `OutlookServiceTests` deckt `OutlookNotActive` ab; `ActiveSelectionToolsTests` deckt Polymorphie + Scope-Filter ab
- Specs aktualisiert: `VISION.md` (in-scope), `API-DESIGN.md` (Tool-Spec + Mapping-Tabelle), `ARCHITECTURE.md` (Datenfluss + Single-Threading-Hinweis)
- Token-Budget: aktiv — `ActiveSelectionTools`-Klasse + 2 Methoden + DTOs in bestehende Tests einplanen (vermutlich +200-350 LOC im Domain/Tools, +300-450 LOC im Interop)
