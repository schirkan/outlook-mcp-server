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

## 2026-07-15 — Lizenz: TBD

**Status**: open

**Optionen**:
- **MIT** (permissiv, oft für Open-Source-Tools)
- **Apache-2.0** (permissiv + Patent-Schutz, gleiche Lizenz wie modelcontextprotocol/csharp-sdk)
- **BSL / Quelle-zur-Pflicht** (kommerziell restriktiver)

Vorschlag: **Apache-2.0** (gleiche Lizenz wie csharp-sdk → konsistent, gute Kompatibilität). Entscheidung in einer Folge-Session vor erstem Pull-Request von außen.

---

## 2026-07-15 — Out of Scope v1 (zur Klarstellung)

- **Kontakte / Aufgaben / Notizen**: andere Objekt-Hierarchie in MAPI, separates Design; erst nach Stabilisierung der Mail+Calendar-Pfade
- **Regeln / Suchordner / QuickSteps**: MAPI-spezifische Erweiterungen jenseits von Graph → würde API-Semantik sprengen
- **Multi-User / Multi-Profil**: würde Server-as-a-Service erfordern (Auth, Isolation); widerspricht „lokal-only, im User-Kontext"
- **Modern/New Outlook**: nutzt andere APIs (Store-App, WebView2), wäre ein paralleler Server
- **Free/Busy über mehrere Postfächer / Delegierte Kalender** (v1.1)
- **MIME-Rekonstruktion aus `PR_TRANSPORT_MESSAGE_HEADERS`** (`getMailMime`, v1.1)
- **Search-Folders / Inbox-Rules** (v1.1)
