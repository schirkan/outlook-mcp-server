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
- [x] MCP-Tools (23 Tools in `MailTools` + `CalendarTools`)
- [x] Unit-Tests xUnit (35/35 grün, `OutlookMcpServer.Domain.Tests`)
- [/] **Echte COM-Interop-Impl** (Karte 3.5) — Phasen 1, 2, 3a, 3b, 3c + Docs P7 (`resolveName`) abgeschlossen; ausstehend: Mail-Send, Mail-Move/Copy/Update, Calendar (Appointment-Lookup/Modify), Attendees, ActiveSelection-Tools + Tests
- [ ] Integration-Tests mit Outlook-Profil (Karte 7)
- [ ] README erweitern (Build, Konfiguration, Verwendung)
- [ ] Beispiel-Config + MCP-Client-Setup (Claude Desktop, Cline)

## Git

- **Repo-Typ:** GitHub (public)
- **Pfad / URL:** https://github.com/schirkan/outlook-mcp-server
- **Remote(s):** `origin` -> https://github.com/schirkan/outlook-mcp-server.git
- **Eingerichtet am:** 2026-07-15
- **Standard-Branch:** `main`
- **`.gitignore`-Status:** vorhanden

## Workboard

- **Board-ID:** `outlook-mcp-server`
- **Stats:** 10 Karten total · 6 done · 1 running · 3 backlog
- **Laufende Karte:**
  - `f78b75ed-6f77-439c-abd2-7b03a1d9f371` — Impl: Echte COM-Interop (24 Methoden + Active-Selection) (high)
- **Backlog:**
  - `022e0b4e-f07a-499a-904d-4c4a49443871` — Tests: Integration (Outlook-Profil, xUnit, COM-Adapter)
  - `737dbaa1-f169-4094-af81-a6204ece9052` — Doku: README erweitern (Build, Konfiguration, Verwendung)
  - `d8753677-91bc-4181-9e39-4c5139d12990` — Doku: Beispiel-Config + MCP-Client-Setup (Claude Desktop, Cline)
- **Done:** Solution-Scaffold · Domain-Schicht · Interop-Adapter Grundgerüst · Konfiguration + Transport · MCP-Tools · Unit-Tests (35/35 grün)

## Project Files

- `README.md` - Landing Page (kurzer Ueberblick, Links)
- `PROJECT.md` - diese Datei (Status, Constraints, Architektur-Uebersicht)
- `specs/VISION.md` - Vision, Scope, Ziele/Nicht-Ziele (Detail)
- `specs/ARCHITECTURE.md` - Schichten, Komponenten, Datenfluesse
- `specs/API-DESIGN.md` - MCP-Tools & Resources fuer Mail + Kalender
- `DECISIONS.md` - Architektur- und Designentscheidungen mit Datum + Begruendung
