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
- [ ] Workboard eingerichtet
- [ ] Research abgeschlossen (Graph API Mail+Calendar, Interop Mail+Calendar, MCP C# SDK)
- [ ] Specs geschrieben (VISION, ARCHITECTURE, API-DESIGN)
- [ ] Implementation gestartet (.NET 8 Projekt, MCP SDK Integration, Interop-Wrapper)

## Git

- **Repo-Typ:** GitHub (public)
- **Pfad / URL:** https://github.com/schirkan/outlook-mcp-server
- **Remote(s):** `origin` -> https://github.com/schirkan/outlook-mcp-server.git
- **Eingerichtet am:** 2026-07-15
- **Standard-Branch:** `main`
- **`.gitignore`-Status:** vorhanden

## Workboard

Wird im Anschluss an dieses Initial-Setup eingerichtet.
Board-ID: `outlook-mcp-server`.

## Project Files

- `README.md` - Landing Page (kurzer Ueberblick, Links)
- `PROJECT.md` - diese Datei (Status, Constraints, Architektur-Uebersicht)
- `specs/VISION.md` - Vision, Scope, Ziele/Nicht-Ziele (Detail)
- `specs/ARCHITECTURE.md` - Schichten, Komponenten, Datenfluesse
- `specs/API-DESIGN.md` - MCP-Tools & Resources fuer Mail + Kalender
- `DECISIONS.md` - Architektur- und Designentscheidungen mit Datum + Begruendung
