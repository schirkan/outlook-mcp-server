# Outlook MCP Server

MCP-Server fuer **klassische Outlook-Installationen** auf Windows. Liest und schreibt **E-Mails** und **Kalender-Eintraege** ueber `Microsoft.Office.Interop.Outlook` (COM) - ohne Cloud, ohne Graph API.

> Status: **Setup-Phase** - Specs werden gerade erarbeitet (siehe `PROJECT.md`).

## Schnellueberblick

- **Sprache / Stack:** C# / .NET 8+ mit [`modelcontextprotocol/csharp-sdk`](https://github.com/modelcontextprotocol/csharp-sdk)
- **Outlook-Anbindung:** lokal via COM-Interop (kein Graph, kein Cloud)
- **Scope v1:** E-Mails (lesen/schreiben) + Kalender (lesen/schreiben)
- **Plattform:** Windows (x64)

## Dokumentation

| Datei | Zweck |
|---|---|
| [`PROJECT.md`](PROJECT.md) | Status, Constraints, Architektur-Uebersicht |
| [`specs/VISION.md`](specs/VISION.md) | Vision, Scope, Ziele/Nicht-Ziele |
| [`specs/ARCHITECTURE.md`](specs/ARCHITECTURE.md) | Schichten, Komponenten, Datenfluesse |
| [`specs/API-DESIGN.md`](specs/API-DESIGN.md) | MCP-Tools/Resources (Ziel-API) |
| [`DECISIONS.md`](DECISIONS.md) | Entscheidungs-Logbuch |

## Voraussetzungen

- Windows 10/11 x64
- Outlook (klassisch) 2016 / 2019 / 2021 (LTSC oder Retail) mit eingerichtetem MAPI-Profil
- .NET 8 SDK (fuer Build aus Source)

## Lizenz

[Apache License 2.0](LICENSE) — gleiche Lizenz wie [`modelcontextprotocol/csharp-sdk`](https://github.com/modelcontextprotocol/csharp-sdk), inkl. explizitem Patent-Schutz.
