# Vision — Outlook MCP Server

## Wozu

AI-Agenten (Claude, IDE-Plugins, eigene Bots) sollen mit einem **lokalen klassischen Outlook** (Microsoft Outlook für Windows, klassischer Desktop-Client) interagieren können, **ohne** dass der Agent Zugriff auf Cloud, Graph API oder Exchange-Online-Credentials braucht.

## Was wir bauen (v1)

Einen **MCP-Server** (`OutlookMcpServer`), der als lokaler Prozess auf Windows läuft und folgende Fähigkeiten über standardisierte MCP-Tools/Resources bereitstellt:

- **E-Mails lesen** (Ordner auflisten, einzelne Mails, Suche, Anhänge auflisten)
- **E-Mails schreiben** (neue Mail, Antworten, Allen antworten, Weiterleiten, Entwurf speichern)
- **Kalender lesen** (Kalender auflisten, Termine abfragen, freie Zeitfenster — Self)
- **Kalender schreiben** (Termin erstellen, aktualisieren, löschen, Teilnehmer einladen, auf Einladung antworten)

## Was wir NICHT bauen (v1)

- Cloud-Variante (Microsoft 365, Outlook.com, Exchange Online via Graph)
- Microsoft Graph API als Backend
- Moderner „New Outlook" / One Outlook / Outlook (Preview)
- **Kontakte, Aufgaben, Notizen, Regeln** (Follow-up-Phase, nicht v1)
- Mobile-/Web-Frontend
- Multi-User / Multi-Profil
- Linux/macOS (Windows-only wegen COM)
- Free/Busy über mehrere Postfächer / Delegierte Kalender (v1.1)

## Design-Philosophie

- **API-Semantik nahe an Microsoft Graph** (Mail + Calendar) → Agent-Code, der bereits mit Graph funktioniert, soll mit möglichst wenig Anpassung auch gegen diesen Server laufen. Properties wie `subject`, `body`, `from`, `toRecipients`, `start`, `end`, `attendees`, `location`, `isAllDay`, `importance`, `sensitivity`, `categories` etc. werden 1:1 übernommen.
- **Technologie unabhängig von Graph** → COM-Interop, kein Token/OAuth/Graph-Endpoint involviert. MAPI-spezifische Felder (z. B. `EntryID`) werden als optionale Erweiterungen zurückgegeben.
- **Lokal-only, kein Credential-Store** → der Server läuft im Sicherheitskontext des angemeldeten Windows-Benutzers und nutzt dessen aktives Outlook-Profil. Kein separates Auth-Setup.

## Zielgruppe / Use Cases

1. **Persönlicher AI-Assistent**: „Was steht heute in meinem Kalender?" / „Fass die letzten 5 Mails vom Chef zusammen." / „Schreib eine Antwort auf die Mail vom Chef."
2. **IDE-Copilot / Coding-Agent**: „Lege einen Termin mit Team A für nächsten Donnerstag 14:00 an."
3. **Eigener Workflow-Bot**: Mail-Triage, Kalender-Optimierung, automatische Antwortentwürfe.

## Erfolgskriterien

- Agent kann mit *denselben* Tool-Namen, die er für Microsoft Graph kennt, gegen diesen Server arbeiten (Mapping-Tabelle in `API-DESIGN.md`).
- Server läuft stabil bei Outlook mit >10k Mails und >1000 Terminen (Performance-Test in Implementation-Phase).
- Server kann als Hintergrund-Dienst gestartet und per MCP-Client (Claude Desktop, Continue, Cline) verbunden werden.
- Tests (Unit + Integration) decken mindestens alle MCP-Tools ab.
- Keine Regression bei Outlook-Updates innerhalb derselben Major-Version (Outlook 2016/2019/2021/2024).

## Nicht-Ziele — revisited (Begründung)

- **Kontakte, Aufgaben, Notizen**: andere Objekt-Hierarchie in MAPI, separates Design; erst nach Stabilisierung der Mail+Calendar-Pfade.
- **Regeln, Suchordner, QuickSteps**: MAPI-spezifische Erweiterungen jenseits von Graph → würde API-Semantik sprengen.
- **Multi-User**: würde Server-as-a-Service erfordern (Auth, Isolation); widerspricht „lokal-only, im User-Kontext".
- **Modern/New Outlook**: nutzt andere APIs (Store-App, WebView2), wäre ein paralleler Server.
