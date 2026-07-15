# Architektur — Outlook MCP Server

## Schichten (Überblick)

```
+------------------------------------------------+
| MCP-Client (Claude Desktop, Continue, Cline,   |
| eigenes Tool, etc.)                            |
+--------------------||--------------------------+
                     || MCP / JSON-RPC
                     || (stdio default, HTTP/SSE loopback optional)
+--------------------vv--------------------------+
| OutlookMcpServer (C# / .NET 8+ / win-x64)      |
|                                                |
| +-----------------+    +---------------------+ |
| | MailTools       |    | CalendarTools       | |
| | (MCP-Tools)     |    | (MCP-Tools)         | |
| +--------||-------+    +---------||-----------+ |
|          ||                       ||           |
| +--------vv-----------------------vv---------+ |
| | OutlookService (Domain)                    | |
| | - MailOperations                           | |
| | - CalendarOperations                       | |
| +-------------------||-----------------------+ |
|                     ||                         |
+---------------------||--------------------------+
                      || IOutlookService (Mock-fähig)
+---------------------vv--------------------------+
| InteropOutlookAdapter                          |
| - bindet Microsoft.Office.Interop.Outlook ein  |
| - versteckt COM-Boilerplate (Release, Marshal) |
+---------------------||--------------------------+
                      || COM (IDispatch, late-bound wo nötig)
+---------------------vv--------------------------+
| Microsoft.Office.Interop.Outlook (PIA / NuGet) |
+---------------------||--------------------------+
                      || MAPI / Outlook-Profil
+---------------------vv--------------------------+
| Outlook (klassisch) — laufendes Windows-Profil  |
+------------------------------------------------+
```

## Komponenten

### 1. OutlookMcpServer (Top-Level)

- **Verantwortung**: MCP-Protokoll-Handling, Tool-Registrierung, JSON-Schema-Generierung, Konfiguration
- **Stack**: `ModelContextProtocol` NuGet, attribute-based registration
- **Transport**: stdio (Default für lokale Clients) oder HTTP/SSE (loopback) — Auswahl per `appsettings.json`
- **DI**: `Microsoft.Extensions.Hosting` + `Microsoft.Extensions.DependencyInjection`
- **Logging**: `Microsoft.Extensions.Logging` → `ILogger<T>` → Serilog/Sink konfigurierbar

### 2. MailTools / CalendarTools (MCP-Tool-Provider)

- **Verantwortung**: exponieren MCP-Tools (stateless Wrapper)
- **Pattern**: eine Klasse pro Domain (`MailTools`, `CalendarTools`), jede Methode = ein MCP-Tool
- **Attribute**: `[McpServerToolType]` (Klasse), `[McpServerTool]` + `[Description]` (Methode)
- **Parameter-Binding**: `[Description]`-annotierte Parameter, komplexe Typen via System.Text.Json
- **Beispiel**:
  ```csharp
  [McpServerToolType]
  public sealed class MailTools(IOutlookService svc, ILogger<MailTools> log)
  {
      [McpServerTool(Name = "listMails"),
       Description("Listet Mails im angegebenen Ordner auf, neueste zuerst.")]
      public async Task<MailList> ListMails(
          [Description("Well-known folder name (inbox|drafts|sentItems|...) oder EntryID")]
          string folder,
          [Description("Max. Anzahl Treffer (1-100)")]
          int top = 25,
          [Description("Anzahl zu ueberspringender Treffer fuer Paginierung")]
          int skip = 0)
      {
          return await svc.ListMailsAsync(folder, top, skip);
      }
  }
  ```

### 3. OutlookService (Domain)

- **Verantwortung**: Business-Logik, Graph-konforme DTOs, Validierung, Fehler-Mapping
- **DTOs** (siehe `API-DESIGN.md`): `MailMessage`, `MailFolder`, `CalendarEvent`, `Attendee`, `DateTimeTimeZone`, `ItemBody`, `Recipient`, `Location`, `ResponseStatus` …
- **Interfaces**: `IOutlookService` (Mock-fähig für Tests)
- **Implementierungen**:
  - `OutlookService` → produktiv (delegiert an `IInteropOutlookAdapter`)
  - `FakeOutlookService` → in-memory für Unit-Tests
- **Validierung**: Input-Validierung (Pflichfelder, Datums-Ranges, Body-Size-Limits, E-Mail-Format)
- **Fehler-Mapping**: Adapter wirft `OutlookInteropException` → Service mappt auf `OutlookServiceException` mit Code

### 4. InteropOutlookAdapter (COM-Boundary)

- **Verantwortung**: **einzige** Stelle, die COM-Objekte anfasst; versteckt COM-Boilerplate
- **Pattern**:
  - `Application`/`NameSpace`/`MAPIFolder`/`MailItem`/`AppointmentItem` als `dynamic` (late-bound) oder stark typisiert mit PIA
  - COM-Objekte in `try/finally` mit `Marshal.ReleaseComObject` (oder `using`-Wrapper)
  - Helper-Methoden wie `WithItem<T>(EntryID, action)`, `WithFolder<T>(path, action)`
- **Versteckt**: Application-Bootstrapping, `NameSpace("MAPI")`-Lookup, Folder-Traversal, `PropertyAccessor`-Aufrufe, MAPI-String-Property-Namen (DASL)
- **Exception-Mapping**: `COMException` → `OutlookInteropException` mit Code (z. B. `OutlookBusy`, `ItemNotFound`, `PermissionDenied`)
- **Lifecycle**: Outlook muss bereits laufen (Use Case: persönlicher Assistent, kein Service für headless-Backend)
- **Start-Strategie**: Wenn Outlook nicht läuft → automatisch starten via `Process.Start("outlook.exe")`, mit kurzem Retry-Loop (max 30s), dann Fehler

### 5. Models / DTOs

- **Graph-kompatibel**: Property-Namen und Typen 1:1 wie Microsoft Graph
- **Beispiele**:
  - `MailMessage { Id, ConversationId, Subject, BodyPreview, Body, From, ToRecipients, CcRecipients, BccRecipients, SentDateTime, ReceivedDateTime, HasAttachments, Importance, IsRead, Categories }`
  - `CalendarEvent { Id, Subject, Body, Start (DateTimeTimeZone), End (DateTimeTimeZone), Location, Attendees, Organizer, IsAllDay, Importance, Sensitivity, ShowAs, Categories, IsCancelled, IsReminderOn, ReminderMinutesBeforeStart }`
- **JSON-Serialisierung**: `System.Text.Json` mit `JsonNamingPolicy.CamelCase` + `JsonPropertyName` für Graph-konforme Keys
- **Datum/Zeit**: `DateTimeOffset` in UTC für Input/Output; Mapping auf Outlook-`DateTime` lokal oder UTC je nach Property

## Datenflüsse

### Mail lesen: „Zeige letzte 10 Inbox-Mails"

```
Client (Claude)
  -> MCP: tools/call listMails {folder: "inbox", top: 10}
OutlookMcpServer
  -> MailTools.ListMails("inbox", 10, 0)
  -> IOutlookService.ListMailsAsync("inbox", 10, 0)
  -> InteropOutlookAdapter:
     - NameSpace.GetDefaultFolder(olFolderInbox)
     - Items.Restrict("@SQL=(...ReceivedTime...)")
     - foreach MailItem: map -> MailMessage DTO
  -> JSON-serialisieren
Client
  <- { value: [...], nextSkip: 10 }
```

### Mail senden: „Schick Antwort an Chef"

```
Client
  -> MCP: tools/call sendMail { to:["chef@..."], subject:"Re: ...", body:..., replyToId:"EntryID-..." }
OutlookMcpServer
  -> MailTools.SendMail(...)
  -> IOutlookService.SendMailAsync(...)
  -> InteropOutlookAdapter:
     - if replyToId: GetItemFromID(replyToId) -> Reply() or ReplyAll() -> .To, .Subject, .Body
     - else: Application.CreateItem(olMailItem), .To, .Subject, .Body, .HTMLBody
     - .Send()
  -> return: { id: EntryID-of-SentItem }
Client
  <- { sent: true, id: "..." }
```

### Termin anlegen: „Meeting Donnerstag 14:00 mit Team A"

```
Client
  -> MCP: tools/call createEvent { subject:"...", start:{dateTime, timeZone}, end:..., attendees:[...], location:"..." }
OutlookMcpServer
  -> CalendarTools.CreateEvent(...)
  -> IOutlookService.CreateEventAsync(...)
  -> InteropOutlookAdapter:
     - GetDefaultFolder(olFolderCalendar)
     - AppointmentItem appt = (AppointmentItem)calendar.Items.Add(OlItemType.olAppointmentItem)
     - Setze Subject, Start, End, Location, AllDayEvent
     - foreach attendee: Recipients.Add(address); ResolveAll()
     - appt.Send()  // verschickt Einladungen
Client
  <- { id: "EntryID-..." }
```

### Active-Item / Selection: „Verarbeite die offene Mail / die markierten Mails"

```
Client
  -> MCP: tools/call getActiveItem
OutlookMcpServer
  -> ActiveSelectionTools.GetActiveItemAsync()
  -> IOutlookService.GetActiveItemAsync()
  -> InteropOutlookAdapter:
     - var insp = _application.ActiveInspector()   // null wenn kein Inspector
     - if (insp == null) return null
     - switch (insp.CurrentItem):
         MailItem mi        -> map -> ActiveMail  { kind: "mail",  item: MailMessage }
         AppointmentItem ai -> map -> ActiveEvent { kind: "event", item: CalendarEvent }
         _                  -> null (z. B. Contact/Task in v1.1)
Client
  <- { kind: "mail", item: {...} } | null

Client
  -> MCP: tools/call getSelectedItems { scope: "mail", top: 50 }
OutlookMcpServer
  -> ActiveSelectionTools.GetSelectedItemsAsync("mail", 50)
  -> IOutlookService.GetSelectedItemsAsync("mail", 50)
  -> InteropOutlookAdapter:
     - var exp = _application.ActiveExplorer()
     - if (exp == null) throw OutlookServiceException(OutlookNotActive, "kein Explorer aktiv")
     - var sel = exp.Selection
     - fuer i=1..min(sel.Count, top): Map -> ActiveItem (Mail/Event, andere gefiltert)
Client
  <- { value: [...], count: N }
```

**Wichtige Eigenschaft**: Beide Tools sind reine Reads auf den UI-State. Sie verändern weder Selection noch offene Inspektoren. Das hält die Semantik klar und macht sie kombinierbar mit klassischen Mutationen (z. B. „verschicke die Reply auf die selektierte Mail").

### Active-Selection und Outlook-Single-Threading

COM-Single-Threading-Apartment: alle Outlook-COM-Aufrufe laufen im STA-Thread. Der MCP-Server ist per Default `Main(args).Run()` ein Top-Level STA-Thread (siehe Karte 5: `Host.CreateApplicationBuilder` + `app.Run()`); Outlook-COM-Aufrufe müssen in den richtigen Apartment gemarshallt werden. Wir verwenden die PIA-eigenen Calls (sie sind MTA-frei), nicht `Task.Run` für COM-Operationen.

## Sicherheit & Isolation

- **Prozessgrenze**: Server läuft im User-Kontext des angemeldeten Windows-Benutzers. Outlook-Profil ist implizit.
- **Kein Credential-Store**: keine Tokens, keine Secrets im Repo oder in Config-Dateien.
- **Kein Netzwerk-Endpoint by default**: stdio-Transport, Client startet Server als Subprozess.
- **HTTP/SSE nur loopback**: bei `Transport=http` immer `Host=127.0.0.1`, niemals `0.0.0.0`.
- **Operations-Whitelist** in `appsettings.json`: `AllowSend`, `AllowDelete` — für Read-only-Setups einzeln deaktivierbar.
- **Logging**: keine Mail-Bodies in Logs (nur Subject + IDs); kein PII ohne expliziten Log-Level.
- **Attachment-Streaming**: max 25 MB pro Attachment, größere als Fehler zurück (statt im Speicher halten).

## Performance & Limits

- **Paginierung**: `top` (1-100 Mail, 1-250 Calendar), `skip` für Listen
- **Body-Size**: bis 25 MB pro Mail (Outlook-Default), größere Mails als Anhang-Stub zurückgeben
- **Attachment-Downloads**: separat via `getAttachment` Tool, nie inline in der Mail-DTO
- **Caching**: keins in v1; COM-Aufrufe sind hinreichend schnell für persönliche Assistenten-Workloads
- **Parallelität**: Outlook-COM ist single-threaded-affine; Service serialisiert Operationen pro Outlook-Application-Instanz über `SemaphoreSlim`

## Konfiguration (`appsettings.json`)

```json
{
  "OutlookMcpServer": {
    "Transport": "stdio",
    "Http": {
      "Host": "127.0.0.1",
      "Port": 51204
    },
    "Outlook": {
      "ProfileName": null,
      "AutoStartOutlook": true,
      "StartupTimeoutSeconds": 30,
      "AllowSend": true,
      "AllowDelete": true,
      "AllowCreate": true,
      "MaxAttachmentBytes": 26214400
    },
    "Logging": {
      "LogLevel": {
        "Default": "Information",
        "Microsoft.Office.Interop": "Warning"
      }
    }
  }
}
```

## Projektstruktur

```
outlook-mcp-server/
├── OutlookMcpServer.sln
├── src/
│   ├── OutlookMcpServer/                 # Top-Level (MCP-Host, DI, Config)
│   │   ├── Program.cs
│   │   ├── appsettings.json
│   │   ├── Tools/
│   │   │   ├── MailTools.cs
│   │   │   └── CalendarTools.cs
│   │   └── OutlookMcpServer.csproj
│   ├── OutlookMcpServer.Domain/          # DTOs, Interfaces, OutlookService
│   │   ├── Models/
│   │   │   ├── Mail/
│   │   │   └── Calendar/
│   │   ├── Abstractions/
│   │   │   └── IOutlookService.cs
│   │   ├── Services/
│   │   │   └── OutlookService.cs
│   │   └── OutlookMcpServer.Domain.csproj
│   └── OutlookMcpServer.Interop/        # COM-Adapter (einzige COM-Stelle)
│       ├── InteropOutlookAdapter.cs
│       ├── Exceptions/
│       │   └── OutlookInteropException.cs
│       └── OutlookMcpServer.Interop.csproj
├── tests/
│   ├── OutlookMcpServer.Domain.Tests/    # Unit-Tests mit FakeOutlookService
│   └── OutlookMcpServer.IntegrationTests/ # Integration-Tests (Outlook erforderlich)
├── specs/                                 # Design-Specs (siehe specs/VISION.md)
├── README.md
├── PROJECT.md
├── DECISIONS.md
└── .gitignore
```

## Tests

- **Unit** (`OutlookMcpServer.Domain.Tests`): `FakeOutlookService` mit deterministischen Daten; testet Tool-Layer, DTO-Mapping, Validierung, Fehler-Codes. **Läuft überall**, kein Outlook nötig.
- **Integration** (`OutlookMcpServer.IntegrationTests`): `InteropOutlookAdapter` gegen lokales Outlook-Profil (xUnit `[Fact(Skip=...)]` wenn Outlook fehlt); testet echte COM-Interaktion mit Test-Profil oder Default-Profil.
- **E2E**: manuell — MCP-Client (Claude Desktop/Cline) → MCP-Server → Outlook. Optional: Smoke-Test-Script das `sendDraft` → Mail in Drafts prüft.
- **Coverage-Ziel**: >80% der Domain- und Adapter-Schicht.
