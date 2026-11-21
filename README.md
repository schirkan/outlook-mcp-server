# Outlook MCP Server

MCP-Server (Model Context Protocol) for accessing a **classic** Outlook installation on Windows via COM-Interop. Reads and writes **emails** and **calendar entries** without cloud, without Microsoft Graph API, without the modern "New Outlook".

> Status: **v1 in active development.** This README describes the current source on `main`. The local server exposes 28 MCP tools over stdio; live Outlook validation of the write paths remains open.

## Features (v1)

### Mail (read)
- `list_mail_folders` / `get_mail_folder` - list + retrieve folders incl. well-known names (`inbox`, `drafts`, `sentItems`, `deletedItems`, `junkEmail`, `outbox`); `wellKnownName` is populated in returned folder objects for caller-friendly lookups. **Note**: `archive` is not an official Outlook default folder and requires an explicit folder EntryID.
- `list_mails` / `get_mail` / `get_mails` (bulk) / `search_mails` / `list_mails_recursive` — list/get/search mails with body, headers, attachments; bulk-get reads up to 50 EntryIDs in one call (returns `{ value, notFoundIds }`); recursive variant walks folder hierarchy with property-filter support (e.g. `[Unread]=true`). **`get_mail`/`get_mails`** haben einen `bodyFormat`-Parameter (`markdown`/`text`/`html`, Default `null` = kein Body); **`list_mails`/`list_mails_recursive`/`search_mails`** liefern keinen Body (Header-only Listen-Calls — spart Tokens + Performance). Mail-Body wird on-the-fly aus Outlook-HTML via `HtmlBodyConverter` (ReverseMarkdown-basiert) konvertiert.
- `get_mail_headers` — internet headers (From/To headers, DKIM, routing)
- `list_attachments` / `get_attachment` — attachment summary + Base64-encoded content

### Mail (write)
- `send_mail` — send directly, supports reply/forward (auto-adds `Re:` / `Fwd:` prefix), attachments via Base64
- `create_draft` — persist draft without sending
- `update_mail` — PATCH-style: `isRead`, `categories`, `importance`
- `move_mail` / `copy_mail` — returns new EntryID (Outlook changes EntryID on Move!)
- `delete_mail` — soft delete moves to the owning mailbox’s Deleted Items. Hard delete removes the message from that folder; server-side recovery or retention depends on mailbox policy.

### Calendar (read)
- `list_calendars` / `get_calendar` — all calendars in profile + default-calendar flag
- `list_events` — time-window query (Start/End, IANA-TZ) with optional calendar scope; returned timestamps are UTC; **kein Body** (Header-only Listen-Call)
- `get_event` — full event incl. attendees, organizer, body, recurrence; supports `bodyFormat`-Parameter (`markdown`/`text`/`html`, Default `null` = kein Body)

### Calendar (write)
- `create_event` — with optional attendees (Required/Optional/Resource) and invitations; optional body is plain text
- `update_event` — PATCH-style fields incl. attendees, reminders, categories; optional body is plain text
- `delete_event` — deletes standalone appointments. Active meetings you organize must be cancelled or deleted in Outlook itself; the COM adapter rejects these requests instead of claiming a cancellation was sent.
- `respond_to_event` — `accepted` / `tentativelyAccepted` / `declined`
- `find_meeting_times` — finds free slots in the default calendar, accounting for overlapping and busy all-day appointments

### Active-Inspector / Selection (COM-only, no Graph equivalent)
- `get_active_item` — currently open item in Outlook's Inspector (mail or event), null if no inspector or out-of-scope type
- `get_selected_items` — items selected in Explorer with scope-filter (`mail` / `calendar` / `any`) and top-cap
- `open_item_by_id` — opens an Outlook item by EntryID in an Inspector window

## Requirements

- Windows 10/11 x64
- **Outlook classic desktop** (2016 / 2019 / 2021 LTSC or Retail / 2024) with configured MAPI profile
- .NET 8 SDK (for building from source)

The server runs in the Windows user context of the active Outlook profile — no separate auth, no cloud credentials.

## Quick Start

### Build from source

```bash
git clone https://github.com/schirkan/outlook-mcp-server.git
cd outlook-mcp-server
dotnet build OutlookMcpServer.sln
```

### Run tests

```bash
# Unit and regression tests (no Outlook profile required)
dotnet test tests/OutlookMcpServer.Domain.Tests/ -c Release
```

Live integration tests are skipped unless explicitly opted in. Use a dedicated classic Outlook test profile:

```powershell
$env:OUTLOOKMCP_RUN_LIVE_TESTS = '1'
dotnet test tests/OutlookMcpServer.IntegrationTests/
```

### Publish (self-contained, minimal size)

A `minimal.pubxml` profile is shipped that produces a single-file, self-contained Windows-x64 exe:

```bash
dotnet publish src/OutlookMcpServer/OutlookMcpServer.csproj \
  -c Release -r win-x64 \
  -p:PublishProfile=minimal
```

Output: `src/OutlookMcpServer/bin/Release/net8.0-windows/win-x64/publish/OutlookMcpServer.exe` (self-contained; no separate .NET runtime install needed). Trimming is disabled to preserve COM activation. Globalization is enabled so calendar operations can handle named time zones and Outlook’s locale-specific dates.

### GitHub releases

The [Release workflow](.github/workflows/release.yml) builds only for newly pushed stable version tags (`vMAJOR.MINOR.PATCH`). Branch pushes and pull requests do not publish releases. The workflow runs the tests without Outlook, embeds the tag version in the Windows-x64 executable, and publishes `outlook-mcp-server.zip` containing `OutlookMcpServer.exe` and `appsettings.json`. The release notes include the ZIP’s SHA-256 checksum.

[Release `v0.1.0`](https://github.com/schirkan/outlook-mcp-server/releases/tag/v0.1.0) is already published. For subsequent releases, create and push a new tag; do not move an existing tag. See [the release procedure](specs/ci-cd-pipeline.md).

## Configuration

Configuration sources (highest priority first):
1. Command-line arguments, e.g. `--OutlookMcpServer:Outlook:AllowSend=false`
2. Environment variables, e.g. `OUTLOOKMCPSERVER__OUTLOOK__ALLOWSEND=false`
3. `appsettings.json` next to the executable, independent of the working directory

The canonical double-underscore environment names override legacy prefixed names such as `OUTLOOKMCPSERVER_OutlookMcpServer__Outlook__AllowSend`. The optional file logger still accepts `Logging:FilePath` in JSON or `OUTLOOKMCPSERVER_Logging__FilePath` in the environment.

### Outlook settings

`appsettings.json`:

```json
{
  "OutlookMcpServer": {
    "Outlook": {
      "ProfileName": null,
      "AutoStartOutlook": true,
      "StartupTimeoutSeconds": 30,
      "AllowSend": true,
      "AllowDelete": true,
      "AllowCreate": true,
      "MaxAttachmentBytes": 26214400
    }
  }
}
```

| Setting | Default | Description |
|---------|---------|-------------|
| `ProfileName` | `null` | Specific Outlook profile name to use, or null for the default profile |
| `AutoStartOutlook` | `true` | If Outlook is not running, start it automatically via `outlook.exe` |
| `StartupTimeoutSeconds` | `30` | Max seconds to wait for Outlook to start |
| `AllowSend` | `true` | Allows mail, meeting invitations, updates with notification, cancellations and meeting responses to be sent |
| `AllowCreate` | `true` | Allows creating drafts/events, updating/copying/moving items and opening an Outlook inspector |
| `AllowDelete` | `true` | Master switch for `delete_mail` and `delete_event` |
| `MaxAttachmentBytes` | `26214400` (25 MB) | Max size for inline attachments in `send_mail` / `create_draft` |

Set all three `Allow*` flags to `false` for read-only mode. When `AllowSend` is `false`, local calendar edits require `sendInvitations: false` or `sendUpdate: false` as applicable; requests that would send a notification are rejected.

### Transport

The only MCP transport is **stdio**. The client starts the server as a local subprocess. No transport setting is needed, and the server opens no network listener.

## MCP Client Setup

Adjust the `command` path in your client's configuration to point to your published executable.

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Program Files\\Outlook MCP Server\\OutlookMcpServer.exe",
      "args": [],
      "env": {
        "OUTLOOKMCPSERVER__OUTLOOK__ALLOWSEND": "true"
      }
    }
  }
}
```

### Cline (VSCode)

Edit `.vscode/settings.json`:

```json
{
  "cline.mcpServers": {
    "outlook": {
      "command": "C:\\Program Files\\Outlook MCP Server\\OutlookMcpServer.exe",
      "args": [],
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

### Continue.dev

Edit `~/.continue/config.json`:

```json
{
  "mcpServers": [
    {
      "name": "outlook",
      "command": "C:\\Program Files\\Outlook MCP Server\\OutlookMcpServer.exe",
      "args": []
    }
  ]
}
```

### Cursor

Edit `~/.cursor/mcp.json` (global) or `.cursor/mcp.json` (per-project):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Program Files\\Outlook MCP Server\\OutlookMcpServer.exe",
      "args": [],
      "env": {
        "OUTLOOKMCPSERVER__OUTLOOK__ALLOWSEND": "true"
      }
    }
  }
}
```

In Cursor, the `outlook` MCP server is then available in Composer (Cmd+I) and in chat via `@outlook`.

### Tool allowlists

The `autoApprove` array in Cline (or analogous setting in other clients) controls which tools can be called without explicit user confirmation. **Always review carefully before allowing `send_mail`, `delete_mail`, `delete_event`**.

## Example Tool Calls

The server speaks MCP/JSON-RPC 2.0 over stdio (one JSON object per line). These are real payloads you can adapt when building custom MCP clients, debugging, or writing integration tests against the contract.

### List unread inbox mails

Request:

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "list_mails",
    "arguments": {
      "folder": "inbox",
      "filter": "isRead eq false",
      "top": 10
    }
  }
}
```

Response (abbreviated):

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "result": {
    "content": [
      {
        "type": "text",
        "text": "[{\"id\":\"AAMkAGI2...\",\"subject\":\"Weekly Status\",\"from\":{\"name\":\"Alice\",\"address\":\"alice@example.com\"},\"receivedDateTime\":\"2026-07-25T09:12:00Z\",\"isRead\":false}]"
      }
    ]
  }
}
```

### Send a mail

Request:

```json
{
  "jsonrpc": "2.0",
  "id": 2,
  "method": "tools/call",
  "params": {
    "name": "send_mail",
    "arguments": {
      "to": ["alice@example.com"],
      "subject": "Weekly Status",
      "bodyText": "Hi Alice,\n\nAttached is this week's status report.\n\nBest,\nBob"
    }
  }
}
```

### Create a calendar event with attendees

Request:

```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "method": "tools/call",
  "params": {
    "name": "create_event",
    "arguments": {
      "subject": "Sprint Review",
      "start": "2026-07-25T14:00:00",
      "end": "2026-07-25T15:00:00",
      "timeZone": "Europe/Berlin",
      "location": "Room 4.1",
      "attendees": [
        { "email": "alice@example.com", "type": "required" },
        { "email": "bob@example.com",   "type": "optional" }
      ]
    }
  }
}
```

## Architecture

```
+-----------------------------+
|  MCP-Client (Claude, Cline) |
+-------------||-------------+
              ||  MCP / JSON-RPC / stdio
+-------------vv-------------+
|  OutlookMcpServer (C#)     |
|  - MailTools (16)           |
|  - CalendarTools (9)        |
|  - ActiveSelectionTools (2) |
+-------------||-------------+
              ||  C# / COM-Interop (dynamic)
+-------------vv-------------+
|  Microsoft.Office.          |
|  Interop.Outlook (NuGet PIA)|
+-------------||-------------+
              ||  MAPI / RPC
+-------------vv-------------+
|  Outlook (classic, running) |
+-----------------------------+
```

Three-layer architecture:
1. **Tools-Layer** (`MailTools`, `CalendarTools`, `ActiveSelectionTools`): MCP-Tool-Definitionen with `[McpServerTool]` + `[Description]`. Stateless, no COM. Discovered automatically via `WithToolsFromAssembly()`.
2. **Domain-Layer** (`OutlookService`, `IOutlookService`): Microsoft-Graph-compatible DTOs, validation (`ValidationHelpers`), policy enforcement (`AllowSend` / `AllowDelete` / `AllowCreate`).
3. **COM-Adapter-Layer** (`OutlookInteropAdapter`, `IInteropOutlookAdapter`): einzige Stelle mit `Marshal.ReleaseComObject`. Single-threaded via `SemaphoreSlim _comLock` (Outlook COM is single-threaded-affine).

API-Semantik follows Microsoft Graph (snake_case tool names, Graph-compatible DTOs with `JsonPropertyName`). Outlook access uses COM interop, not the Microsoft Graph API.

## Development

### Project structure

```
outlook-mcp-server/
+- src/
|  +- OutlookMcpServer/                 # Main project, MCP-Tools + Program.cs (DI, stdio)
|  |  +- Tools/                          # MailTools.cs, CalendarTools.cs, ActiveSelectionTools.cs
|  |  +- Properties/PublishProfiles/    # minimal.pubxml
|  |  +- appsettings.json
|  |  +- Program.cs
|  +- OutlookMcpServer.Domain/          # DTOs, IOutlookService, OutlookService
|  |  +- Abstractions/
|  |  +- Configuration/
|  |  +- Exceptions/                    # OutlookServiceException + ErrorCode enum
|  |  +- Models/                         # Mail/, Calendar/, Common/
|  |  +- Services/                       # OutlookService.cs
|  |  +- Validation/
|  +- OutlookMcpServer.Interop/         # OutlookInteropAdapter (COM-Boundary)
+- tests/
|  +- OutlookMcpServer.Domain.Tests/         # Unit-Tests (FakeOutlookService, xUnit) - 35/35 gruen
|  +- OutlookMcpServer.IntegrationTests/     # Integration-Tests (echtes Outlook, SkippableFact)
+- specs/
|  +- VISION.md                         # Vision, Scope, Ziele/Nicht-Ziele
|  +- ARCHITECTURE.md                   # Schichten, Datenfluesse
|  +- API-DESIGN.md                     # Tool-Specs, Graph-zu-COM-Mapping-Tabelle
+- Properties/PublishProfiles/        # Publish-Profile (verwaltet im jeweiligen src-Projekt)
+- PROJECT.md                          # Status, Constraints, Workboard
+- DECISIONS.md                        # Entscheidungs-Logbuch
+- LICENSE
+- README.md                           # Diese Datei
```

### Adding a new MCP-Tool

1. Add the method signature to `IOutlookService` (Domain) and `IInteropOutlookAdapter` (Interop)
2. Implement the method in `OutlookService` (validation + policy) and `OutlookInteropAdapter` (real COM)
3. Add the method to a `*Tools` class with `[McpServerTool]` (snake_case name) and `[Description]` attributes
4. Add unit tests using `FakeOutlookService`
5. (Optional) Add integration tests using real Outlook via `OutlookIntegrationTestBase`

## Roadmap / Status

Done:
- [x] Solution scaffold (.NET 8 + MCP SDK + Outlook Interop)
- [x] Domain layer (DTOs, IOutlookService, OutlookService with validation)
- [x] Interop-Adapter Grundgerüst (24 Mail/Calendar methods)
- [x] Configuration + DI + stdio transport
- [x] MCP-Tools (MailTools, CalendarTools)
- [x] Unit-Tests xUnit (35/35 grün, FakeOutlookService)
- [x] Echte COM-Impl für alle Mail/Calendar/Active-Selection-Methoden (Karte 3.5)
- [x] Publish-Profil `minimal.pubxml` für minimale Dateigröße (single-file, self-contained, win-x64)

In progress / backlog:
- [ ] Integration-Tests (Projektstruktur steht, Tests skippen ohne Outlook) — Karte 7
- [ ] README erweitern (Build, Konfiguration, MCP-Client-Setup) — Karte 8 (dieser Commit)
- [ ] Beispiel-Config + MCP-Client-Setup (Claude Desktop, Cline, Continue.dev) — Karte 9
- [ ] Phase-3h-spezifische Unit-Tests (Polymorphie / Scope-Filter / Top-Cap / Empty-Selection)
- [ ] Manuelle Smoke-Test-Verifikation der COM-Pfade gegen echtes Outlook-Profil

## License

[MIT License](LICENSE) — permissive open-source license.

## See also

- [`PROJECT.md`](PROJECT.md) — Status, Constraints, Architektur-Details, Workboard
- [`specs/VISION.md`](specs/VISION.md) — Vision, Scope, Ziele/Nicht-Ziele
- [`specs/ARCHITECTURE.md`](specs/ARCHITECTURE.md) — Schichten, Datenflüsse, COM-Lifecycle
- [`specs/API-DESIGN.md`](specs/API-DESIGN.md) — Tool-Specs, Graph-zu-COM-Mapping
- [`DECISIONS.md`](DECISIONS.md) — Entscheidungs-Logbuch mit Datum + Begründung