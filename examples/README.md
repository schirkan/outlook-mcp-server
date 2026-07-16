# examples/

Isolierte Beispiel-Configs fuer OutlookMcpServer. Kopiere die jeweilige Datei an
den Zielort und passe Pfade / Profile-Namen an.

## MCP-Client-Configs

| Datei | Zielort | Client |
|---|---|---|
| `claude-desktop-config.json` | `%APPDATA%\Claude\claude_desktop_config.json` | Claude Desktop |
| `cline-mcp-settings.json` | Cline-Sidebar (MCP-Settings) | Cline (VSCode) |

Pfad-Anpassung: `C:\\Program Files\\Outlook MCP Server\\OutlookMcpServer.exe` →
tatsaechlicher Installationsort. Bei Publish via `minimal.pubxml` ist die Binary
self-contained, ~17,7 MB, kein .NET-Install auf Zielmaschine noetig.

Cline `autoApprove`-Liste enthaelt nur Read-Tools (`list_*`, `get_*`, `search_*`,
`find_meeting_times`). Write-Tools (`send_mail`, `create_draft`, `update_*`,
`delete_*`, `move_*`, `copy_*`, `respond_to_event`, `create_event`, `update_event`,
`delete_event`) und Active-Selection-Tools (`get_active_item`, `get_selected_items`)
sind NICHT in autoApprove → Cline fragt bei jedem Call explizit nach.

## OutlookMcpServer-Configs (appsettings.json-Varianten)

Diese Dateien ersetzen die mit der Binary ausgelieferte `appsettings.json` und
werden via `OUTLOOKMCPSERVER_*`-Env-Vars oder CLI-Args ueberschrieben (siehe
README.md § Configuration).

| Datei | Zweck |
|---|---|
| `appsettings.http.json` | Transport=HTTP/SSE loopback (127.0.0.1:51204), alle Allow*-Flags aktiv |
| `appsettings.readonly.json` | Read-only-Betrieb: `AllowSend`/`AllowDelete`/`AllowCreate` = false. Empfohlen fuer geteilte oder produktive Setups, in denen nur Lese-Zugriff benoetigt wird |

## Schnelltest (HTTP/SSE)

Nach `OutlookMcpServer.exe --config appsettings.http.json`:

```bash
# MCP initialize
curl -X POST http://127.0.0.1:51204/mcp -H 'Content-Type: application/json' \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"curl","version":"1.0"}}}'
```

Erwartete Antwort: JSON-RPC-Response mit `serverInfo.name="OutlookMcpServer"`,
`capabilities.tools={}`. Bei `OutlookNotRunning` Fehler: `Outlook.exe` starten
oder `AutoStartOutlook=true` setzen + 30s warten.

## Verifikation

```powershell
# Claude Desktop: MCP-Settings editieren, Claude neu starten, "list_mail_folders" testen
# Cline: Settings oeffnen, MCP-Server taucht unter "outlook" auf, Tool-Calls testen
```

Bei Fehlern siehe README.md § Troubleshooting.