# examples/

Isolierte Beispiel-Configs fuer OutlookMcpServer. Kopiere die jeweilige Datei an
den Zielort und passe Pfade / Profile-Namen an.

## MCP-Client-Configs

Der MCP-Client startet den Server lokal über stdio. Eine Transporteinstellung ist nicht nötig.

| Datei | Zielort | Client |
|---|---|---|
| `claude-desktop-config.json` | `%APPDATA%\Claude\claude_desktop_config.json` | Claude Desktop |
| `cline-mcp-settings.json` | Cline-Sidebar (MCP-Settings) | Cline (VSCode) |

Pfad-Anpassung: `C:\\Program Files\\Outlook MCP Server\\OutlookMcpServer.exe` →
tatsaechlicher Installationsort. Bei Publish via `minimal.pubxml` ist die Binary
self-contained; keine separate .NET-Installation notwendig.

Cline `autoApprove`-Liste enthaelt nur Read-Tools (`list_*`, `get_*`, `search_*`,
`find_meeting_times`). Write-Tools (`send_mail`, `create_draft`, `update_*`,
`delete_*`, `move_*`, `copy_*`, `respond_to_event`, `create_event`, `update_event`,
`delete_event`) und Active-Selection-Tools (`get_active_item`, `get_selected_items`)
sind NICHT in autoApprove → Cline fragt bei jedem Call explizit nach.

## OutlookMcpServer-Configs (appsettings.json-Varianten)

Kopiere die gewählte Datei als `appsettings.json` neben die ausführbare Datei.
CLI-Argumente haben Vorrang vor `OUTLOOKMCPSERVER__…`-Umgebungsvariablen, diese
vor `appsettings.json` (siehe [Configuration](../README.md#configuration)).

| Datei | Zweck |
|---|---|
| `appsettings.readonly.json` | Read-only-Betrieb: `AllowSend`/`AllowDelete`/`AllowCreate` = false. Empfohlen fuer geteilte oder produktive Setups, in denen nur Lese-Zugriff benoetigt wird |

## Verifikation

Starte den MCP-Client mit der gewählten Config neu und prüfe, ob er die Tools
unter `outlook` anzeigt. Das Auflisten der Tools benötigt kein Outlook-Profil.

```powershell
# Claude Desktop: MCP-Settings editieren, Claude neu starten, "list_mail_folders" testen
# Cline: Settings oeffnen, MCP-Server taucht unter "outlook" auf, Tool-Calls testen
```

Bei Fehlern siehe README.md § Troubleshooting.