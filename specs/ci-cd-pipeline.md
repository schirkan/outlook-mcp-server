# GitHub Actions: CI und Releases

`ci.yml` prüft Pushes und Pull Requests für `main`, veröffentlicht aber keine Releases.

`release.yml` startet den Release-Job nur, wenn ein neuer stabiler Versionstag `vMAJOR.MINOR.PATCH` gepusht wird. Verschobene oder gelöschte Tags lösen keinen Release-Build aus. Einen veröffentlichten Tag nicht wiederverwenden.

1. Den gewünschten Commit auf `main` prüfen. Die nächste freie Version nach `v0.1.0` ist `v0.1.1`.
2. Den neuen Tag anlegen und einzeln pushen, z. B. `git tag -a v0.1.1 -m "Outlook MCP Server v0.1.1"` und `git push origin refs/tags/v0.1.1`.
3. Den [GitHub-Actions-Lauf](https://github.com/schirkan/outlook-mcp-server/actions/workflows/release.yml) prüfen. Nach Erfolg liegt `outlook-mcp-server.zip` im [GitHub-Release](https://github.com/schirkan/outlook-mcp-server/releases).

Der Workflow läuft auf Windows mit dem SDK aus `global.json`. Er baut und testet ohne Outlook-Profil, erstellt mit `minimal.pubxml` eine selbstenthaltende Windows-x64-EXE und prüft ihre Versionsnummer gegen den Tag. Im ZIP liegen nur `OutlookMcpServer.exe` und `appsettings.json`. Der SHA-256-Hash des ZIP steht in den Release-Notizen.

Zur Abnahme: ZIP herunterladen, Hash vergleichen, beide Dateien gemeinsam entpacken und die EXE mit einem lokalen MCP-Client über stdio initialisieren. Ein Outlook-Profil ist für den MCP-Handshake nicht erforderlich.
