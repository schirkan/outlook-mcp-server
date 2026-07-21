# DECISIONS — Outlook MCP Server

Architektur- und Designentscheidungen mit Datum + Begruendung. Wird iterativ erweitert.

Hinweis: Diese Datei enthaelt einen Wiederherstellungs-Stub mit Verweis auf den Code-Stand von 2026-07-21.
Der vollstaendige DECISIONS-Inhalt der vorherigen Sitzung (13 Eintraege) wird in einer separaten Sitzung nachgetragen.

## 2026-07-21 — Tool `list_mails_recursive` (Property-Filter ueber Ordner-Hierarchie)

Status: accepted

Kontext: Martins Use-Case "alle ungelesenen Mails in allen Ordnern" war mit den existierenden Tools `list_mails` (nur ein Ordner) und `search_mails` (Subject/Sender-Textsuche) nicht abdeckbar. Nach dem Fix von `search_mails` mit `folderId=null` war die rekursive Reichweite da, aber nur fuer Text-Felder — kein `[UnRead]`, kein `[HasAttachments]`, kein `[Importance]`. User-Anfrage (verbatim): "kann man die suche dafuer verwenden? also ordneruebergreifend nach ungelesenen mails suchen?" / "der call hat technisch funktioniert, aber kein ergebnis geliefert, weil die Mails in Unterordnern liegen."

Entscheidung: Neues MCP-Tool `list_mails_recursive(scope=..., top=25, filter="...")` mit folgenden Eigenschaften:
- Tool-Name `list_mails_recursive` — Naming analog zu `list_mails`/`list_mail_folders`; Suffix `_recursive` signalisiert Ordner-Hierarchie-Walk.
- scope-Parameter optional, Komma-getrennte Liste von Well-Known-Ordner-Namen (inbox, drafts, sentItems, deletedItems, junkEmail, archive, outbox). Default = alle genannten. Validierung pro Eintrag gegen `WellKnownFolder.IsKnownMailFolder(...)` im Domain-Layer.
- filter optionaler DASL-Ausdruck, 1:1 zur `list_mails`-Semantik (`Items.Restrict(filter)`). Haeufig: `[UnRead] = true`, `[Unread] = true AND [HasAttachments] = true`, `[Importance] = 2`, `[ReceivedTime] >= '7/1/2026 12:00:00 AM'`. Bei ungueltiger Syntax Fallback auf ungefilterte Iteration pro Folder + Warnung im Log.
- Reichweite: pro Scope-Eintrag `MAPINamespace.GetDefaultFolder(olDefaultFolders.X)`, dann rekursiv in alle Unterordner via `MAPIFolder.Folders`. Alle Stores (Standard-OST, PSTs) werden durchsucht.
- Deduplizierung: `HashSet<string>` auf `MailItem.EntryID` (Outlook-eindeutig, auch ueber Stores).
- Sortierung: `ReceivedTime DESC` (neueste zuerst) erst NACH Voll-Sammeln aller Folder, dann OrderByDescending + Take(top). Pro-Folder-Sortierung wuerde aeltere Mails Top-Plaetze wegfressen.
- Top-Cap: 1-100, validated. `NextSkip = null` (Hard-Cap). Keine Pagination.
- Item-Strategy wie bei `list_mails`: `Items.Restrict(filter)` mit Fallback, `MapMailItem(item, includeBody:false)` (kein Body bei Listen), Class=43-Filter (olMail).
- Logging: START/DONE mit Statistiken (`visitedFolders`, `skippedFolders`, `fallbackFolders`, `skippedNonMail`, `failedItems`, `returned`) — fuer Smoke-Tests.

Alternativen (verworfen):
- `search_mails` um `[UnRead]`-Property-Filter erweitern: zwei orthogonale Achsen (Subject-Volltextsuche + Property-Filter) in einem Tool vermischt.
- `list_mails` um `recursive`-Parameter erweitern: macht Signatur uneindeutig.
- `list_unread_mails_recursive` als Spezialtool: zu eng.
- `AdvancedSearch` von Outlook: erfordert Folder-Scope, kein Hierarchie-Walk.
- `MAPITable` mit Restrict: liefert nur Spalten-Tupel.

Konsequenzen:
- Tool-Count: 25 -> 26.
- DTO: keine neuen Records noetig (`MailMessage` + `PagedResult<MailMessage>` reichen).
- Interface: `IOutlookService` und `IInteropOutlookAdapter` bekommen `ListMailsRecursiveAsync(scope, top, filter, ct)`.
- `WellKnownFolder` erweitert um `MailFolderNames` und `IsKnownMailFolder(name)`.
- Tests: +5 in `OutlookServiceTests` (Top-Range, invalid scope entry, empty scope entry, valid passthrough, null scope passthrough). FakeInteropAdapter hat `OnListMailsRecursiveAsync`-Hook. FakeOutlookService kompiliert mit leerer Default-Impl.
- Performance bei grossen Postfaechern: vollstaendiger Folder-Walk O(N). Bei Exchange-Konten Server-Side-Filter via `Items.Restrict`; bei lokalen PSTs Client-Iteration. Hard-Cap `top=100` sorgt fuer <= 100 Returns, aber Iteration kann bei 100k-Mail-Postfaechern Minuten dauern. Akzeptabel fuer manuellen Use-Case; nicht fuer Bulk.

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer/Tools/MailTools.cs](src/OutlookMcpServer/Tools/MailTools.cs) — Tool-Definition `ListMailsRecursive` mit `[Description(...)]`-Attributen auf jedem Parameter.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs) — `ListMailsRecursiveAsync` (Entry-Punkt), `CollectMailsFromFolderRecursive` (Rekursion), `CollectItemsInto` (per-Item-Mapping).
- [src/OutlookMcpServer.Domain/Models/Mail/MailModels.cs](src/OutlookMcpServer.Domain/Models/Mail/MailModels.cs) — `WellKnownFolder.MailFolderNames` + `IsKnownMailFolder(name)`.
- [src/OutlookMcpServer.Domain/Services/OutlookService.cs](src/OutlookMcpServer.Domain/Services/OutlookService.cs) — `ListMailsRecursiveAsync` mit Validation.
