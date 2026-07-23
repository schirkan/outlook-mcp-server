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

## 2026-07-22 — `fix recursive` (`f90aa05`)

Status: accepted

Kontext: `OutlookInteropAdapter.GetActiveItemAsync` und verwandte rekursive Active-Selection-Pfade hatten ein Edge-Case-Polymorphie-Problem in der Source-Generation-JSON-Serialisierung. Konkret: `ActiveMail` und `ActiveEvent` werden via `JsonDerivedType`-Attribute diskriminiert — die Polymorphie-Aufloesung im Source-Generator brauchte explizite Registrierung im `OutlookMcpJsonContext`, sonst fehlte die TypeInfoMetadata und Serialisierung scheiterte still.

Entscheidung: `OutlookMcpJsonContext` in `src/OutlookMcpServer.Domain/Serialization/OutlookMcpJsonContext.cs` erweitert um explizite Registrierung der polymophen Typen via `[JsonSerializable(typeof(...))]`. `OutlookInteropAdapter.GetActiveItemAsync` auf den CallSite-Pfad umgestellt (vorher Reflection, das bei `__ComObject`-RCW-Objekten nicht zuverlaessig funktionierte).

Konsequenzen:
- Polymorphe Serialisierung von `ActiveItem`-Varianten jetzt zuverlaessig (vorher: `NotSupportedException` bei `GetActiveItem` mit MailItem).
- Kleinere Refactorings in `OutlookInteropAdapter` (Reflection → CallSite, der von `6b053ed` etablierte Pfad).

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer.Domain/Serialization/OutlookMcpJsonContext.cs](src/OutlookMcpServer.Domain/Serialization/OutlookMcpJsonContext.cs) — explizite `[JsonSerializable(typeof(ActiveItem))]`-Registrierung.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs) — `GetActiveItemAsync` auf CallSite-Pfad umgestellt (statt Reflection-`GetType().GetProperty()`).

## 2026-07-22 — BodyFormat + Multi-Store Folder-Resolve + UnRead-Alias (`24e4e46`)

Status: accepted

Kontext: Drei Sub-Themen, alle am 2026-07-22 20:43 in einem Commit:

1. **BodyFormat-Enum + HtmlBodyConverter**: Vor diesem Commit hat der Server Outlook-HTML 1:1 durchgereicht (gleiches Problem fuer Mail UND Calendar). Bei LLM-Consumern fuehrte das zu Word/Outlook-Styling-Bloat (CSS-Inline, Class-Attribute, conditional comments), was Tokens verschwendet und die Lesbarkeit verschlechtert. HTML → Plain-Text ohne Struktur war auch nicht ideal (Listen, Tabellen, Code-Blocks gingen verloren).
2. **Multi-Store Folder-Resolve**: In Multi-Store-Profilen (Cached-Mode/Exchange + PST) liefert `session.GetDefaultFolder(olDefaultFolders.X)` nicht zwingend den kanonischen Folder. Beispiel: Ein Profil mit Cached Exchange + lokalem PST kann `olFolderInbox` auf den PST-Inbox mappen, nicht den Exchange-Inbox. Das fuehrte zu Scope-vs-Ordner-Mismatch bei `list_mails_recursive`: der User wollte Exchange-Inbox, bekam aber PST-Inbox, oder umgekehrt.
3. **UnRead-Alias-Normalisierung**: Outlook-DASL-Filter ist case-sensitive. User schrieben `[UnRead] = true` (PascalCase), Outlook akzeptiert aber nur `[Unread]` (lowercase 'r'). Resultat: `Items.Restrict("[UnRead] = true")` lieferte leeres Resultat, weil das Property `[UnRead]` nicht existiert — Outlooks echtes Property ist `[Unread]`. Still: die Filter fielen auf ungefilterte Iteration zurueck und gaben ALLE Mails zurueck, nicht nur ungelesene.

Entscheidung:
1. **BodyFormat**: Neuer interner Komponent `HtmlBodyConverter` (`src/OutlookMcpServer.Interop/HtmlBodyConverter.cs`) basiert auf ReverseMarkdown 3.7.0. Konvertiert Outlook-HTML zu GitHub-kompatibles Markdown. Neues Enum `BodyFormat` (`markdown`|`text`|`html`) in `CommonModels.cs` mit `BodyFormatExtensions.ParseBodyFormat`-Helper fuer case-insensitive String-Parsing (erlaubte Werte: `markdown`, `md`, `text`, `plain`, `plaintext`, `html`). Default fuer alle Tools = `markdown` (kompakt, gut lesbar fuer LLMs). Calendar verwendet denselben Konverter — verhindert das gleiche Styling-Bloat-Problem bei Termin-Body.
2. **Multi-Store**: Neuer Helper `ResolveDefaultFolderSmart(olId)` in `OutlookInteropAdapter`. Iteriert ueber `session.Stores[*].GetDefaultFolder(olId)` und nimmt den ersten Store, der die olId kennt. So bekommt man immer den kanonischen Default-Folder des ersten passenden Stores — was typischerweise der Cached-Exchange-Store ist.
3. **UnRead-Alias**: Server-seitige Normalisierung in `OutlookService.ListMailsRecursiveAsync`: vor dem `Items.Restrict(filter)`-Aufruf wird `[UnRead]` zu `[Unread]` ersetzt. Caller koennen beide Schreibweisen verwenden, der Server macht den Normalisierungs-Pass.

Alternativen (verworfen):
- BodyFormat: Eigenbau HTML-Parser (fehleranfaellig, viel Wartungsaufwand fuer Edge-Cases), HtmlAgilityPack + manuelle Konvertierung (zusaetzliche Dependency, weniger robust gegen Outlook-spezifische HTML-Idiosynkrasien), nur Plain-Text (Listen/Tabellen/Code-Blocks gehen verloren).
- Multi-Store: Konfigurierbarer Default-Store (zu frueh fuer v1, spaeter konfigurierbar machen in v1.1 wenn Use-Cases auftauchen).
- UnRead-Alias: Caller auf korrekte Schreibweise hinweisen (UX-schlecht, User merkt sich nicht welche korrekt ist), case-insensitive DASL-Property-Lookup (Outlook-intern nicht zuverlaessig).

Konsequenzen:
- Neuer interner File `src/OutlookMcpServer.Interop/HtmlBodyConverter.cs` (~91 LOC).
- Dependency: ReverseMarkdown 3.7.0 (NuGet).
- Calendar verwendet jetzt denselben Konverter wie Mail (vorher: HTML 1:1 durchgereicht, gleiche Word-Styling-Bloat-Problem).
- `list_mails_recursive` mit scope=[inbox,archive] funktioniert jetzt korrekt in Multi-Store-Profilen.
- API-DESIGN.md: expliziter Hinweis auf `[Unread]` (lowercase 'r', ACHTUNG: nicht `[UnRead]!`), Der Server normalisiert `UnRead` automatisch — beides funktioniert.
- `BodyFormat` und `string[]` als `[JsonSerializable(typeof(...))]` im `OutlookMcpJsonContext` registriert (Source-Gen-Polymorphie-Support).

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer/Tools/MailTools.cs](src/OutlookMcpServer/Tools/MailTools.cs) — `bodyFormat`-Parameter an `list_mails` + Logging der BodyFormat-Konvertierung.
- [src/OutlookMcpServer/Tools/CalendarTools.cs](src/OutlookMcpServer/Tools/CalendarTools.cs) — Calendar-Body via `HtmlBodyConverter.Convert()` statt 1:1 HTML durchreichen.
- [src/OutlookMcpServer.Interop/HtmlBodyConverter.cs](src/OutlookMcpServer.Interop/HtmlBodyConverter.cs) — neuer File, ~91 LOC, ReverseMarkdown-Wrapper.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs) — `ResolveDefaultFolderSmart`-Helper, `HtmlBodyConverter.Convert()`-Integration in `MapMailItem`/`MapAppointmentItem`, `TryMapMailRecipients`-Fix (siehe naechster Eintrag).
- [src/OutlookMcpServer.Domain/Models/Common/CommonModels.cs](src/OutlookMcpServer.Domain/Models/Common/CommonModels.cs) — `BodyFormat`-Enum + `BodyFormatExtensions`-Helper.
- [src/OutlookMcpServer.Domain/Services/OutlookService.cs](src/OutlookMcpServer.Domain/Services/OutlookService.cs) — UnRead-Alias-Normalisierung in `ListMailsRecursiveAsync`.
- [src/OutlookMcpServer.Domain/Serialization/OutlookMcpJsonContext.cs](src/OutlookMcpServer.Domain/Serialization/OutlookMcpJsonContext.cs) — `BodyFormat` + `string[]` Source-Gen-Registrierung.

## 2026-07-22 — `bodyFormat`-Tool-Parameter + Recipients-Type-Filter (`df11443`)

Status: accepted

Kontext: BodyFormat-Konvertierung (aus `24e4e46`) sollte nicht nur intern wirken, sondern auch vom MCP-Caller explizit steuerbar sein — verschiedene Use-Cases brauchen verschiedene Formate (LLM-Caller bevorzugen Markdown, Mobile-Clients vielleicht HTML, Text-Exporte Plain). Ausserdem: Outlook-MAPI-MailItems haben keine separaten `To`/`CC`/`BCC`-Properties — nur eine `Recipients`-Collection mit `.Type` pro Element (2=To, 3=CC, 4=BCC, 1=Originator). Der vorherige Versuch, `mail.To`/`mail.CC`/`mail.BCC` per IDispatch zu lesen, lieferte daher leere Arrays — Recipients-Sortierung ging verloren.

Entscheidung: `bodyFormat`-Parameter als expliziter Tool-Input an acht Tools exponiert: `list_mails`, `get_mail`, `get_mails`, `list_mails_recursive` (Mail); `get_event`, `list_events` (Calendar); `get_active_item`, `get_selected_items` (ActiveMail-Teil — wirkt nur auf Mails, nicht auf Termine). Default = `markdown` (kompakt, LLM-freundlich). `TryMapMailRecipients` liest einmal `mail.Recipients` und filtert nach `.Type` (2=To, 3=CC, 4=BCC).

Alternativen (verworfen):
- `bodyFormat` als Domain-DTO-Feld (in `MailMessage`/`CalendarEvent`): verwischt Layer-Trennung — Domain bleibt 1:1 zu Microsoft Graph (`ItemBody`-DTO), Format-Konvertierung ist Tool-Layer-Concern.
- Recipients-Sortierung via `mail.To`/`mail.CC`/`mail.BCC` Reflection: liefert leere Arrays (Reflection funktioniert nicht fuer `System.__ComObject`-RCW mit IDispatch-Properties — derselbe Bug wie in `6b053ed`).

Konsequenzen:
- Tool-Signaturen: 8 Tools bekommen zusaetzlichen `bodyFormat`-Parameter (string?, default null → Default `markdown`).
- Production-Default `markdown` weil LLM-Caller typischerweise Markdown-Body bevorzugen.
- `ToRecipients`/`CcRecipients`/`BccRecipients` in `MailMessage`-DTO jetzt korrekt befuellt (vorher: leer).
- Tool-Count bleibt **27** (keine neuen Tools, nur Parameter-Erweiterungen).
- API-DESIGN.md: `bodyFormat`-Parameter in 8 Tool-Sections dokumentiert + neuer Enum `BodyFormat`.

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer/Tools/MailTools.cs](src/OutlookMcpServer/Tools/MailTools.cs) — `bodyFormat`-Parameter an `list_mails`, `get_mail`, `get_mails`, `list_mails_recursive`.
- [src/OutlookMcpServer/Tools/CalendarTools.cs](src/OutlookMcpServer/Tools/CalendarTools.cs) — `bodyFormat`-Parameter an `list_events`, `get_event`.
- [src/OutlookMcpServer/Tools/ActiveSelectionTools.cs](src/OutlookMcpServer/Tools/ActiveSelectionTools.cs) — `bodyFormat`-Parameter an `get_active_item`, `get_selected_items` (nur Mail-Teil; ActiveEvent hat keinen Body).
- [src/OutlookMcpServer.Domain/Abstractions/IOutlookService.cs](src/OutlookMcpServer.Domain/Abstractions/IOutlookService.cs) — `bodyFormat`-Parameter auf 8 Service-Methoden.
- [src/OutlookMcpServer.Domain/Abstractions/IInteropOutlookAdapter.cs](src/OutlookMcpServer.Domain/Abstractions/IInteropOutlookAdapter.cs) — `bodyFormat`-Parameter auf 8 Interop-Methoden.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs) — `TryMapMailRecipients` liest einmal `mail.Recipients` und filtert nach `.Type` (2=To, 3=CC, 4=BCC).

## 2026-07-23 — list_mails-Stabilization (`45de587`)

Status: accepted

Kontext: Nach den BodyFormat- und Multi-Store-Features aus `24e4e46`/`df11443` zeigte sich beim Smoke-Test, dass `list_mails`/`list_mails_recursive` in manchen Profil-Konfigurationen nicht robust lief: (a) `ResolveDefaultFolderSmart` hatte Reihenfolge-Probleme — Store-Iteration vor `ns.GetDefaultFolder`, was bei Multi-Store-Profilen mit Cached-Mode/Exchange + PST je nach Store-Reihenfolge unterschiedliche Folder für dieselbe `olId` liefern konnte; (b) Notes/Journal/Tasks-Default-Folder wurden fälschlich als Mail-Folder interpretiert und lieferten `Items.Class != 43`; (c) DASL-Filter `[UnRead]` (mit großem R) wurde in manchen Outlook-Versionen stillschweigend abgelehnt und fiel auf ungefilterte Iteration zurück; (d) `ListMailsAsync`-Methode hatte inkonsistente `try/finally`-Struktur, sodass `Marshal.ReleaseComObject` in Fehlerpfaden teilweise fehlte.

Entscheidung:
1. **`ResolveDefaultFolderSmart` umstrukturiert**: Reihenfolge geändert — zuerst `ns.GetDefaultFolder(olId)`, dann Store-Iteration als Fallback (vorher umgekehrt). `ns.GetDefaultFolder` liefert im Normalfall den kanonischen Folder schneller; Store-Iteration ist nur Fallback für Profile, wo `ns.GetDefaultFolder` versagt.
2. **Neuer `ResolveDefaultFolderByOlIdWithSchemaCheck` + `IsMailFolder`-Heuristik**: iteriert über alle Stores und prüft jeden gefundenen Folder auf `DefaultItemType == 0` (= `olMailItem`). Verhindert, dass Notes/Journal/Tasks-Folder fälschlich als Mail-Folder interpretiert werden. Wird von `ListMailsRecursiveAsync` (OutlookService-Layer) für die `olFolderId`-Auflösung genutzt.
3. **Neuer `NormalizeDaslFilter` (adapter-side)**: Regex-basiertes `[UnRead]` → `[Unread]`-Replacement im Adapter. Redundant zum Service-Layer in `OutlookService.ListMailsRecursiveAsync`, damit auch direkte `list_mails`-Calls den Alias-Support erhalten.
4. **`ListMailsAsync` umstrukturiert**: konsistentes `try/finally` mit defensiven `Marshal.ReleaseComObject`-Aufrufen für `items` und `folder`, Error-Handling für `Restrict`-, `Sort`-, `Count`- und `Item(i+1)`-Operationen. `Sort` an `[ReceivedTime]` kann auf manchen Foldern `0x80020009` werfen (z. B. wenn der Folder kein `ReceivedTime` unterstützt) — wird jetzt mit `COMException`-Catch abgefangen und Fallback auf unsortierte Iteration gemacht.
5. **`ListMailsRecursiveAsync` (OutlookService-Layer)** nutzt jetzt `ResolveDefaultFolderByOlIdWithSchemaCheck` statt `ResolveDefaultFolderSmart` für die `olFolderId`-Auflösung — Schema-Check ist robuster als nur Folder-Smart-Resolve.

Alternativen (verworfen):
- `ResolveDefaultFolderSmart` komplett löschen, nur `ns.GetDefaultFolder` nutzen: verworfen, weil Multi-Store-Profile genau diesen Fallback brauchen.
- Schema-Check im Service-Layer (statt im Adapter): verworfen, weil der Adapter näher an der COM-Boundary ist und Schema-Information (`DefaultItemType`) COM-spezifisch ist.
- `Sort`-Aufruf ganz weglassen: verworfen, weil Caller explizit sortierte Ergebnisse erwarten (Pagination, Reihenfolge-Stabilität).

Konsequenzen:
- `OutlookInteropAdapter` umstrukturiert: `ResolveDefaultFolderSmart`, `ResolveDefaultFolderByOlIdWithSchemaCheck`, `IsMailFolder`, `NormalizeDaslFilter` neu, `ListMailsAsync` neu geschrieben.
- `OutlookService.ListMailsAsync` und `ListMailsRecursiveAsync`: ListMailsAsync nutzt jetzt `WellKnownFolder.IsKnownMailFolder(folderId)` um zu entscheiden, ob der adapter-seitige DASL-Alias-Filter angewendet wird.
- Robustheit: Notes-/Journal-/Tasks-Folder werden korrekt gefiltert, `[UnRead]`-Schreibweise wird toleriert, Sort-Failure wird graceful behandelt.
- Keine API-Änderungen (alle Änderungen sind intern).

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs) — `ResolveDefaultFolderSmart` umstrukturiert (Reihenfolge geändert), neuer `ResolveDefaultFolderByOlIdWithSchemaCheck`, neue `IsMailFolder`-Heuristik, neuer `NormalizeDaslFilter` (adapter-side DASL-Normalisierung), `ListMailsAsync` komplett umstrukturiert (konsistentes `try/finally`, defensives `Marshal.ReleaseComObject`, Error-Handling für Sort/Count/Item-Fetch), `ListMailsRecursiveAsync`-Aufruf nutzt jetzt `ResolveDefaultFolderByOlIdWithSchemaCheck` statt `ResolveDefaultFolderSmart`.
