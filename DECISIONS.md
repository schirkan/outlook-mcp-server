# DECISIONS — Outlook MCP Server

Architektur- und Designentscheidungen mit Datum + Begruendung. Wird iterativ erweitert.

Hinweis: Diese Datei enthaelt einen Wiederherstellungs-Stub mit Verweis auf den Code-Stand von 2026-07-21.
Der vollstaendige DECISIONS-Inhalt der vorherigen Sitzung (13 Eintraege) wird in einer separaten Sitzung nachgetragen.

## 2026-11-21 — Unsichere Fallbacks im Outlook-MCP-Vertrag entfernen

- Fehlgeschlagene Mailfilter sind Fehler, keine Freigabe für ungefilterte Ergebnisse.
- Die vorhandenen `Allow*`-Schalter werden für alle Schreibpfade durchgesetzt; `AllowCreate` deckt auch nicht löschende Änderungen ab. Einladungen, Updates mit Versand und Meeting-Antworten benötigen `AllowSend`.
- Kalenderzeiten werden auf UTC normalisiert; für IANA-Zonen und Outlooks lokale Datumsfilter bleibt Globalisierung im Publish-Profil eingeschaltet. Nicht zuverlässig steuerbare Absagen eigener Meetings werden abgelehnt statt als erfolgreich gemeldet.

## 2026-11-21 — stdio als einziger MCP-Transport

Der MCP-Client startet den Server lokal als Subprozess und kommuniziert über stdio. Das reicht für den Zugriff auf das lokale Outlook-Profil; der Server öffnet keinen Netzwerk-Listener und benötigt keine Transporteinstellung.

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

## 2026-07-24 — BodyFormat-Nullability + includeBody-Entfernung (`f40a4a8`)

Status: accepted

Kontext: BodyFormat-API bisher: Default = `markdown`, alle 8 Read-Tools hatten einen `bodyFormat`-Parameter. Zwei Probleme: (a) Listen-Tools wie `list_mails` lieferten Bodies automatisch mit — bei 100 Mails viel Token-Verbrauch (Outlook-HTML gross), und das Parsen/Serialisieren verschwendet CPU; (b) separater `includeBody`-Parameter war redundant — wer `includeBody=false` setzt, will keinen Body, also ist `bodyFormat=null` (= kein Body) die sauberere Loesung. Der neue Default `null` macht "kein Body" zum Standardverhalten und zwingt Caller, explizit `bodyFormat="markdown"` zu setzen wenn sie Bodies brauchen.

Entscheidung: Komplette bodyFormat-API-Umstellung:
- **Vier Tools OHNE `bodyFormat`**: `list_mails`, `list_mails_recursive`, `search_mails`, `list_events` — liefern nie einen Body (Header-only Listen-Calls).
- **Fuenf Tools MIT `bodyFormat` (nullable)**: `get_mail`, `get_mails`, `get_event`, `get_active_item`, `get_selected_items`. `null` = kein Body, sonst `markdown`/`text`/`html`.
- `includeBody`-Parameter komplett entfernt aus allen 6 betroffenen Tools (war redundant zu `bodyFormat=null`).
- `BodyFormatExtensions.ParseBodyFormat` gibt jetzt `BodyFormat?` zurueck (null bei null/empty statt Default `markdown`).

Alternativen (verworfen):
- `bodyFormat` als verbindliches Pflicht-Feld fuer Item-Tools: zwingt Caller zum Mitsenden, schlechtere DX.
- `includeBody` + `bodyFormat` als zwei separate Parameter beibehalten: redundante API.
- Default = `markdown` beibehalten: Performance-Problem bei grossen Listen bleibt.

Konsequenzen:
- 14 Files, +182/-149.
- `specs/API-DESIGN.md`, `README.md`, `PROJECT.md`, `DECISIONS.md` alle vier angepasst (Breaking Change → vollstaendige Doku-Sync-Pflicht).
- Tool-Count bleibt **27** (keine neuen Tools, nur Parameter-Umstellung).
- Performance-Gewinn bei `list_mails`-Calls: keine Body-Konvertierung mehr in 99% der Faelle.
- Token-Ersparnis: Outlook-Bodies koennen 5-50 KB pro Mail sein, bei Listen mit 100 Mails sind das 500 KB - 5 MB pro Call.

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer/Tools/MailTools.cs](src/OutlookMcpServer/Tools/MailTools.cs) — `bodyFormat`-Parameter aus `list_mails`/`list_mails_recursive` entfernt; `includeBody` aus `get_mail`/`get_mails` entfernt; `bodyFormat`-Default-Semantik auf `null` umgestellt.
- [src/OutlookMcpServer/Tools/CalendarTools.cs](src/OutlookMcpServer/Tools/CalendarTools.cs) — `bodyFormat`-Parameter aus `list_events` entfernt; `get_event` auf nullable `BodyFormat?` umgestellt.
- [src/OutlookMcpServer/Tools/ActiveSelectionTools.cs](src/OutlookMcpServer/Tools/ActiveSelectionTools.cs) — `get_active_item`/`get_selected_items` auf nullable `BodyFormat?` umgestellt (Default = kein Body).
- [src/OutlookMcpServer.Domain/Models/Common/CommonModels.cs](src/OutlookMcpServer.Domain/Models/Common/CommonModels.cs) — `BodyFormatExtensions.ParseBodyFormat` gibt jetzt `BodyFormat?` zurueck.
- [src/OutlookMcpServer.Domain/Abstractions/IOutlookService.cs](src/OutlookMcpServer.Domain/Abstractions/IOutlookService.cs) — Signaturen angepasst: `BodyFormat?`-Parameter auf 5 Methoden, kein Parameter auf 4 Methoden.
- [src/OutlookMcpServer.Domain/Abstractions/IInteropOutlookAdapter.cs](src/OutlookMcpServer.Domain/Abstractions/IInteropOutlookAdapter.cs) — gleiche Signatur-Anpassungen auf Adapter-Layer.
- [tests/OutlookMcpServer.Domain.Tests/Fakes/*](tests/OutlookMcpServer.Domain.Tests/Fakes/) — `FakeOutlookService` + `FakeInteropAdapter` an neue Signaturen angepasst.
- [tests/OutlookMcpServer.Domain.Tests/Tools/MailToolsTests.cs](tests/OutlookMcpServer.Domain.Tests/Tools/MailToolsTests.cs) — Tests an neue Signaturen angepasst (keine `includeBody`-Tests mehr, `bodyFormat`-Tests pruefen nullable-Verhalten).

## 2026-07-24 — Active-Mail EntryID-Reload (`e4b2591`)

Status: accepted

Kontext: `get_active_item` fuer `MailItem` hatte eine experimentelle WordEditor-basierte Body-Extraktion, die unzuverlaessige Ergebnisse lieferte (Word-Range-Objekte sind fluechtig, Inspector-Kontext kann waehrend des Reads kippen). Die anderen Mail-Lese-Tools (`get_mail`, `get_mails`) nutzen den `Namespace.GetItemFromID(EntryID)`-Pfad — stabil, getestet, konsistent.

Entscheidung: `OutlookInteropAdapter.GetActiveItemAsync` fuer Mail-Pfad umgestellt: aktives `MailItem.EntryID` extrahieren → via `Namespace.GetItemFromID(EntryID)` neu laden → dann wie `get_mail` mappen. WordEditor-Pfad entfernt. `get_selected_items` unveraendert (kein Inspector-Kontext dort, EntryID-Reload nicht noetig — Selection-Items sind bereits EntryID-stabil).

Alternativen (verworfen):
- WordEditor-Pfad stabilisieren: hoher Aufwand, marginaler Nutzen (Word-Editor-API ist nicht fuer Server-Use-Case gedacht).
- Auf Inspector-Item direkt arbeiten (ohne Reload): funktioniert in 90% der Faelle, aber die Edge-Cases (Inspector-Wechsel waehrend Reads, COM-RCW-Inkonsistenzen) sind genau die, die in Produktion weh tun.

Konsequenzen:
- Konsistentes Body-Verhalten zwischen `get_active_item` und `get_mail`/`get_mails`.
- Eliminiert eine Klasse von sporadischen Body-Lese-Fehlern.
- 1 File, +35/-1.

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs) — `GetActiveItemAsync` Mail-Pfad: WordEditor entfernt, `EntryID` + `Namespace.GetItemFromID`-Reload eingebaut. (Datei existiert nur bis `87e5896`; danach wandert die Logik in `OutlookInteropAdapter.ActiveSelection.cs`.)

## 2026-07-24 — `OutlookInteropAdapter`-Partial-Split (`87e5896`)

Status: accepted

Kontext: `OutlookInteropAdapter.cs` war mit 2971 Zeilen die groesste Klasse im Repo — ein "Gott-Objekt" mit Mail-Reads, Mail-Writes, Mail-Mapping, Calendar-Reads, Calendar-Writes, Calendar-Mapping, Active-Selection, Common-Helpers in einer Datei. Navigation, Refactoring, und Code-Reviews litten darunter; jede Aenderung touchte potentiell unrelated Code.

Entscheidung: Aufteilung in 8 thematische Partials via C# `partial class`-Mechanik:
- `OutlookInteropAdapter.ActiveSelection.cs` (~59 LOC) — `GetActiveItemAsync`, `GetSelectedItemsAsync`
- `OutlookInteropAdapter.Calendar.cs` (~201 LOC) — `ListCalendarsAsync`, `GetCalendarAsync`, `ListEventsAsync`, `GetEventAsync`
- `OutlookInteropAdapter.CalendarMapping.cs` (~369 LOC) — `MapAppointmentItem` + Sub-Helpers
- `OutlookInteropAdapter.CalendarMutations.cs` (~331 LOC) — `CreateEventAsync`, `UpdateEventAsync`, `DeleteEventAsync`, `RespondToEventAsync`, `FindMeetingTimesAsync`
- `OutlookInteropAdapter.Common.cs` (~132 LOC) — Shared Helpers (`EnsureOutlookInitialized`, `ReleaseComObject`, etc.)
- `OutlookInteropAdapter.Mail.cs` (~1001 LOC) — `ListMailsAsync`, `GetMailAsync`, `GetMailsAsync`, `ListMailsRecursiveAsync`, `SearchMailsAsync`
- `OutlookInteropAdapter.MailMapping.cs` (~316 LOC) — `MapMailItem` + Sub-Helpers (`TryMapMailRecipients`, etc.)
- `OutlookInteropAdapter.MailMutations.cs` (~326 LOC) — `SendMailAsync`, `CreateDraftAsync`, `UpdateMailAsync`, `MoveMailAsync`, `CopyMailAsync`, `DeleteMailAsync`, `GetMailHeadersAsync`, `ListAttachmentsAsync`, `GetAttachmentAsync`
- `OutlookInteropAdapter.cs` (Core-Stub, ~12 LOC) — Partial-Klassen-Deklaration + ctor + Fields

**Keine API-Aenderungen** — alle oeffentlichen Methoden via Partial-Mechanik erhalten, gleiche Signaturen, gleiche Semantik. Bestehende Tests bleiben unveraendert gruen.

Alternativen (verworfen):
- Komplett-Refactor in mehrere separate Klassen (Composition ueber Vererbung): groesserer API-Bruch, alle Caller muessten angepasst werden.
- Aufteilung nach Feature-Set (Reads vs. Mutations): orthogonal zur Themen-Aufteilung, wuerde Cross-Cutting-Concerns (z. B. `MapMailItem` fuer Reads + Mutations) duplizieren.

Konsequenzen:
- 9 Files, +2767/-2961.
- Bessere Navigation (jedes Partial ~60-1000 LOC statt 2971 in einem File).
- Compile-Zeiten leicht reduziert (Partials werden parallel kompiliert).
- Reviewer koennen thematische Aenderungen isoliert betrachten (z. B. nur `MailMutations` diff beim Review eines Mutation-Patches).
- Vorbereitung fuer weitere Sub-Aufteilungen (z. B. wenn `OutlookInteropAdapter.Mail.cs` mit 1001 LOC selbst noch zu gross wird).

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.cs) — Core-Stub (Partial-Klassen-Deklaration + ctor + Fields).
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.ActiveSelection.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.ActiveSelection.cs) — `GetActiveItemAsync`, `GetSelectedItemsAsync`.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.Calendar.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.Calendar.cs) — `ListCalendarsAsync`, `GetCalendarAsync`, `ListEventsAsync`, `GetEventAsync`.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.CalendarMapping.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.CalendarMapping.cs) — `MapAppointmentItem` + Sub-Helpers.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.CalendarMutations.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.CalendarMutations.cs) — `CreateEventAsync`, `UpdateEventAsync`, `DeleteEventAsync`, `RespondToEventAsync`, `FindMeetingTimesAsync`.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.Common.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.Common.cs) — Shared Helpers (`EnsureOutlookInitialized`, `ReleaseComObject`, etc.).
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.Mail.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.Mail.cs) — `ListMailsAsync`, `GetMailAsync`, `GetMailsAsync`, `ListMailsRecursiveAsync`, `SearchMailsAsync`.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.MailMapping.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.MailMapping.cs) — `MapMailItem` + Sub-Helpers (`TryMapMailRecipients`, etc.).
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.MailMutations.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.MailMutations.cs) - `SendMailAsync`, `CreateDraftAsync`, `UpdateMailAsync`, `MoveMailAsync`, `CopyMailAsync`, `DeleteMailAsync`, `GetMailHeadersAsync`, `ListAttachmentsAsync`, `GetAttachmentAsync`.

## 2026-07-28 — `open_item_by_id` UI-Tool (`6edbf78`)

Status: accepted

Kontext: Bisher waren alle Tools in der Active-Inspector/Selection-Sektion (`get_active_item`, `get_selected_items`) **read-only** auf UI-State — sie lesen `Application.ActiveInspector()`, `ActiveExplorer().Selection`, etc., öffnen oder verändern aber keine Fenster selbst. Das ist eine bewusste Design-Entscheidung (Sektion-Einleitung in `specs/API-DESIGN.md`: „Wichtig: alle diese Tools sind read-only auf den UI-State"). Aber ein häufiger Use-Case ist „öffne die Mail mit EntryID X" (nach `list_mails`/`search_mails`) oder „öffne den Termin mit EntryID Y" (nach `list_events`) — das ging bisher nicht via MCP-Tool. Caller mussten manuell in Outlook suchen oder via UI-Automation-Skripts arbeiten.

Entscheidung: Neues MCP-Tool `open_item_by_id(id)` in `ActiveSelectionTools.cs`. Architektur-Muster:
- **COM-Adapter** (`OutlookInteropAdapter.OpenItemByIdAsync` in `OutlookInteropAdapter.ActiveSelection.cs`): `GetOutlookApplicationAsync()` für Init → `RunComAsync(() => { dynamic item = GetMapiNamespace().GetItemFromID(id, Type.Missing); try { item.Display(false); return new OpenItemResult { Opened = true, Id = id }; } finally { ReleaseComObject(item); } }, nameof(OpenItemByIdAsync))`. Standard-Pattern `GetItemFromID + Item.Display(false) + ReleaseComObject im finally` — robust gegen Inspector-Switches und COM-RCW-Leaks.
- **Service** (`OutlookService.OpenItemByIdAsync`): Validation `ValidateStringNotEmpty(id, nameof(id))` → delegation an Adapter.
- **Domain-Model** (`OpenItemResult` in `CommonModels.cs`): `record { bool Opened, string Id }` — JSON-serialisierbar mit `[JsonPropertyName]`.
- **Tool** (`ActiveSelectionTools.OpenItemById`): `[McpServerTool(Name = "open_item_by_id")]` + `[Description("Oeffnet ein Outlook-Item per EntryID in einem Inspector-Fenster. Geeignet fuer MailItem und AppointmentItem, sofern Outlook den Typ anzeigen kann.")]`.
- **Section-Einleitung aktualisiert** (in `specs/API-DESIGN.md`): „`get_active_item` und `get_selected_items` sind read-only auf den UI-State. `open_item_by_id` darf dagegen explizit ein Inspector-Fenster öffnen."

Alternativen (verworfen):
- **COM-`Item.Display(true)` (modal)**: blockiert MCP-Server-Thread bis User das Fenster schließt — unbrauchbar für Server-Use-Case.
- **Tool in eigenes `UiTools.cs` auslagern statt in `ActiveSelectionTools`**: thematisch verwandt (Inspector-Operationen), aktuelle Co-Location mit `get_active_item`/`get_selected_items` ist semantisch passender.
- **Tool ohne `Opened`-Flag direkt `void` returnen**: weniger Feedback für Caller; `Opened=false` signalisiert klar „Inspector konnte nicht geöffnet werden".

Konsequenzen:
- 10 Files, +108/−1.
- Tool-Count: 27 → **28** (ActiveSelectionTools 2 → 3: `get_active_item`, `get_selected_items`, `open_item_by_id`).
- Martin hat README.md + specs/API-DESIGN.md direkt im Commit mit-aktualisiert (kein separater Doku-Commit nötig).
- Erste UI-Mutation in der bisher read-only Sektion — Section-Einleitung präzisiert.
- Validation greift früh (Service-Layer) → kein unnötiger COM-Round-Trip bei leerer/null ID.
- `Item.Display(false)` ist nicht-modal: das Inspector-Fenster öffnet sich, der MCP-Call returnt sofort mit `{opened: true, id: "..."}` — Caller kann weitere Tools aufrufen, während das Fenster offen ist.

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer/Tools/ActiveSelectionTools.cs](src/OutlookMcpServer/Tools/ActiveSelectionTools.cs) — neues `[McpServerTool(Name = "open_item_by_id")]` mit `OpenItemById(id, ct)` Methode (10 LOC).
- [src/OutlookMcpServer.Domain/Services/OutlookService.cs](src/OutlookMcpServer.Domain/Services/OutlookService.cs) — `OpenItemByIdAsync(id, ct)` mit `ValidateStringNotEmpty(id, nameof(id))` + Delegation an Adapter.
- [src/OutlookMcpServer.Domain/Abstractions/IOutlookService.cs](src/OutlookMcpServer.Domain/Abstractions/IOutlookService.cs) — Interface-Method `Task<OpenItemResult> OpenItemByIdAsync(string id, CancellationToken ct)`.
- [src/OutlookMcpServer.Domain/Abstractions/IInteropOutlookAdapter.cs](src/OutlookMcpServer.Domain/Abstractions/IInteropOutlookAdapter.cs) — gleiche Interface-Method auf Adapter-Layer.
- [src/OutlookMcpServer.Domain/Models/Common/CommonModels.cs](src/OutlookMcpServer.Domain/Models/Common/CommonModels.cs) — neuer `OpenItemResult { Opened, Id }` Record mit `[JsonPropertyName]` Attributen.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.ActiveSelection.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.ActiveSelection.cs) — COM-Impl: `GetItemFromID + Item.Display(false) + ReleaseComObject` Pattern, eingebettet in `RunComAsync` für einheitliches Error-Handling.
- [tests/OutlookMcpServer.Domain.Tests/Fakes/FakeInteropAdapter.cs](tests/OutlookMcpServer.Domain.Tests/Fakes/FakeInteropAdapter.cs) — Fake-Impl mit `_openByIdResults`-Dictionary für Test-Setup.
- [tests/OutlookMcpServer.Domain.Tests/Fakes/FakeOutlookService.cs](tests/OutlookMcpServer.Domain.Tests/Fakes/FakeOutlookService.cs) — Fake-Impl auf Service-Layer.
- [README.md](README.md) — neue Zeile in Active-Inspector/Selection-Sektion: „`open_item_by_id` — opens an Outlook item by EntryID in an Inspector window".
- [specs/API-DESIGN.md](specs/API-DESIGN.md) — neue Sektion `### open_item_by_id` mit Input/Output/Errors-Block; Section-Einleitung präzisiert (read-only-Klausel auf die anderen beiden Tools eingeschränkt).

## 2026-07-31 — OlDefaultFolder-IDs-Korrektur + WellKnownName-Population (`5254f19`)

Status: accepted (von Martin approviert, Author GitHub Copilot)

Kontext: Beim ersten lokalen Smoke-Test mit echtem Outlook-Profil (Outlook 2016 + Exchange-Mailbox) fiel auf, dass `ListMailFolders_ReturnsAtLeastInbox` und `GetMailFolder_Inbox_ReturnsValidFolder` rot waren. Die Service-Code-Pfade waren korrekt, aber die Folder-Resolution lieferte konsistent falsche oder leere Results. Ursache: `OlEnumMappings.ToOlDefaultFolder(string wellKnownName)` lieferte für jeden Well-Known-Namen die falsche `olFolderId`. Verifikation gegen Microsoft Learn / Outlook-PIA (`Microsoft.Office.Interop.Outlook.OlDefaultFolders`-Enum) zeigte: die ursprünglichen Mapping-Werte waren Copy-Paste-Error — kein Outlook-Versions-Drift (die Werte sind seit Outlook 2010 stabil). Konkret falsch: `Inbox 12 → 6`, `SentItems 18 → 5`, `Outbox 6 → 4`, `DeletedItems 5 → 3`, `JunkEmail 16 → 23`, `Drafts 11 → 16`. Plus: `Archive` hat gar keine offizielle `OlDefaultFolders`-Konstante (Outlook stellt keinen globalen Archive-Default bereit) — der ursprüngliche Wert 23 war in Wirklichkeit `olFolderJunk`.

Zusätzliche Folgeprobleme:
1. `ListMailFoldersAsync` iterierte `ns.Folders` (Postfächer/Stores selbst, nicht die sichtbaren Mail-Ordner). In Exchange-Profilen lieferte das eine Liste von Postfach-Namen wie `"Martin.Pietschmann@itergo.com"` ohne sichtbare Mail-Ordner — komplett falsche Semantik für einen Caller der „alle Mail-Ordner" will.
2. `MapMailFolder` setzte `WellKnownName` immer auf `null`, sodass Caller Folder nur per `displayName` (lokalisierungsabhängig!) oder `id` identifizieren konnten — fragil gegen Outlook-Sprachvarianten.
3. `GetMailFolderAsync` setzte `WellKnownName` auch nicht, wenn der Caller einen Well-Known-Namen statt EntryID übergab.

Entscheidung:
- **OlDefaultFolder-IDs korrigiert** in `OlEnumMappings.ToOlDefaultFolder` (siehe Mapping-Tabelle oben).
- **`Archive` aus dem Mapping entfernt** (`null` zurückgeben). Caller müssen expliziten Folder-Pfad verwenden — dokumentiert in README.md und specs/API-DESIGN.md (Hinweis-Box).
- **Neue `KnownFolderMappings`-Liste** in `OlEnumMappings` (`IReadOnlyList<(string Name, int OlFolderId)>`) für Service-/Adapter-Iteration. Quelle-of-Truth für `ListMailFoldersAsync`-Vergleich.
- **`ListMailFoldersAsync` umstrukturiert**: sammelt vorab WellKnown-COM-Referenzen via `Namespace.GetDefaultFolder(olId)`, baut `Dictionary<string, string>` (EntryID → WellKnownName, `OrdinalIgnoreCase`), iteriert dann `ns.Folders` (Postfächer) → pro Store `storeRoot.Folders` (sichtbare Mail-Ordner). Jeder gefundene Folder wird per `Dictionary.TryGetValue` mit WellKnownName befüllt. COMException pro WellKnown-Lookup wird mit `LogDebug` toleriert (nicht alle Profile haben jeden WellKnown-Folder — z. B. Exchange-Cached-Mode ohne Junk).
- **`MapMailFolder` erweitert**: optionaler `wellKnownName`-Parameter (`null` Default), wird im DTO gesetzt. Callers müssen beim Aufruf den WellKnownName mitgeben.
- **`GetMailFolderAsync` erweitert**: `wellKnown = WellKnownFolder.IsKnownMailFolder(folderId) ? folderId : null` — wenn Caller einen Well-Known-Namen übergibt, wird `WellKnownName` direkt gesetzt; bei EntryID-Übergabe bleibt es `null` (Lookup wäre zu teuer für Single-Folder-Calls).
- **Integration-Tests grün**: `ListMailFolders_ReturnsAtLeastInbox`, `GetMailFolder_Inbox_ReturnsValidFolder` (Outlook 2016 + Exchange). Full Test-Suite jetzt **69/69 passing** (63 domain + 6 integration).

Alternativen (verworfen):
- **Mapping-Werte aus anderer Quelle beziehen** (z. B. NuGet-Package, externe Lib): keine zuverlässige, gewartete Quelle gefunden; Werte direkt im Code mit Microsoft-Learn-Verifikation ist transparenter.
- **Cache der WellKnown-COM-Referenzen** über mehrere Calls hinweg: würde Init-Logik komplexer machen, ohne messbaren Performance-Vorteil (Outlook-Cache ist schnell genug für typische Call-Frequenz).
- **`WellKnownName` per DisplayName-Vergleich setzen** (z. B. `"Inbox"` → `"inbox"`): sprachabhängig — `"Inbox"` heißt auf deutschem Outlook `"Posteingang"`, französisch `"Boîte de réception"`. EntryID-Vergleich ist die einzige locale-unabhängige Methode.

Konsequenzen:
- 3 Files, +146/−30.
- **Keine API-Changes** — alle Signaturen bleiben gleich (Tool-Surface unverändert).
- **Verhaltensänderungen ohne API-Change**:
  - `list_mail_folders` (ohne `parentFolderId`) liefert jetzt sichtbare Top-Level-Mailordner aller Postfächer (`Inbox`, `Drafts`, ...) — vorher lieferte es eine Liste von Postfächern.
  - `list_mail_folders`/`get_mail_folder`: `WellKnownName` ist jetzt zuverlässig befüllt (vorher immer `null`).
  - `list_mails`/`get_mail_folder` mit `folderId="archive"`: liefert jetzt einen Fehler statt falschen Folder (kein offizielles OlDefaultFolder für Archive). Caller müssen expliziten Folder-Pfad verwenden.
- **Integration-Tests grün** — erster automatisierter End-to-End-Beweis dass das Toolset gegen echtes Outlook funktioniert. Karte 7 Acceptance nähert sich (`lokal grün` durch automatisierte Integration-Tests bestätigt; Martin's manuelle Verifikation steht weiterhin aus).
- Author = **GitHub Copilot** (von Martin approviert); Commit-Message dokumentiert Microsoft-Learn-Verifikation explizit.
- Tool-Count bleibt **28** (keine neuen Tools).
- Pattern `GetItemFromID` + EntryID-Match für locale-unabhängige Identifikation ist auch für andere Outlook-Entities anwendbar (Categories, Views, Stores).

Vollstaendige Begruendung pro Code-Stelle:
- [src/OutlookMcpServer.Interop/EnumMappings.cs](src/OutlookMcpServer.Interop/EnumMappings.cs) — `ToOlDefaultFolder` korrigiert (6 Mapping-Werte), `Archive => null`, neue `KnownFolderMappings`-Liste als `IReadOnlyList<(string Name, int OlFolderId)>`, Inline-Kommentar mit Microsoft-Learn-Referenz.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.Mail.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.Mail.cs) — `ListMailFoldersAsync` komplett umstrukturiert: pre-collection der WellKnown-COM-Referenzen mit COMException-Toleranz, Iteration über `storeRoot.Folders` pro Postfach, `wellKnownByEntryId.TryGetValue` für `WellKnownName`-Population. `GetMailFolderAsync` setzt `wellKnown` wenn Caller Well-Known-Namen übergibt.
- [src/OutlookMcpServer.Interop/OutlookInteropAdapter.MailMapping.cs](src/OutlookMcpServer.Interop/OutlookInteropAdapter.MailMapping.cs) — `MapMailFolder` bekommt optionalen `wellKnownName`-Parameter (Default `null`), DTO-`WellKnownName` wird aus diesem Parameter gesetzt statt hardcoded `null`.
- [tests/OutlookMcpServer.IntegrationTests/MailFolderIntegrationTests.cs](tests/OutlookMcpServer.IntegrationTests/MailFolderIntegrationTests.cs) — `ListMailFolders_ReturnsAtLeastInbox`, `GetMailFolder_Inbox_ReturnsValidFolder` jetzt grün (Outlook 2016 + Exchange); Full Test-Suite 69/69 passing (laut Commit-Message, ohne Diff-Beleg — Integration-Tests wurden evtl. in einer vorherigen Session angepasst).
