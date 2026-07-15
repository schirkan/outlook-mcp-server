# API-Design — MCP-Tools & Resources

Alle Tools folgen der Naming-Convention von Microsoft Graph (Mail: `messages`, `mailFolders`, Calendar: `events`, `calendars`), sind aber technisch unabhängig (COM-Interop statt Graph/HTTP). Property-Namen sind 1:1 zu Microsoft Graph, damit Graph-Client-Code mit minimalen Anpassungen auch gegen diesen Server läuft.

Tool-Eingaben sind JSON-Objekte; Tool-Ausgaben sind JSON-Objekte. Datums-/Zeit-Werte: ISO 8601 mit Time-Zone (z. B. `2026-07-15T14:00:00` + `timeZone: "Europe/Berlin"`). UTC-Werte mit `Z`-Suffix.

## Mail

### `listMailFolders`

**Beschreibung**: Listet alle Mail-Ordner im Standard-Postfach auf (Inbox, Drafts, SentItems, DeletedItems + Custom-Ordner).

**Input**:
- `parentFolderId` (optional, string): EntryID eines Parent-Ordners. Default = Root.
- `includeHidden` (optional, bool): Default `false`.

**Output**:
```json
{
  "value": [
    {
      "id": "00000000...EntryID",
      "displayName": "Inbox",
      "wellKnownName": "inbox",
      "parentFolderId": "00000000...",
      "childFolderCount": 3,
      "totalItemCount": 1234,
      "unreadItemCount": 7
    }
  ]
}
```

### `getMailFolder`

**Input**: `folderId` (string, required)
**Output**: einzelner Folder (gleiche DTO wie oben).

### `listMails`

**Beschreibung**: Listet Mails in einem Ordner. Sortierung: neueste zuerst (ReceivedTime desc).

**Input**:
- `folderId` (string, required) — oder well-known name (`inbox`, `drafts`, `sentItems`, `deletedItems`, `junkEmail`, `archive`, `outbox`)
- `top` (int, default 25, max 100)
- `skip` (int, default 0)
- `filter` (optional, string): Outlook-DASL-Filter (z. B. `"@SQL=urn:schemas:httpmail:subject LIKE '%urgent%'"`)
- `search` (optional, string): Volltext-Quicksearch über Subject + BodyPreview (Outlook-InstanteSearch)

**Output**:
```json
{
  "value": [
    {
      "id": "00000000...EntryID",
      "conversationId": "...",
      "subject": "Re: Projekt-Status",
      "bodyPreview": "Hi Martin, anbei der Status...",
      "from": { "emailAddress": { "name": "Chef", "address": "chef@firma.de" } },
      "toRecipients": [{ "emailAddress": { "name": "Martin", "address": "martin@firma.de" } }],
      "receivedDateTime": "2026-07-15T10:23:00Z",
      "sentDateTime": "2026-07-15T10:22:47Z",
      "hasAttachments": true,
      "importance": "high",
      "isRead": false,
      "categories": ["Projekt-X"]
    }
  ],
  "nextSkip": 25
}
```

### `getMail`

**Input**: `id` (string, required), `includeBody` (bool, default true)
**Output**:
```json
{
  "id": "...",
  "subject": "...",
  "body": { "contentType": "html", "content": "..." },
  "from": { "emailAddress": { "name": "...", "address": "..." } },
  "toRecipients": [...],
  "ccRecipients": [...],
  "bccRecipients": [...],
  "receivedDateTime": "...",
  "sentDateTime": "...",
  "hasAttachments": true,
  "importance": "normal",
  "isRead": true,
  "categories": [],
  "conversationId": "..."
}
```

### `getMailHeaders`

**Input**: `id` (string, required)
**Output**: `{ "internetMessageHeaders": [{ "name": "X-Mailer", "value": "..." }, ...] }`
Implementierung: `MailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")` (PR_TRANSPORT_MESSAGE_HEADERS).

### `listAttachments`

**Input**: `mailId` (string, required)
**Output**: `[{ "id": "...", "name": "report.pdf", "contentType": "application/pdf", "size": 12345, "isInline": false }]`

### `getAttachment`

**Input**: `mailId` (string, required), `attachmentId` (string, required)
**Output**: `{ "id": "...", "name": "...", "contentType": "...", "size": 12345, "contentBase64": "..." }`
Limit: `MaxAttachmentBytes` aus Config.

### `sendMail`

**Beschreibung**: Erstellt und versendet eine neue Mail. Bei `replyToId` wird der Reply/ReplyAll/Forward erzeugt.

**Input**:
```json
{
  "to": ["empfaenger@firma.de"],
  "cc": [],
  "bcc": [],
  "subject": "Statusupdate",
  "body": { "contentType": "text", "content": "Hi, ..." },
  "importance": "normal",
  "attachments": [
    { "name": "report.pdf", "contentBase64": "...", "contentType": "application/pdf" }
  ],
  "replyToId": null,
  "replyAll": false,
  "forwardFromId": null,
  "sendAt": null,
  "saveToSentItems": true
}
```

**Output**: `{ "sent": true, "id": "EntryID-..." }` (id der gesendeten Mail in SentItems)

### `createDraft`

**Input**: wie `sendMail` ohne `saveToSentItems`, plus optional `replyToId`
**Output**: `{ "id": "EntryID-..." }`

### `updateMail`

**Input**: `id` (string, required), beliebige patchbare Felder (`isRead`, `categories`, `importance`)
**Output**: `{ "id": "..." }`

### `moveMail` / `copyMail`

**Input**: `id`, `destinationFolderId`
**Output**: `{ "newId": "..." }` (EntryID im neuen Folder; ID kann sich ändern — wie bei Graph)

### `deleteMail`

**Input**: `id`, `permanent` (bool, default false — false = in DeletedItems, true = hart löschen)
**Output**: `{ "deleted": true }`

### `searchMails`

**Input**: `query` (string, required), `folderId` (optional), `top` (default 25)
**Output**: wie `listMails` (Advanced-Filter: DASL, alle Ordner wenn `folderId=null`)

## Kalender

### `listCalendars`

**Input**: (keine)
**Output**:
```json
{
  "value": [
    {
      "id": "00000000...EntryID-Cal",
      "name": "Kalender",
      "isDefaultCalendar": true,
      "canEdit": true,
      "owner": "martin@firma.de"
    }
  ]
}
```

### `getCalendar`

**Input**: `id`
**Output**: einzelner Calendar

### `listEvents` (CalendarView)

**Input**:
- `calendarId` (string, optional — default = default calendar)
- `startDateTime` (DateTimeTimeZone, required)
- `endDateTime` (DateTimeTimeZone, required)
- `top` (int, default 50, max 250)
- `skip` (int, default 0)
- `filter` (optional, string): z. B. `"showAs eq 'busy'"` (einfache Equality-Filter, intern auf DASL gemappt)

**Output**:
```json
{
  "value": [
    {
      "id": "00000000...EntryID",
      "subject": "Team-Meeting",
      "bodyPreview": "Agenda: Status, Roadmap",
      "start": { "dateTime": "2026-07-16T14:00:00", "timeZone": "Europe/Berlin" },
      "end":   { "dateTime": "2026-07-16T15:00:00", "timeZone": "Europe/Berlin" },
      "isAllDay": false,
      "location": { "displayName": "Konferenzraum 3" },
      "organizer": { "emailAddress": { "name": "Chef", "address": "chef@firma.de" } },
      "attendees": [
        { "emailAddress": { "name": "Alice", "address": "alice@firma.de" }, "type": "required", "status": { "response": "accepted" } }
      ],
      "importance": "normal",
      "sensitivity": "normal",
      "showAs": "busy",
      "isCancelled": false,
      "isReminderOn": true,
      "reminderMinutesBeforeStart": 15,
      "categories": []
    }
  ],
  "nextSkip": 50
}
```

### `getEvent`

**Input**: `id`
**Output**: einzelnes Event (Properties wie oben + `body`, `hasAttachments`, `recurrence`)

### `createEvent`

**Input**:
```json
{
  "subject": "Team-Meeting",
  "body": { "contentType": "html", "content": "<b>Agenda</b>: ..." },
  "start": { "dateTime": "2026-07-17T14:00:00", "timeZone": "Europe/Berlin" },
  "end":   { "dateTime": "2026-07-17T15:00:00", "timeZone": "Europe/Berlin" },
  "isAllDay": false,
  "location": "Konferenzraum 3",
  "attendees": [
    { "email": "alice@firma.de", "name": "Alice", "type": "required" },
    { "email": "bob@firma.de",   "name": "Bob",   "type": "optional" }
  ],
  "reminderMinutesBeforeStart": 15,
  "categories": ["Team"],
  "showAs": "busy",
  "importance": "normal",
  "sensitivity": "normal",
  "sendInvitations": true
}
```

**Output**: `{ "id": "EntryID-..." }` (gesendete Termineinladung an alle Attendees, wenn `sendInvitations=true`)

### `updateEvent` (PATCH)

**Input**: `id` + patchbare Felder. Mit `sendUpdate=true` werden Attendees benachrichtigt. `forceUpdateToAllAttendees` setzt das MAPI-Flag für „Update an alle erzwingen" (Outlook-Standard: nur Delta).

### `deleteEvent`

**Input**: `id`, `sendCancellation` (bool, default true)

### `respondToEvent` (Accept / Tentative / Decline)

**Input**: `id`, `response` (`accepted|tentativelyAccepted|declined`), `comment` (optional, wird in die Antwort-Mail an den Organizer übernommen)
**Output**: `{ "ok": true }`
Implementierung: `MeetingItem.Respond(OlResponseStatus, true, true)` (letzter Parameter = keine Mail an Organizer wenn false; wir wollen true → Antwort an Organizer).

### `findMeetingTimes` (free/busy, Self only in v1)

**Input**:
- `durationMinutes` (int)
- `timeWindow`: { `start`: DateTimeTimeZone, `end`: DateTimeTimeZone } (required)
- `maxCandidates` (int, default 10)

**Output**:
```json
{
  "value": [
    { "start": "...", "end": "...", "confidence": 100 }
  ]
}
```
Implementierung: iteriere `Items.Find` über den Self-Calendar im Zeitfenster, berechne Lücken >= `durationMinutes`. v1: Self only. v1.1: optional auch über shared/delegated Kalender (via `Recipient.FreeBusy`).

## Active-Inspector / Selection (Outlook-UI-State, COM-only)

Diese Tools sind ein **Alleinstellungsmerkmal von COM-Interop gegenüber Microsoft Graph** — Graph ist Server-seitig und hat keinen Zugriff auf den lokalen Outlook-UI-State. Die COM-Properties `Application.ActiveInspector()`, `Inspector.CurrentItem`, `Application.ActiveExplorer()` und `Explorer.Selection` liefern genau diesen State.

Wichtig: alle diese Tools sind **read-only auf den UI-State**. Sie öffnen oder verändern keine Fenster selbst.

### `getActiveItem`

**Beschreibung**: Liefert das Item, das aktuell im Vordergrund-Inspector-Fenster offen ist (Doppelklick auf eine Mail / einen Termin). Ein typischer Use-Case: „Fasse die Mail zusammen, die ich gerade offen habe" oder „lege einen Folgetermin an, basierend auf dem Termin, der gerade offen ist".

**Input**: (keine)

**Output** (polymorph, diskriminiert via `kind`):
```json
{ "kind": "mail",  "item": { ...MailMessage... } }
| { "kind": "event", "item": { ...CalendarEvent... } }
| null
```

**Errors**: keine — kein Inspector offen → `null` (statt Exception). Inspektor, der ein Item öffnet, das wir nicht kennen (z. B. ein leeres `MailItem` im Entwurfsmodus ohne Empfänger) → wird wie `null` behandelt und geloggt.

**Implementierung**:
```csharp
var insp = Application.ActiveInspector();   // ggf. null
if (insp is null) return null;
var current = insp.CurrentItem;            // _MailItem | _AppointmentItem | _ContactItem | _TaskItem | ...
return current switch
{
    MailItem mi        => new ActiveItem { Kind = "mail",  Item = MapToMail(mi) },
    AppointmentItem ai => new ActiveItem { Kind = "event", Item = MapToEvent(ai) },
    _                  => null  // für Tasks/Contacts/Notes (V1.1-Backlog)
};
```

**v1-Scope**: `mail` (gelesen oder im Entwurfs-/Edit-Modus), `event` (gelesen oder in Bearbeitung). Tasks/Contacts sind v1.1.

### `getSelectedItems`

**Beschreibung**: Liefert alle Items, die im aktuell aktiven Explorer-Fenster markiert sind (Posteingang mit 3 Mails angeklickt, Kalender-View mit 2 Terminen markiert, etc.). Klassischer Bulk-Workflow: „Lösch die 12 markierten Spam-Mails" oder „Erstelle Follow-up-Termine zu allen markierten Mails".

**Input**:
- `scope` (optional, string, default `"any"`): `"mail"` | `"calendar"` | `"any"` — filtert auf Item-Typen.
- `top` (optional, int, default 50, max 250): harte Obergrenze, schützt vor zu großen Selection-Batches.

**Output**:
```json
{
  "value": [
    { "kind": "mail",  "item": { ...MailMessage... } },
    { "kind": "event", "item": { ...CalendarEvent... } }
  ],
  "count": 3
}
```

**Errors**:
- `OutlookNotActive` — wenn `Application.ActiveExplorer()` `null` ist (z. B. nur ein Inspector offen, kein Explorer-Fenster)
- `InvalidInput` — wenn `scope` kein gültiger Wert ist oder `top` außerhalb `[1, 250]`

**Implementierung**:
```csharp
var explorer = Application.ActiveExplorer();   // ggf. null
if (explorer is null) throw OutlookServiceException(OutlookNotActive, "Kein Explorer aktiv");
var selection = explorer.Selection;           // _Selection-Collection
var items = new List<ActiveItem>();
for (int i = 1; i <= Math.Min(selection.Count, top); i++) {
    var sel = selection[i];
    switch (sel) {
        case MailItem mi:        if (scope is "any" or "mail")     items.Add(MapMail(mi)); break;
        case AppointmentItem ai: if (scope is "any" or "calendar") items.Add(MapEvent(ai)); break;
    }
}
return items;
```

**Edge Cases**:
- `Selection.Count == 0` → `value: []`, `count: 0`, **kein** Fehler (valide Selection, einfach nichts markiert)
- Selection im Suchordner oder RSS — wird wie normale Items behandelt (Outlook-eigenes Verhalten, kein Sonderfall)
- Mixed Selection (technisch nicht möglich im Standard-Explorer, da Folder nur einen Typ führt) — wir werfen keinen Fehler, sondern filtern per `scope`

## Adressbuch-Auflösung (Exchange GAL, COM-only)

Löst **Namen in SMTP-Mailadressen auf** via Exchange-GAL. COM-only — Graph kennt dafür `/users?$search` und `/me/people`, aber Cloud-only; der Outlook-COM-Adapter löst rein lokal über das geladene Outlook-Profil.

**Use-Case**: "Schicke eine Mail an Martin" (Display-Name → SMTP), "Prüfe ob ein Name gültig ist", "Löse einen Verteiler zur Anzeige auf".

### `resolveName`

**Input**:
- `query` (string, required, MinLength 1): Suchbegriff — Substring-Match auf `ExchangeUser.Name` ODER `PrimarySmtpAddress` (case-insensitive)
- `top` (int, default 10, max 50): harte Obergrenze gegen zu große GAL-Trefferlisten

**Output**: `PagedResult<ResolvedRecipient>` mit Properties:
- `DisplayName` (string)
- `EmailAddress` (string — der **resolved** SMTP-Primary, nicht der Anzeigename)
- `Type` (`User|Group|Room|Other`) — Enum
- `JobTitle` (string?, optional)
- `Department` (string?, optional)
- `OfficeLocation` (string?, optional)
- `Alias` (string?, optional, Exchange-Alias)

**Errors**:
- `InvalidInput` wenn `query` leer oder `top` außerhalb [1, 50]
- `OutlookNotRunning` wenn Session nicht initialisiert

**Implementierung** (COM):
- `Application.Session.GetGlobalAddressList()` → `AddressList`
- `addressList.AddressEntries` iterieren via `GetFirst()` / `GetNext()`
- pro `AddressEntry`:
  - `entry.GetExchangeUser()` (kann `null` sein bei Verteilerlisten/Räumen) → `entry.GetExchangeDistributionList()` für `Type=Group`
  - Match auf `Name.Contains(query, IgnoreCase)` ODER `PrimarySmtpAddress.Contains(query, IgnoreCase)`
- Bei `== top` abbrechen, `Marshal.ReleaseComObject` für jedes COM-Objekt im finally

**Edge Cases**:
- Verteilerlisten → `Type=Group`, Members NICHT rekursiv aufgelöst (Use-Case nur Anzeige)
- Room-Adressen → `Type=Room`
- `query="*"` → alle GAL-Einträge (durch `top` gecappt)
- `query=""` → `InvalidInput`

**Was NICHT in v1 ist** (per DECISIONS.md):
- Adressbuch-CRUD (Create/Update/Delete) — bleibt v1.1
- `listAddressBooks()` (Multi-Buch-Browser) — Martin hat nur EIN Tool gewünscht
- `showSelectNamesDialog` (UI-getrieben, für MCP-Clients schwer darstellbar)

### Polymorphe DTO-Serialisierung (`ActiveItem`)

`ActiveItem` ist ein `abstract record` mit `JsonDerivedType`-Attributen (STJ, ab .NET 7) — Diskriminator ist `kind`:

```csharp
[JsonDerivedType(typeof(ActiveMail),  "mail")]
[JsonDerivedType(typeof(ActiveEvent), "event")]
public abstract record ActiveItem
{
    [JsonPropertyName("kind")]
    public abstract string Kind { get; init; }
}

public sealed record ActiveMail : ActiveItem
{
    public override string Kind { get; init; } = "mail";
    [JsonPropertyName("item")] public required MailMessage Item { get; init; }
}

public sealed record ActiveEvent : ActiveItem
{
    public override string Kind { get; init; } = "event";
    [JsonPropertyName("item")] public required CalendarEvent Item { get; init; }
}

public sealed record ActiveMail : ActiveItem
{
    public override string Kind { get; init; } = "mail";
    [JsonPropertyName("item")] public required MailMessage Item { get; init; }
}

public sealed record ActiveEvent : ActiveItem
{
    public override string Kind { get; init; } = "event";
    [JsonPropertyName("item")] public required CalendarEvent Item { get; init; }
}
```

Der MCP-Client braucht keine Type-Hints — `kind` im Response reicht zum Dispatchen.

## Mapping: Microsoft Graph → OutlookMcpServer

| Graph Endpoint | MCP-Tool | Interop-Calls |
|---|---|---|
| `GET /me/mailFolders` | `listMailFolders` | `Namespace.Folders` traversal |
| `GET /me/mailFolders/{id}` | `getMailFolder` | `Namespace.GetFolderFromID` |
| `GET /me/mailFolders/{id}/messages` | `listMails` | `Folder.Items.Restrict` + Sort |
| `GET /me/messages/{id}` | `getMail` | `Namespace.GetItemFromID` |
| `GET /me/messages/{id}/attachments` | `listAttachments` | `MailItem.Attachments` |
| `POST /me/sendMail` | `sendMail` | `Application.CreateItem(olMailItem)` → `.Send()` |
| `POST /me/messages` (draft) | `createDraft` | `Application.CreateItem(olMailItem)` → `.Save()` |
| `PATCH /me/messages/{id}` | `updateMail` | `MailItem.PropertyAccessor` + speichern |
| `POST /me/messages/{id}/move` | `moveMail` | `MailItem.Move(destFolder)` |
| `DELETE /me/messages/{id}` | `deleteMail` | `MailItem.Delete()` |
| `GET /me/events` | `listEvents` | `Calendar.Items.Find("[Start] >= ...")` + Restrict |
| `GET /me/events/{id}` | `getEvent` | `Namespace.GetItemFromID` |
| `POST /me/events` | `createEvent` | `Calendar.Items.Add(olAppointmentItem)` → `.Send()` |
| `PATCH /me/events/{id}` | `updateEvent` | `AppointmentItem` mutieren → `.Save()` / `.Send()` |
| `DELETE /me/events/{id}` | `deleteEvent` | `AppointmentItem.Delete()` |
| `POST /me/events/{id}/accept` | `respondToEvent` (accepted) | `MeetingItem.Respond(olResponseAccept, ...)` |
| — | `getActiveItem` | `Application.ActiveInspector()?.CurrentItem` |
| — | `getSelectedItems` | `Application.ActiveExplorer()?.Selection` |
| — | `resolveName` | `Session.GetGlobalAddressList()` + `AddressEntries.GetExchangeUser()` |

## Fehler-Schema (einheitlich für alle Tools)

Bei Fehlern gibt das Tool ein JSON-Objekt mit `isError=true` zurück (MCP-Standard) und folgendem Body:

```json
{
  "error": {
    "code": "FolderNotFound",
    "message": "Folder 'InboxX' not found",
    "details": null
  }
}
```

### Fehler-Codes

| Code | Bedeutung | Retry? |
|---|---|---|
| `FolderNotFound` | `mailFolder` ID existiert nicht | nein |
| `MailNotFound` | `message` ID existiert nicht | nein |
| `EventNotFound` | `event` ID existiert nicht | nein |
| `CalendarNotFound` | `calendar` ID existiert nicht | nein |
| `AttachmentNotFound` | `attachment` ID existiert nicht | nein |
| `InvalidInput` | Input-Validierung fehlgeschlagen (fehlendes Feld, ungültiges Format) | nein (nach Korrektur) |
| `OutlookNotRunning` | Outlook-Prozess nicht verfügbar + `AutoStartOutlook=false` | ja (nach Start) |
| `OutlookBusy` | COM-Call wegen „Another operation in progress" abgewiesen | ja (mit Backoff) |
| `PermissionDenied` | COM-Security-Block (z. B. Anti-Virus blockiert In-Process-COM) | nein |
| `AttachmentTooLarge` | > `MaxAttachmentBytes` | nein (nach Verkleinern) |
| `SendDisabled` | `AllowSend=false` in Config | nein |
| `DeleteDisabled` | `AllowDelete=false` in Config | nein |
| `InternalError` | COM-Exception nicht klassifiziert | ja (mit Backoff) |

## Pagination-Konvention

Listen-Tools liefern `value: [...]` + `nextSkip: <int> | null`. Client setzt `skip = nextSkip` für nächste Seite. `nextSkip = null` (oder weggelassen) signalisiert „letzte Seite erreicht".

## Time-Zone-Handling

- Input: `dateTime` als ISO-8601 (lokal ohne Offset) + `timeZone` als IANA-Name (z. B. `"Europe/Berlin"`) ODER `dateTime` als UTC mit `Z`-Suffix
- Intern: Outlook speichert Termine in zwei Formen — lokal (für Anzeige in Outlook) und UTC (für Vergleiche). `AppointmentItem.Start`/`End` ist die lokale Variante; `.StartUTC`/`.EndUTC` ist UTC.
- Mapping: `DateTimeTimeZone { dateTime, timeZone }` → COM-`Start`/`End` mit passender `StartTimeZone`/`EndTimeZone`-Property.
- Output: für Termine aus Outlook → lokale `dateTime` + `timeZone` (aus `StartTimeZone.Name`); UTC-Variante als `originalStartUtc` in `getEvent` (für Vergleiche).
