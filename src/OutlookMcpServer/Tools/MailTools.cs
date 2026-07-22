using System.ComponentModel;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;
using OutlookMcpServer.Domain.Validation;

namespace OutlookMcpServer.Tools;

/// <summary>
/// MCP-Tools fuer Mail-Operationen. Naming: <c>snake_case</c>
/// (z. B. <c>list_mails</c>, <c>get_mail</c>, <c>send_mail</c>).
/// Spezifikation siehe <c>specs/API-DESIGN.md</c>.
/// </summary>
[McpServerToolType]
public sealed class MailTools
{
    private readonly IOutlookService _service;
    private readonly ILogger<MailTools> _logger;

    public MailTools(IOutlookService service, ILogger<MailTools> logger)
    {
        _service = service;
        _logger = logger;
    }

    // ===== Mail: Ordner =====

    [McpServerTool(Name = "list_mail_folders")]
    [Description("Listet Mail-Ordner im aktuellen Outlook-Profil. Eine Ebene: liefert die direkten Kinder des Parent-Ordners (oder Top-Level, wenn parentFolderId=null/fehlt). Fuer volle Baumliste rekursiv selbst aufrufen: fuer jeden zurueckgegebenen Ordner erneut list_mail_folders(parentFolderId=<ordner.id>) und iterieren.")]
    public async Task<PagedResult<MailFolder>> ListMailFolders(
        [Description("Optional: Parent-Folder-ID fuer Unterordner-Auflistung. Null/leer = Top-Level-Ordner des Profils (Posteingang, Gesendete, Entwuerfe, Archiv, etc.).")] string? parentFolderId = null,
        [Description("Hidden-Folder mit aufnehmen (Default false).")] bool includeHidden = false,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("list_mail_folders parentFolderId={Parent} includeHidden={Inc}", parentFolderId, includeHidden);
        return await _service.ListMailFoldersAsync(parentFolderId, includeHidden, cancellationToken);
    }

    [McpServerTool(Name = "get_mail_folder")]
    [Description("Liest einen einzelnen Mail-Ordner per ID oder Well-Known-Name (z. B. 'inbox').")]
    public async Task<MailFolder> GetMailFolder(
        [Description("Folder-ID (EntryID) oder Well-Known-Name (inbox, drafts, sentItems, deletedItems, junkEmail, archive, outbox).")] string folderId,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("get_mail_folder folderId={FolderId}", folderId);
        return await _service.GetMailFolderAsync(folderId, cancellationToken);
    }

    // ===== Mail: Nachrichten =====

    [McpServerTool(Name = "list_mails")]
    [Description("Listet Mails in einem Ordner (Subject, From, To, SentDateTime, ReceivedDateTime, IsRead, HasAttachments, Importance, BodyPreview).")]
    public async Task<PagedResult<MailMessage>> ListMails(
        [Description("Folder-ID (EntryID) oder Well-Known-Name. Erlaubte Well-Known-Namen: inbox, drafts, sentItems, deletedItems, junkEmail, archive, outbox. WICHTIG: durchsucht NUR diesen Ordner, NICHT seine Unterordner. Fuer alle ungelesenen Mails ueber mehrere Ordner: search_mails mit folderId=null, oder mehrere list_mails-Calls pro Ordner (z. B. via list_mail_folders rekursiv).")] string folderId,
        [Description("Max Anzahl zurueckgegebener Mails (1-100, Default 25). Bei mehr Ergebnissen: skip = N * top fuer Seite N (Pagination).")] int top = 25,
        [Description("Skip-Count fuer Pagination (Default 0). Fuer Seite N: skip = N * top.")] int skip = 0,
        [Description(@"Optional: DASL-Filterausdruck fuer Outlooks Items.Restrict(). Der Ausdruck wird unverändert an Outlook weitergereicht — bei ungültiger Syntax fällt der Server auf ungefilterte Iteration zurück.

Syntax: Outlook-DASL (JET-Query-Dialekt). Property-Namen in eckigen Klammern, Operatoren =, <>, >, <, >=, <=, LIKE, AND, OR, NOT.

Häufige Properties:
  [UnRead]             — boolean (true = ungelesen)
  [Subject]            — string (mit LIKE '%suchbegriff%')
  [Body]               — string (mit LIKE '%suchbegriff%')
  [SenderEmailAddress] — string
  [ReceivedTime]       — datetime (US-Format: 'M/d/yyyy h:mm:ss tt', z.B. '7/21/2026 6:00:00 PM')
  [SentOn]             — datetime
  [Importance]         — 0=Low, 1=Normal, 2=High
  [HasAttachments]     — boolean

Beispiele:
  Nur ungelesene Mails:
    [UnRead] = true
  Heute empfangen:
    [ReceivedTime] >= '7/21/2026 12:00:00 AM'
  Wichtige Mails von heute:
    [Importance] = 2 AND [ReceivedTime] >= '7/21/2026 12:00:00 AM'
  Mails mit Anhang aus bestimmtem Absender:
    [HasAttachments] = true AND [SenderEmailAddress] LIKE '%@example.com'

Hinweis: dies ist KEIN OData-Filter. Filter wie 'isRead eq false' oder 'receivedDateTime ge 2026-07-21' werden unveraendert an Outlook weitergegeben, was eine COMException ausloest. Verwende stattdessen die DASL-Syntax oben.")] string? filter = null,
        [Description(@"Optional: Volltext-Suchausdruck. Wird intern als DASL-LIKE-Filter ueber Subject und Body eingesetzt: ([Subject] LIKE '%query%' OR [Body] LIKE '%query%'). Sonderzeichen '%' und '_' werden als Wildcards interpretiert. Beispiel: 'Rechnung' findet Mails mit 'Rechnung' im Subject oder Body.")] string? search = null,
        [Description(@"Format des Mail-Body (siehe get_mail). Default 'markdown'.")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        _logger.LogInformation("list_mails folderId={FolderId} top={Top} skip={Skip} filter={Filter} search={Search} bodyFormat={Bf}", folderId, top, skip, filter, search, bf);
        try
        {
            return await _service.ListMailsAsync(folderId, top, skip, filter, search, bf, cancellationToken);
        }
        catch (OutlookServiceException ex)
        {
            // OutlookServiceException enthaelt bereits Code + Message + InnerException
            _logger.LogWarning(ex, "list_mails OutlookServiceException code={Code}", ex.Code);
            throw;
        }
        catch (OperationCanceledException)
        {
            throw; // Aufrufer hat abgebrochen — nicht transformieren
        }
        catch (Exception ex)
        {
            // Unerwarteter Fehler: vollstaendigen Kontext loggen und als
            // OutlookServiceException weiterreichen, damit der MCP-Client
            // die Originalmeldung bekommt (statt nur 'An error occurred invoking ...').
            _logger.LogError(ex,
                "list_mails unerwarteter Fehler folderId={FolderId} top={Top} skip={Skip} filter={Filter}",
                folderId, top, skip, filter);
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                $"list_mails: unerwarteter Fehler ({ex.GetType().Name}): {ex.Message}. " +
                $"Siehe Server-Log (stderr) fuer vollstaendigen Stack-Trace.",
                ex);
        }
    }

    [McpServerTool(Name = "get_mail")]
    [Description("Liest eine vollstaendige Mail inkl. Body.")]
    public async Task<MailMessage> GetMail(
        [Description("Mail EntryID.")] string id,
        [Description("Body einlesen (Default true) — bei false nur Header/Metadata zur Performance.")] bool includeBody = true,
        [Description(@"Format des Mail-Body. Outlook speichert intern immer HTML — der Server konvertiert on-the-fly. Default 'markdown' (kompakt, gut lesbar fuer LLM). 'text' = Plain Text (maximal kompakt, keine Struktur). 'html' = Outlook-Original-HTML 1:1 (Word/Outlook-Styling erhalten — fuer Tabellen, eingebettete Bilder, komplexe Layouts).")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        _logger.LogInformation("get_mail id={Id} includeBody={Inc} bodyFormat={Bf}", id, includeBody, bf);
        try
        {
            return await _service.GetMailAsync(id, includeBody, bf, cancellationToken);
        }
        catch (OutlookServiceException ex)
        {
            _logger.LogWarning(ex, "get_mail OutlookServiceException code={Code}", ex.Code);
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "get_mail unerwarteter Fehler id={Id} includeBody={Inc}", id, includeBody);
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                $"get_mail: unerwarteter Fehler ({ex.GetType().Name}): {ex.Message}. " +
                $"Siehe Server-Log (stderr) fuer vollstaendigen Stack-Trace.",
                ex);
        }
    }

    [McpServerTool(Name = "get_mails")]
    [Description("Liest mehrere Mails in einem Aufruf (Bulk-Variante von get_mail) — liefert { value: [...], notFoundIds: [...] }. Vermeidet N round-trips wenn Caller bereits eine ID-Liste hat (z. B. aus list_mails).")]
    public async Task<BulkMailResult> GetMails(
        [Description("Liste von Mail EntryIDs (1-50). Wird intern dedupliziert.")] string[] ids,
        [Description("Body einlesen (Default false) — bei true nur wenn wirklich noetig (Performance/Cost).")] bool includeBody = false,
        [Description(@"Format des Mail-Body (siehe get_mail). Default 'markdown'.")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        _logger.LogInformation("get_mails count={Count} includeBody={Inc} bodyFormat={Bf}", ids?.Length ?? 0, includeBody, bf);
        return await _service.GetMailsAsync(ids ?? Array.Empty<string>(), includeBody, bf, cancellationToken);
    }

    [McpServerTool(Name = "get_mail_headers")]
    [Description("Liest Internet-Header einer Mail (From/Sender/To Header, Routing, DKIM etc.).")]
    public async Task<IReadOnlyList<InternetMessageHeader>> GetMailHeaders(
        [Description("Mail EntryID.")] string id,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("get_mail_headers id={Id}", id);
        return await _service.GetMailHeadersAsync(id, cancellationToken);
    }

    [McpServerTool(Name = "list_attachments")]
    [Description("Listet Anlagen einer Mail als Summary (Name, ContentType, Size, IsInline).")]
    public async Task<IReadOnlyList<AttachmentSummary>> ListAttachments(
        [Description("Mail EntryID.")] string mailId,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("list_attachments mailId={MailId}", mailId);
        return await _service.ListAttachmentsAsync(mailId, cancellationToken);
    }

    [McpServerTool(Name = "get_attachment")]
    [Description("Liest eine einzelne Anlage inkl. Base64-codiertem Inhalt (große Anlagen verbrauchen Tokens).")]
    public async Task<AttachmentData> GetAttachment(
        [Description("Mail EntryID.")] string mailId,
        [Description("Attachment-ID oder Index innerhalb der Mail.")] string attachmentId,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("get_attachment mailId={MailId} attachmentId={AttachmentId}", mailId, attachmentId);
        return await _service.GetAttachmentAsync(mailId, attachmentId, cancellationToken);
    }

    [McpServerTool(Name = "search_mails")]
    [Description("Inhaltssuche (Volltext) in Subject + SenderEmailAddress. Liefert Mails, deren Subject oder Sender-Email den Query enthaelt (case-insensitive). Bei folderId=null: REKURSIV ueber ALLE Mail-Ordner des Profils inkl. Unterordner. Keine Filterung auf UnRead/HasAttachments etc. — dafuer list_mails mit DASL-Filter verwenden.")]
    public async Task<PagedResult<MailMessage>> SearchMails(
        [Description("Suchausdruck (Keywords oder Phrase, case-insensitive). Mindestens 1 Zeichen. Wird in Subject und SenderEmailAddress gesucht. Beispiel: 'rechnung' findet Mails mit 'rechnung' im Subject oder Sender; 'pietschmann' findet alle Mails mit @pietschmann-Absender.")] string query,
        [Description(@"Optional: Folder-ID (EntryID) oder Well-Known-Name zur Scope-Einschraenkung (inbox, drafts, sentItems, deletedItems, junkEmail, archive, outbox). Wenn leer/Null: REKURSIVE Suche ueber ALLE Mail-Ordner des Profils (Posteingang, Archiv, Gesendete, etc. plus aller Unterordner).")] string? folderId = null,
        [Description(@"Max Anzahl zurueckgegebener Mails (1-100, Default 25). Bei mehr Treffern: Pagination wird NICHT unterstuetzt (NextSkip ist null) — Suche bricht bei top=100 ab. Fuer vollstaendiges Ergebnis: query enger fassen, oder folderId auf einen Teilbereich einschraenken.")] int top = 25,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("search_mails query={Query} folderId={FolderId} top={Top}", query, folderId, top);
        try
        {
            return await _service.SearchMailsAsync(query, folderId, top, cancellationToken);
        }
        catch (OutlookServiceException ex)
        {
            _logger.LogWarning(ex, "search_mails OutlookServiceException code={Code}", ex.Code);
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "search_mails unerwarteter Fehler query={Query} folderId={FolderId} top={Top}",
                query, folderId, top);
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                $"search_mails: unerwarteter Fehler ({ex.GetType().Name}): {ex.Message}. " +
                $"Siehe Server-Log (stderr) fuer vollstaendigen Stack-Trace.",
                ex);
        }
    }

    [McpServerTool(Name = "list_mails_recursive")]
    [Description(@"Listet Mails REKURSIV ueber mehrere Mail-Ordner UND deren Unterordner. Im Unterschied zu list_mails (nur ein Ordner) und search_mails (Subject/Sender-Textsuche) ist dieses Tool fuer property-basierte Filter ueber den gesamten Posteingangs-Bereich gedacht: filter=[UnRead] = true liefert z. B. alle ungelesenen Mails in Inbox + Archiv + allen anderen gewaehlten Scopes, inkl. Unterordner.

Reichweite: scope (Default = alle Standard-Mailordner): inbox, drafts, sentItems, deletedItems, junkEmail, archive, outbox. Mehrere als Komma-Liste. Pro Folder wird rekursiv in alle Unterordner abgestiegen. Hinweis: es werden alle Outlook-Stores (OST/PST) beruecksichtigt.

Filter: DASL-Ausdruck (gleiche Syntax wie list_mails). Beispiele:
  [UnRead] = true                                                    — nur ungelesene
  [UnRead] = true AND [HasAttachments] = true                       — ungelesene mit Anhang
  [Importance] = 2                                                   — nur High-Importance
  [ReceivedTime] >= '7/1/2026 12:00:00 AM'                          — seit Wochenbeginn (US-Format)

Ergebnis: sortiert nach ReceivedTime DESC (neueste zuerst), dedupliziert per EntryID, Hard-Cap bei top. Es gibt keine Pagination (NextSkip ist null) — bei mehr Treffern die Top einschraenken oder query enger fassen.

Performance: kann bei grossen PSTs langsam sein (vollstaendiger Ordner-Walk). Bei Exchange-Accounts wird der Server-Side-Filter (Restrict) ausgenutzt, wo moeglich.")]
    public async Task<PagedResult<MailMessage>> ListMailsRecursive(
        [Description(@"Optionale Komma-Liste der Well-Known-Mailordner zur Scope-Einschraenkung. Erlaubte Werte: inbox, drafts, sentItems, deletedItems, junkEmail, archive, outbox. Default = alle genannten (alle Mail-Ordner des Profils). Beispiel: 'inbox,archive' durchsucht nur Posteingang + Archiv (jeweils rekursiv).")] string? scope = null,
        [Description(@"Max Anzahl zurueckgegebener Mails (1-100, Default 25). Hard-Cap; keine Pagination. Bei mehr Treffern wird nach ReceivedTime DESC auf top limitiert; aeltere Mails gehen verloren.")] int top = 25,
        [Description(@"Optional: DASL-Filterausdruck, wird an Outlook Items.Restrict() weitergereicht. Bei ungueltiger Syntax Fallback auf ungefilterte Iteration pro Folder (Warnung im Server-Log). Haeufig genutzte Properties: [Unread] (boolean, ACHTUNG: nicht [UnRead]!), [Importance] (0=Low,1=Normal,2=High), [HasAttachments] (boolean), [ReceivedTime] (datetime, US-Format 'M/d/yyyy h:mm:ss tt'). Der Server normalisiert 'UnRead' automatisch zu 'Unread' — beides funktioniert.")] string? filter = null,
        [Description(@"Format des Mail-Body (siehe get_mail). Default 'markdown'.")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var scopeList = string.IsNullOrWhiteSpace(scope)
            ? Array.Empty<string>()
            : scope.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        // Filter-Normalisierung: Tippfehler 'UnRead' (MAPI kennt nur 'Unread') wird
        // serverseitig korrigiert, damit der Restrict nicht stillschweigend fehlschlaegt.
        var normalizedFilter = NormalizeDaslFilter(filter);
        _logger.LogInformation(
            "list_mails_recursive scope={Scope} top={Top} filter={Filter} (raw={RawFilter}) bodyFormat={Bf}",
            string.Join(",", scopeList), top, normalizedFilter, filter, bf);
        try
        {
            return await _service.ListMailsRecursiveAsync(scopeList, top, normalizedFilter, bf, cancellationToken);
        }
        catch (OutlookServiceException ex)
        {
            _logger.LogWarning(ex, "list_mails_recursive OutlookServiceException code={Code}", ex.Code);
            throw;
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "list_mails_recursive unerwarteter Fehler scope={Scope} top={Top} filter={Filter}",
                string.Join(",", scopeList), top, filter);
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                $"list_mails_recursive: unerwarteter Fehler ({ex.GetType().Name}): {ex.Message}. " +
                $"Siehe Server-Log (stderr) fuer vollstaendigen Stack-Trace.",
                ex);
        }
    }

    // ===== Mail: Mutationen =====

    [McpServerTool(Name = "send_mail")]
    [Description("Sendet eine Mail direkt (nicht als Draft). Erfordert AllowSend=true. Subject + IDs werden geloggt, Body-Inhalt NIE.")]
    public async Task<SendMailResult> SendMail(
        [Description("Subject (Betreff).")] string subject,
        [Description("Body Inhalt.")] string body,
        [Description("Body content type: 'text' oder 'html' (Default text).")] string bodyContentType = "text",
        [Description("Komma-getrennte Liste der To-Empfaenger (E-Mail-Adressen).")] string to = "",
        [Description("Optional: Komma-getrennte CC-Empfaenger.")] string? cc = null,
        [Description("Optional: Komma-getrennte BCC-Empfaenger.")] string? bcc = null,
        [Description("Importance: 'low' | 'normal' | 'high' (Default normal).")] string importance = "normal",
        [Description("Kopie in SentItems ablegen (Default true).")] bool saveToSentItems = true,
        [Description("Optional: EntryID einer Mail, auf die geantwortet wird (Subject wird mit 'Re:' ergaenzt).")] string? replyToId = null,
        [Description("Bei replyToId: true = Reply-All, false = nur Reply-To (Default false).")] bool replyAll = false,
        [Description("Optional: EntryID einer weitergeleiteten Mail (Subject wird mit 'Fwd:' ergaenzt).")] string? forwardFromId = null,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("send_mail subject={Subject} toCount={ToCount} ccCount={CcCount} bccCount={BccCount} replyToId={ReplyToId} forwardFromId={ForwardFromId}",
            subject, Count(to), Count(cc), Count(bcc), replyToId, forwardFromId);
        var request = new SendMailRequest
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = string.Equals(bodyContentType, "html", StringComparison.OrdinalIgnoreCase) ? ItemBodyType.Html : ItemBodyType.Text,
            },
            To = SplitRecipients(to),
            Cc = SplitRecipients(cc),
            Bcc = SplitRecipients(bcc),
            Importance = ParseImportance(importance),
            SaveToSentItems = saveToSentItems,
            ReplyToId = replyToId,
            ReplyAll = replyAll,
            ForwardFromId = forwardFromId,
        };
        return await _service.SendMailAsync(request, cancellationToken);
    }

    [McpServerTool(Name = "create_draft")]
    [Description("Legt einen Draft an (nicht gesendet). Erfordert AllowCreate=true.")]
    public async Task<string> CreateDraft(
        [Description("Subject (Betreff).")] string subject,
        [Description("Body Inhalt.")] string body,
        [Description("Body content type: 'text' oder 'html' (Default text).")] string bodyContentType = "text",
        [Description("Komma-getrennte Liste der To-Empfaenger.")] string to = "",
        [Description("Optional: Komma-getrennte CC-Empfaenger.")] string? cc = null,
        [Description("Optional: Komma-getrennte BCC-Empfaenger.")] string? bcc = null,
        [Description("Importance: 'low' | 'normal' | 'high' (Default normal).")] string importance = "normal",
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("create_draft subject={Subject} toCount={ToCount}", subject, Count(to));
        var request = new SendMailRequest
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = string.Equals(bodyContentType, "html", StringComparison.OrdinalIgnoreCase) ? ItemBodyType.Html : ItemBodyType.Text,
            },
            To = SplitRecipients(to),
            Cc = SplitRecipients(cc),
            Bcc = SplitRecipients(bcc),
            Importance = ParseImportance(importance),
            SaveToSentItems = false,
        };
        return await _service.CreateDraftAsync(request, cancellationToken);
    }

    [McpServerTool(Name = "update_mail")]
    [Description("PATCH auf Mail-Felder (IsRead, Categories, Importance) — IsRead=false setzt Unread-Status zurueck.")]
    public async Task UpdateMail(
        [Description("Mail EntryID.")] string id,
        [Description("Als gelesen/ungelesen markieren. null = unveraendert.")] bool? isRead = null,
        [Description("Optional: Komma-getrennte Kategorien (z. B. 'Wichtig,Projekt') — leer = loeschen.")] string? categories = null,
        [Description("Importance: 'low' | 'normal' | 'high' (null = unveraendert).")] string? importance = null,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("update_mail id={Id} isRead={IsRead} categories={Categories} importance={Importance}", id, isRead, categories, importance);
        var update = new MailUpdate
        {
            IsRead = isRead,
            Categories = categories is null ? null : (categories.Length == 0
                ? Array.Empty<string>()
                : categories.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)),
            Importance = importance is null ? null : ParseImportance(importance),
        };
        await _service.UpdateMailAsync(id, update, cancellationToken);
    }

    [McpServerTool(Name = "move_mail")]
    [Description("Verschiebt eine Mail in einen anderen Ordner (EntryID des verschobenen Items wird zurueckgegeben).")]
    public async Task<string> MoveMail(
        [Description("Mail EntryID.")] string id,
        [Description("Ziel-Folder-ID oder Well-Known-Name.")] string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("move_mail id={Id} destFolderId={Dest}", id, destinationFolderId);
        return await _service.MoveMailAsync(id, destinationFolderId, cancellationToken);
    }

    [McpServerTool(Name = "copy_mail")]
    [Description("Kopiert eine Mail in einen anderen Ordner (EntryID der Kopie wird zurueckgegeben).")]
    public async Task<string> CopyMail(
        [Description("Mail EntryID.")] string id,
        [Description("Ziel-Folder-ID oder Well-Known-Name.")] string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("copy_mail id={Id} destFolderId={Dest}", id, destinationFolderId);
        return await _service.CopyMailAsync(id, destinationFolderId, cancellationToken);
    }

    [McpServerTool(Name = "delete_mail")]
    [Description("Loescht eine Mail. permanent=false → in DeletedItems, permanent=true → endgueltig (kann nicht wiederhergestellt werden). Erfordert AllowDelete=true.")]
    public async Task DeleteMail(
        [Description("Mail EntryID.")] string id,
        [Description("true = endgueltig loeschen, false = in DeletedItems verschieben (Default false).")] bool permanent = false,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("delete_mail id={Id} permanent={Permanent}", id, permanent);
        await _service.DeleteMailAsync(id, permanent, cancellationToken);
    }

    // ===== Helpers =====

    private static IReadOnlyList<string> SplitRecipients(string? addresses) =>
        string.IsNullOrWhiteSpace(addresses)
            ? Array.Empty<string>()
            : addresses.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private static int Count(string? addresses) => string.IsNullOrWhiteSpace(addresses) ? 0 : addresses.Split(',', StringSplitOptions.RemoveEmptyEntries).Length;

    /// <summary>
    /// Normalisiert haeufige Tippfehler in DASL-Filtern, die sonst stillschweigend
    /// als Restrict-Fehler enden wuerden. Outlook-DASL-Property-Namen sind
    /// case-insensitive fuer den Match, aber die kanonische Schreibweise
    /// muss exakt stimmen, sonst HResult=0x80020009.
    ///
    /// Bekannte Korrekturen:
    ///   [UnRead]      -> [Unread]   (MAPI kennt nur 'Unread')
    ///   [Read]        -> [Unread]   (Aliases fuer ungerade Tippfehler — bewusst NICHT)
    ///
    /// Wird nur in list_mails_recursive und search_mails angewandt — list_mails
    /// laesst den Filter 1:1 durch (der Caller dort weiss, was er tut).
    /// </summary>
    private static string? NormalizeDaslFilter(string? filter)
    {
        if (string.IsNullOrWhiteSpace(filter)) return filter;
        // Nur [UnRead] -> [Unread], andere Schreibweisen bleiben unveraendert.
        // Wort-Grenzen beachten, damit z. B. [UnReadReceipt] nicht zerschossen wird.
        return System.Text.RegularExpressions.Regex.Replace(
            filter,
            @"\[UnRead\]",
            "[Unread]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    private static Importance ParseImportance(string? value) => (value ?? "normal").Trim().ToLowerInvariant() switch
    {
        "low" => Importance.Low,
        "high" => Importance.High,
        _ => Importance.Normal,
    };
}
