using System.ComponentModel;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Tools;

/// <summary>
/// MCP-Tools fuer Mail-Operationen. Methoden 1:1 zu
/// <c>specs/API-DESIGN.md</c>. Naming: <c>snake_case</c> wie Microsoft Graph
/// (z. B. <c>list_mails</c>, <c>get_mail</c>, <c>send_mail</c>).
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
    [Description("Listet Mail-Ordner im aktuellen Outlook-Profil (1:1 zu Graph listMailFolders).")]
    public async Task<PagedResult<MailFolder>> ListMailFolders(
        [Description("Optional: Parent-Folder-ID fuer Unterordner-Auflistung.")] string? parentFolderId = null,
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
    [Description("Listet Mails in einem Ordner (1:1 zu Graph listMails) — Subject, From, To, SentDateTime, IsRead, HasAttachments, BodyPreview.")]
    public async Task<PagedResult<MailMessage>> ListMails(
        [Description("Folder-ID oder Well-Known-Name.")] string folderId,
        [Description("Max Anzahl (1-100, Default 25).")] int top = 25,
        [Description("Skip-Count fuer Pagination (Default 0).")] int skip = 0,
        [Description("Optional: OData-Filter-Ausdruck (nur eingeschraenkter Subset, siehe API-DESIGN).")] string? filter = null,
        [Description("Optional: Volltext-Suchausdruck ueber Subject/Body/Sender.")] string? search = null,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("list_mails folderId={FolderId} top={Top} skip={Skip} filter={Filter} search={Search}", folderId, top, skip, filter, search);
        try
        {
            return await _service.ListMailsAsync(folderId, top, skip, filter, search, cancellationToken);
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
    [Description("Liest eine vollstaendige Mail inkl. Body (1:1 zu Graph getMessage).")]
    public async Task<MailMessage> GetMail(
        [Description("Mail EntryID.")] string id,
        [Description("Body einlesen (Default true) — bei false nur Header/Metadata zur Performance.")] bool includeBody = true,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("get_mail id={Id} includeBody={Inc}", id, includeBody);
        try
        {
            return await _service.GetMailAsync(id, includeBody, cancellationToken);
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
    [Description("Volltext-Suche ueber Subject/Body/Sender ueber optional definierbares FolderScope.")]
    public async Task<PagedResult<MailMessage>> SearchMails(
        [Description("Suchausdruck (Keywords oder Phrase).")] string query,
        [Description("Optional: Folder-ID oder Well-Known-Name fuer Scope-Einschraenkung.")] string? folderId = null,
        [Description("Max Anzahl (1-100, Default 25).")] int top = 25,
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

    private static Importance ParseImportance(string? value) => (value ?? "normal").Trim().ToLowerInvariant() switch
    {
        "low" => Importance.Low,
        "high" => Importance.High,
        _ => Importance.Normal,
    };
}
