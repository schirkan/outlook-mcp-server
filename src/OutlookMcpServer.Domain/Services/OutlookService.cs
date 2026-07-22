using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Configuration;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;
using OutlookMcpServer.Domain.Validation;

namespace OutlookMcpServer.Domain.Services;

/// <summary>
/// Standardimplementierung von <see cref="IOutlookService"/>. Validiert alle
/// Inputs, prueft AllowSend/AllowDelete/AllowCreate-Flags aus
/// <see cref="OutlookMcpServerOptions"/>, delegiert an
/// <see cref="IInteropOutlookAdapter"/> und mappt Exceptions auf
/// <see cref="OutlookServiceException"/>.
/// </summary>
public sealed class OutlookService : IOutlookService
{
    private readonly IInteropOutlookAdapter _adapter;
    private readonly OutlookMcpServerOptions _options;
    private readonly ILogger<OutlookService> _logger;

    public OutlookService(
        IInteropOutlookAdapter adapter,
        IOptions<OutlookMcpServerOptions> options,
        ILogger<OutlookService> logger)
    {
        _adapter = adapter;
        _options = options.Value;
        _logger = logger;
    }

    // ===== Mail: Ordner =====

    public Task<PagedResult<MailFolder>> ListMailFoldersAsync(
        string? parentFolderId = null,
        bool includeHidden = false,
        CancellationToken cancellationToken = default)
    {
        return _adapter.ListMailFoldersAsync(parentFolderId, includeHidden, cancellationToken);
    }

    public Task<MailFolder> GetMailFolderAsync(
        string folderId,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateFolderId(folderId);
        return _adapter.GetMailFolderAsync(folderId, cancellationToken);
    }

    // ===== Mail: Nachrichten =====

    public Task<PagedResult<MailMessage>> ListMailsAsync(
        string folderId,
        int top = 25,
        int skip = 0,
        string? filter = null,
        string? search = null,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateFolderId(folderId, nameof(folderId));
        ValidationHelpers.ValidateRange(top, 1, 100, nameof(top));
        ValidationHelpers.ValidateRange(skip, 0, int.MaxValue, nameof(skip));
        return _adapter.ListMailsAsync(folderId, top, skip, filter, search, bodyFormat, cancellationToken);
    }

    public Task<MailMessage> GetMailAsync(
        string id,
        bool includeBody = true,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.GetMailAsync(id, includeBody, bodyFormat, cancellationToken);
    }

    public async Task<BulkMailResult> GetMailsAsync(
        IReadOnlyList<string> ids,
        bool includeBody = false,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        if (ids is null)
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                "ids: Liste ist null");
        }
        // Hard-Cap 50: Bulk-Variante vermeidet N round-trips, aber ein Caller
        // mit 1000 IDs soll weiterhin paginieren (list_mails + iter).
        ValidationHelpers.ValidateRange(ids.Count, 1, 50, "ids.Count");
        foreach (var id in ids)
        {
            ValidationHelpers.ValidateStringNotEmpty(id, "ids[]");
        }
        // Deduplizieren: gleiche EntryID mehrfach -> nur einmal fetchen.
        var distinct = ids.Distinct(StringComparer.Ordinal).ToList();
        return await _adapter.GetMailsAsync(distinct, includeBody, bodyFormat, cancellationToken);
    }

    public Task<IReadOnlyList<InternetMessageHeader>> GetMailHeadersAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.GetMailHeadersAsync(id, cancellationToken);
    }

    public Task<IReadOnlyList<AttachmentSummary>> ListAttachmentsAsync(
        string mailId,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(mailId, nameof(mailId));
        return _adapter.ListAttachmentsAsync(mailId, cancellationToken);
    }

    public Task<AttachmentData> GetAttachmentAsync(
        string mailId,
        string attachmentId,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(mailId, nameof(mailId));
        ValidationHelpers.ValidateStringNotEmpty(attachmentId, nameof(attachmentId));
        return _adapter.GetAttachmentAsync(mailId, attachmentId, cancellationToken);
    }

    public Task<PagedResult<MailMessage>> SearchMailsAsync(
        string query,
        string? folderId = null,
        int top = 25,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(query, nameof(query));
        ValidationHelpers.ValidateRange(top, 1, 100, nameof(top));
        if (folderId is not null) ValidationHelpers.ValidateFolderId(folderId, nameof(folderId));
        return _adapter.SearchMailsAsync(query, folderId, top, cancellationToken);
    }

    public Task<PagedResult<MailMessage>> ListMailsRecursiveAsync(
        IReadOnlyList<string> scope,
        int top = 25,
        string? filter = null,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateRange(top, 1, 100, nameof(top));
        // scope-Validierung: null/leer -> Default-Mailordner = OK.
        // Sonst jeder Eintrag muss ein gueltiger Well-Known-Name sein.
        if (scope is not null)
        {
            foreach (var s in scope)
            {
                if (string.IsNullOrWhiteSpace(s))
                {
                    throw new OutlookServiceException(
                        ErrorCode.InvalidInput,
                        "scope: enthaelt leeren oder null Eintrag");
                }
                if (!WellKnownFolder.IsKnownMailFolder(s))
                {
                    throw new OutlookServiceException(
                        ErrorCode.InvalidInput,
                        $"scope: '{s}' ist kein Well-Known-Mailordner. Erlaubt: " +
                        string.Join(", ", WellKnownFolder.MailFolderNames));
                }
            }
        }
        return _adapter.ListMailsRecursiveAsync(scope ?? Array.Empty<string>(), top, filter, bodyFormat, cancellationToken);
    }

    // ===== Mail: Mutationen =====

    public Task<SendMailResult> SendMailAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default)
    {
        EnsureSendAllowed(request);
        ValidateSendMailRequest(request);
        return _adapter.SendMailAsync(request, cancellationToken);
    }

    public Task<string> CreateDraftAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default)
    {
        EnsureCreateAllowed();
        ValidateSendMailRequest(request);
        return _adapter.CreateDraftAsync(request, cancellationToken);
    }

    public Task UpdateMailAsync(
        string id,
        MailUpdate update,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.UpdateMailAsync(id, update, cancellationToken);
    }

    public Task<string> MoveMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        ValidationHelpers.ValidateFolderId(destinationFolderId, nameof(destinationFolderId));
        return _adapter.MoveMailAsync(id, destinationFolderId, cancellationToken);
    }

    public Task<string> CopyMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        ValidationHelpers.ValidateFolderId(destinationFolderId, nameof(destinationFolderId));
        return _adapter.CopyMailAsync(id, destinationFolderId, cancellationToken);
    }

    public Task DeleteMailAsync(
        string id,
        bool permanent = false,
        CancellationToken cancellationToken = default)
    {
        if (!_options.Outlook.AllowDelete)
        {
            throw new OutlookServiceException(
                ErrorCode.DeleteDisabled,
                "Loeschen ist deaktiviert (OutlookMcpServer:Outlook:AllowDelete=false)");
        }
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.DeleteMailAsync(id, permanent, cancellationToken);
    }

    // ===== Calendar: Kalender =====

    public Task<IReadOnlyList<Calendar>> ListCalendarsAsync(
        CancellationToken cancellationToken = default)
    {
        return _adapter.ListCalendarsAsync(cancellationToken);
    }

    public Task<Calendar> GetCalendarAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.GetCalendarAsync(id, cancellationToken);
    }

    // ===== Calendar: Termine =====

    public Task<PagedResult<CalendarEvent>> ListEventsAsync(
        string? calendarId,
        DateTimeTimeZone start,
        DateTimeTimeZone end,
        int top = 50,
        int skip = 0,
        string? filter = null,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateDateTimeTimeZone(start, nameof(start));
        ValidationHelpers.ValidateDateTimeTimeZone(end, nameof(end));
        ValidationHelpers.ValidateRange(top, 1, 250, nameof(top));
        ValidationHelpers.ValidateRange(skip, 0, int.MaxValue, nameof(skip));
        if (calendarId is not null) ValidationHelpers.ValidateStringNotEmpty(calendarId, nameof(calendarId));
        return _adapter.ListEventsAsync(calendarId, start, end, top, skip, filter, bodyFormat, cancellationToken);
    }

    public Task<CalendarEvent> GetEventAsync(
        string id,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.GetEventAsync(id, bodyFormat, cancellationToken);
    }

    // ===== Calendar: Mutationen =====

    public Task<string> CreateEventAsync(
        CreateEventRequest request,
        CancellationToken cancellationToken = default)
    {
        EnsureCreateAllowed();
        ValidateCreateEventRequest(request);
        return _adapter.CreateEventAsync(request, cancellationToken);
    }

    public Task UpdateEventAsync(
        string id,
        EventUpdate update,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.UpdateEventAsync(id, update, cancellationToken);
    }

    public Task DeleteEventAsync(
        string id,
        bool sendCancellation = true,
        CancellationToken cancellationToken = default)
    {
        if (!_options.Outlook.AllowDelete)
        {
            throw new OutlookServiceException(
                ErrorCode.DeleteDisabled,
                "Loeschen ist deaktiviert (OutlookMcpServer:Outlook:AllowDelete=false)");
        }
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        return _adapter.DeleteEventAsync(id, sendCancellation, cancellationToken);
    }

    public Task RespondToEventAsync(
        string id,
        RespondToEventRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateStringNotEmpty(id, nameof(id));
        if (request.Response == Response.None || request.Response == Response.NotResponded)
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"response: '{request.Response}' ist nicht zulaessig fuer respondToEvent");
        }
        return _adapter.RespondToEventAsync(id, request, cancellationToken);
    }

    public Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(
        FindMeetingTimesRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateRange(request.DurationMinutes, 1, 24 * 60, nameof(request.DurationMinutes));
        ValidationHelpers.ValidateDateTimeTimeZone(request.TimeWindowStart, nameof(request.TimeWindowStart));
        ValidationHelpers.ValidateDateTimeTimeZone(request.TimeWindowEnd, nameof(request.TimeWindowEnd));
        ValidationHelpers.ValidateRange(request.MaxCandidates, 1, 100, nameof(request.MaxCandidates));
        return _adapter.FindMeetingTimesAsync(request, cancellationToken);
    }

    // ===== Active-Inspector / Selection (COM-only) =====

    public Task<ActiveItem?> GetActiveItemAsync(
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        // Passthrough: null ist valides Resultat (kein Inspector oder v1-out-of-scope Typ)
        return _adapter.GetActiveItemAsync(bodyFormat, cancellationToken);
    }

    public async Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top = 50,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        ValidationHelpers.ValidateRange(top, 1, 250, nameof(top));
        return await _adapter.GetSelectedItemsAsync(scope, top, bodyFormat, cancellationToken);
    }

    // ===== Private Helpers =====

    private void EnsureSendAllowed(SendMailRequest request)
    {
        if (!_options.Outlook.AllowSend)
        {
            throw new OutlookServiceException(
                ErrorCode.SendDisabled,
                "Senden ist deaktiviert (OutlookMcpServer:Outlook:AllowSend=false)");
        }
        // replyToId / forwardFromId ohne Mail-Versand waere Quatsch, aber erlaubt.
        // Wenn kein replyToId/forwardFromId: muss 'to' mindestens eine Adresse haben.
        if (string.IsNullOrEmpty(request.ReplyToId) && string.IsNullOrEmpty(request.ForwardFromId))
        {
            if (request.To.Count == 0)
            {
                throw new OutlookServiceException(
                    ErrorCode.InvalidInput,
                    "to: mindestens eine Empfaenger-Adresse erforderlich (ausser bei replyToId/forwardFromId)");
            }
        }
    }

    private void EnsureCreateAllowed()
    {
        if (!_options.Outlook.AllowCreate)
        {
            throw new OutlookServiceException(
                ErrorCode.SendDisabled,
                "Erstellen ist deaktiviert (OutlookMcpServer:Outlook:AllowCreate=false)");
        }
    }

    private void ValidateSendMailRequest(SendMailRequest request)
    {
        ValidationHelpers.ValidateEmails(request.To, "to");
        ValidationHelpers.ValidateEmails(request.Cc, "cc");
        ValidationHelpers.ValidateEmails(request.Bcc, "bcc");
        ValidationHelpers.ValidateStringNotEmpty(request.Subject, nameof(request.Subject));
        if (request.Body is null || string.IsNullOrEmpty(request.Body.Content))
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                "body.content: Body-Inhalt ist erforderlich");
        }
        // Attachment-Groessen pruefen
        foreach (var att in request.Attachments)
        {
            // Base64-Laenge * 3/4 = ungefaehre Byte-Groesse
            var approxBytes = att.ContentBase64.Length * 3L / 4L;
            ValidationHelpers.ValidateAttachmentSize(approxBytes, _options.Outlook.MaxAttachmentBytes);
        }
    }

    private void ValidateCreateEventRequest(CreateEventRequest request)
    {
        ValidationHelpers.ValidateStringNotEmpty(request.Subject, nameof(request.Subject));
        ValidationHelpers.ValidateDateTimeTimeZone(request.Start, nameof(request.Start));
        ValidationHelpers.ValidateDateTimeTimeZone(request.End, nameof(request.End));
        ValidationHelpers.ValidateEmails(request.Attendees.Select(a => a.Email), "attendees");
        if (request.ReminderMinutesBeforeStart is { } reminder && (reminder < 0 || reminder > 40320))
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"reminderMinutesBeforeStart: Wert {reminder} ausserhalb [0, 40320] (max 4 Wochen)");
        }
    }
}