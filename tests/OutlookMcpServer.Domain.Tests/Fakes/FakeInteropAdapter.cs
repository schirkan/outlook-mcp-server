using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Domain.Tests.Fakes;

/// <summary>
/// In-Memory-Implementierung von <see cref="IInteropOutlookAdapter"/> fuer
/// Unit-Tests von <c>OutlookService</c>-Validation/Policy-Pfaden.
/// Wirft standardmaessig nicht; rufer-spezifische Fehler koennen via
/// <see cref="OnSendMailAsync"/> etc. injiziert werden.
/// </summary>
public sealed class FakeInteropAdapter : IInteropOutlookAdapter
{
    public List<string> Calls { get; } = new();

    public Func<SendMailRequest, SendMailResult>? OnSendMailAsync { get; set; }
    public Func<string, Task>? OnDeleteMailAsync { get; set; }

    public Task<PagedResult<MailFolder>> ListMailFoldersAsync(string? parentFolderId = null, bool includeHidden = false, CancellationToken cancellationToken = default)
        => Task.FromResult(new PagedResult<MailFolder> { Value = Array.Empty<MailFolder>() });

    public Task<MailFolder> GetMailFolderAsync(string folderId, CancellationToken cancellationToken = default)
        => throw new NotSupportedException("Adapter-Stub: not used in this test");

    public Task<PagedResult<MailMessage>> ListMailsAsync(string folderId, int top = 25, int skip = 0, string? filter = null, string? search = null, BodyFormat bodyFormat = BodyFormat.Markdown, CancellationToken cancellationToken = default)
    {
        Calls.Add(nameof(ListMailsAsync));
        return Task.FromResult(new PagedResult<MailMessage> { Value = Array.Empty<MailMessage>() });
    }

    public Task<MailMessage> GetMailAsync(string id, bool includeBody = true, BodyFormat bodyFormat = BodyFormat.Markdown, CancellationToken cancellationToken = default)
        => throw new NotSupportedException();

    public Task<BulkMailResult> GetMailsAsync(
        IReadOnlyList<string> ids,
        bool includeBody = false,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetMailsAsync)}:count={ids.Count},includeBody={includeBody},bodyFormat={bodyFormat}");
        // Default: nichts gefunden, alle in notFoundIds. Caller-Tests koennen OnGetMails ueberschreiben
        // (Property existiert auf FakeOutlookService, hier nicht noetig — Service-Validation-Tests zaehlen nur Calls).
        return Task.FromResult(new BulkMailResult
        {
            Value = Array.Empty<MailMessage>(),
            NotFoundIds = ids.ToList(),
        });
    }

    public Task<IReadOnlyList<InternetMessageHeader>> GetMailHeadersAsync(string id, CancellationToken cancellationToken = default)
        => throw new NotSupportedException();

    public Task<IReadOnlyList<AttachmentSummary>> ListAttachmentsAsync(string mailId, CancellationToken cancellationToken = default)
        => throw new NotSupportedException();

    public Task<AttachmentData> GetAttachmentAsync(string mailId, string attachmentId, CancellationToken cancellationToken = default)
        => throw new NotSupportedException();

    public Task<PagedResult<MailMessage>> SearchMailsAsync(string query, string? folderId = null, int top = 25, CancellationToken cancellationToken = default)
        => Task.FromResult(new PagedResult<MailMessage> { Value = Array.Empty<MailMessage>() });

    /// <summary>
    /// Test-Hook: mails, die von <see cref="ListMailsRecursiveAsync"/> zurueckgegeben werden.
    /// </summary>
    public Func<ListMailsRecursiveArgs, IReadOnlyList<MailMessage>>? OnListMailsRecursiveAsync { get; set; }

    public Task<PagedResult<MailMessage>> ListMailsRecursiveAsync(
        IReadOnlyList<string> scope,
        int top,
        string? filter,
        BodyFormat bodyFormat,
        CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(ListMailsRecursiveAsync)}:scope={string.Join(",", scope)},top={top},filter={filter ?? "(null)"},bodyFormat={bodyFormat}");
        IReadOnlyList<MailMessage> items;
        if (OnListMailsRecursiveAsync is null)
        {
            items = Array.Empty<MailMessage>();
        }
        else
        {
            items = OnListMailsRecursiveAsync(new ListMailsRecursiveArgs(
                new List<string>(scope), top, filter, bodyFormat));
        }
        var slice = new List<MailMessage>(Math.Min(top, items.Count));
        for (int i = 0; i < items.Count && i < top; i++) slice.Add(items[i]);
        return Task.FromResult(new PagedResult<MailMessage> { Value = slice, NextSkip = null });
    }

    public Task<SendMailResult> SendMailAsync(SendMailRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(SendMailAsync)}:subject={request.Subject}");
        return Task.FromResult(OnSendMailAsync?.Invoke(request) ?? new SendMailResult { Sent = true, Id = "SENT-1" });
    }

    public Task<string> CreateDraftAsync(SendMailRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add(nameof(CreateDraftAsync));
        return Task.FromResult("DRAFT-1");
    }

    public Task UpdateMailAsync(string id, MailUpdate update, CancellationToken cancellationToken = default)
        => Task.CompletedTask;

    public Task<string> MoveMailAsync(string id, string destinationFolderId, CancellationToken cancellationToken = default)
        => Task.FromResult($"MOVED-{id}");

    public Task<string> CopyMailAsync(string id, string destinationFolderId, CancellationToken cancellationToken = default)
        => Task.FromResult($"COPIED-{id}");

    public async Task DeleteMailAsync(string id, bool permanent = false, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(DeleteMailAsync)}:id={id},permanent={permanent}");
        if (OnDeleteMailAsync is not null) await OnDeleteMailAsync(id);
    }

    public Task<IReadOnlyList<Calendar>> ListCalendarsAsync(CancellationToken cancellationToken = default)
        => Task.FromResult<IReadOnlyList<Calendar>>(Array.Empty<Calendar>());

    public Task<Calendar> GetCalendarAsync(string id, CancellationToken cancellationToken = default)
        => throw new NotSupportedException();

    public Task<PagedResult<CalendarEvent>> ListEventsAsync(string? calendarId, DateTimeTimeZone start, DateTimeTimeZone end, int top = 50, int skip = 0, string? filter = null, BodyFormat bodyFormat = BodyFormat.Markdown, CancellationToken cancellationToken = default)
        => Task.FromResult(new PagedResult<CalendarEvent> { Value = Array.Empty<CalendarEvent>() });

    public Task<CalendarEvent> GetEventAsync(string id, BodyFormat bodyFormat = BodyFormat.Markdown, CancellationToken cancellationToken = default)
        => throw new NotSupportedException();

    public Task<string> CreateEventAsync(CreateEventRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add(nameof(CreateEventAsync));
        return Task.FromResult("EVT-1");
    }

    public Task UpdateEventAsync(string id, EventUpdate update, CancellationToken cancellationToken = default)
        => Task.CompletedTask;

    public Task DeleteEventAsync(string id, bool sendCancellation = true, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(DeleteEventAsync)}:id={id},sendCancellation={sendCancellation}");
        return Task.CompletedTask;
    }

    public Task RespondToEventAsync(string id, RespondToEventRequest request, CancellationToken cancellationToken = default)
        => Task.CompletedTask;

    public Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(FindMeetingTimesRequest request, CancellationToken cancellationToken = default)
        => Task.FromResult<IReadOnlyList<MeetingTimeCandidate>>(Array.Empty<MeetingTimeCandidate>());

    // ===== Active-Inspector / Selection =====

    public Func<BodyFormat, CancellationToken, ActiveItem?>? OnGetActiveItem { get; set; }
    public Func<SelectionScope, int, BodyFormat, CancellationToken, IReadOnlyList<ActiveItem>>? OnGetSelectedItems { get; set; }

    public Task<ActiveItem?> GetActiveItemAsync(
        BodyFormat bodyFormat,
        CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetActiveItemAsync)}:bodyFormat={bodyFormat}");
        if (OnGetActiveItem is not null) return Task.FromResult(OnGetActiveItem(bodyFormat, cancellationToken));
        return Task.FromResult<ActiveItem?>(null);
    }

    public Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top,
        BodyFormat bodyFormat,
        CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetSelectedItemsAsync)}:scope={scope},top={top},bodyFormat={bodyFormat}");
        if (OnGetSelectedItems is not null) return Task.FromResult(OnGetSelectedItems(scope, top, bodyFormat, cancellationToken));
        return Task.FromResult<IReadOnlyList<ActiveItem>>(Array.Empty<ActiveItem>());
    }
}
