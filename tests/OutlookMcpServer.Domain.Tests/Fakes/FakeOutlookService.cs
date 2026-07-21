using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;
using OutlookMcpServer.Domain.Validation;

namespace OutlookMcpServer.Domain.Tests.Fakes;

/// <summary>
/// In-Memory-Implementierung von <see cref="IOutlookService"/> fuer
/// Unit-Tests der MCP-Tools (MailTools + CalendarTools).
/// Zaehlt alle Calls, deterministische Builder-Daten, optionale Callbacks
/// fuer Failure-Pfade.
/// </summary>
public sealed class FakeOutlookService : IOutlookService
{
    private readonly Dictionary<string, MailMessage> _mails = new();
    private readonly Dictionary<string, CalendarEvent> _events = new();
    private readonly Dictionary<string, MailFolder> _folders = new();
    private readonly Dictionary<string, Calendar> _calendars = new();

    public List<string> Calls { get; } = new();
    public Func<string, MailMessage?>? OnGetMail { get; set; }
    public Func<SendMailRequest, SendMailResult>? OnSendMail { get; set; }
    public Exception? InjectException { get; set; }

    public void SeedMail(MailMessage mail) => _mails[mail.Id] = mail;
    public void SeedEvent(CalendarEvent evt) => _events[evt.Id] = evt;
    public void SeedFolder(MailFolder folder) => _folders[folder.Id] = folder;
    public void SeedCalendar(Calendar calendar) => _calendars[calendar.Id] = calendar;

    public Task<PagedResult<MailFolder>> ListMailFoldersAsync(string? parentFolderId = null, bool includeHidden = false, CancellationToken cancellationToken = default)
    {
        Calls.Add(nameof(ListMailFoldersAsync));
        var all = _folders.Values
            .Where(f => includeHidden || !IsHiddenName(f.DisplayName))
            .Where(f => parentFolderId is null || f.ParentFolderId == parentFolderId)
            .ToList();
        return Task.FromResult(new PagedResult<MailFolder> { Value = all });
    }

    public Task<MailFolder> GetMailFolderAsync(string folderId, CancellationToken cancellationToken = default)
    {
        Calls.Add(nameof(GetMailFolderAsync));
        if (_folders.TryGetValue(folderId, out var f)) return Task.FromResult(f);
        throw new KeyNotFoundException(folderId);
    }

    public Task<PagedResult<MailMessage>> ListMailsAsync(string folderId, int top = 25, int skip = 0, string? filter = null, string? search = null, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(ListMailsAsync)}:folder={folderId},top={top},skip={skip},filter={filter},search={search}");
        var query = _mails.Values.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(search))
        {
            query = query.Where(m => (m.Subject ?? string.Empty).Contains(search, StringComparison.OrdinalIgnoreCase));
        }
        var list = query.Skip(skip).Take(top).ToList();
        return Task.FromResult(new PagedResult<MailMessage>
        {
            Value = list,
            NextSkip = query.Count() > skip + top ? skip + top : null,
        });
    }

    public Task<MailMessage> GetMailAsync(string id, bool includeBody = true, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetMailAsync)}:id={id},includeBody={includeBody}");
        ThrowIfInjected();
        if (OnGetMail is not null)
        {
            var result = OnGetMail(id);
            if (result is not null) return Task.FromResult(result);
        }
        if (_mails.TryGetValue(id, out var m)) return Task.FromResult(m);
        throw new KeyNotFoundException(id);
    }

    public Func<IReadOnlyList<string>, bool, IReadOnlyList<MailMessage>>? OnGetMails { get; set; }
    private readonly Dictionary<string, MailMessage> _seededBulkMails = new();

    public void SeedBulkMails(IEnumerable<MailMessage> mails)
    {
        foreach (var m in mails) _seededBulkMails[m.Id] = m;
    }

    public Task<BulkMailResult> GetMailsAsync(
        IReadOnlyList<string> ids,
        bool includeBody = false,
        CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetMailsAsync)}:count={ids.Count},includeBody={includeBody}");
        ThrowIfInjected();
        if (OnGetMails is not null)
        {
            return Task.FromResult(BuildBulkResultFromOverride(OnGetMails(ids, includeBody), ids));
        }
        return Task.FromResult(BuildBulkResultFromSeed(ids));
    }

    /// <summary>
    /// Baut ein <see cref="BulkMailResult"/> wenn der Caller einen
    /// <see cref="OnGetMails"/> Override registriert hat. Die Override-Methode
    /// liefert nur die gefundenen Mails; alle nicht abgedeckten IDs wandern
    /// in <c>notFoundIds</c>.
    /// </summary>
    private static BulkMailResult BuildBulkResultFromOverride(
        IReadOnlyList<MailMessage> overrideResult,
        IReadOnlyList<string> requestedIds)
    {
        var foundIds = overrideResult.Select(m => m.Id).ToHashSet(StringComparer.Ordinal);
        var missing = requestedIds.Where(id => !foundIds.Contains(id)).ToList();
        return new BulkMailResult { Value = overrideResult, NotFoundIds = missing };
    }

    /// <summary>
    /// Default-Verhalten: iteriert ueber die (vom Service bereits deduplizierten)
    /// IDs und schaut in <c>_seededBulkMails</c> nach. Gefundene landen in
    /// <c>value</c>, nicht-gefundene in <c>notFoundIds</c>.
    /// </summary>
    private BulkMailResult BuildBulkResultFromSeed(IReadOnlyList<string> ids)
    {
        var found = new List<MailMessage>(ids.Count);
        var missing = new List<string>();
        foreach (var id in ids)
        {
            if (_seededBulkMails.TryGetValue(id, out var mail))
            {
                found.Add(mail);
            }
            else
            {
                missing.Add(id);
            }
        }
        return new BulkMailResult { Value = found, NotFoundIds = missing };
    }

    public Task<IReadOnlyList<InternetMessageHeader>> GetMailHeadersAsync(string id, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetMailHeadersAsync)}:id={id}");
        return Task.FromResult<IReadOnlyList<InternetMessageHeader>>(Array.Empty<InternetMessageHeader>());
    }

    public Task<IReadOnlyList<AttachmentSummary>> ListAttachmentsAsync(string mailId, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(ListAttachmentsAsync)}:mailId={mailId}");
        return Task.FromResult<IReadOnlyList<AttachmentSummary>>(Array.Empty<AttachmentSummary>());
    }

    public Task<AttachmentData> GetAttachmentAsync(string mailId, string attachmentId, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetAttachmentAsync)}:mailId={mailId},attId={attachmentId}");
        return Task.FromResult(new AttachmentData { Id = attachmentId, Name = "stub.txt", ContentType = "text/plain", Size = 0, ContentBase64 = string.Empty });
    }

    public Task<PagedResult<MailMessage>> SearchMailsAsync(string query, string? folderId = null, int top = 25, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(SearchMailsAsync)}:query={query},folder={folderId}");
        var list = _mails.Values
            .Where(m => (m.Subject ?? string.Empty).Contains(query, StringComparison.OrdinalIgnoreCase))
            .Take(top)
            .ToList();
        return Task.FromResult(new PagedResult<MailMessage> { Value = list });
    }

    public Task<PagedResult<MailMessage>> ListMailsRecursiveAsync(
        IReadOnlyList<string> scope,
        int top = 25,
        string? filter = null,
        CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(ListMailsRecursiveAsync)}:scope={string.Join(",", scope)},top={top},filter={filter ?? "(null)"}");
        // Default-Fake: leere Liste (Tests, die Inhalte brauchen, nehmen besser FakeInteropAdapter direkt).
        return Task.FromResult(new PagedResult<MailMessage> { Value = Array.Empty<MailMessage>(), NextSkip = null });
    }

    public Task<SendMailResult> SendMailAsync(SendMailRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(SendMailAsync)}:subject={request.Subject},to={request.To.Count}");
        ThrowIfInjected();
        if (OnSendMail is not null) return Task.FromResult(OnSendMail(request));
        return Task.FromResult(new SendMailResult { Sent = true, Id = "SENT-1" });
    }

    public Task<string> CreateDraftAsync(SendMailRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(CreateDraftAsync)}:subject={request.Subject}");
        return Task.FromResult("DRAFT-1");
    }

    public Task UpdateMailAsync(string id, MailUpdate update, CancellationToken cancellationToken = default)
    {
        var categories = update.Categories is null ? "<null>" : string.Join("|", update.Categories);
        Calls.Add($"{nameof(UpdateMailAsync)}:id={id},isRead={update.IsRead},categories=[{categories}],importance={update.Importance}");
        return Task.CompletedTask;
    }

    public Task<string> MoveMailAsync(string id, string destinationFolderId, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(MoveMailAsync)}:id={id},dest={destinationFolderId}");
        return Task.FromResult($"MOVED-{id}");
    }

    public Task<string> CopyMailAsync(string id, string destinationFolderId, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(CopyMailAsync)}:id={id},dest={destinationFolderId}");
        return Task.FromResult($"COPIED-{id}");
    }

    public Task DeleteMailAsync(string id, bool permanent = false, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(DeleteMailAsync)}:id={id},permanent={permanent}");
        return Task.CompletedTask;
    }

    public Task<IReadOnlyList<Calendar>> ListCalendarsAsync(CancellationToken cancellationToken = default)
    {
        Calls.Add(nameof(ListCalendarsAsync));
        return Task.FromResult<IReadOnlyList<Calendar>>(_calendars.Values.ToList());
    }

    public Task<Calendar> GetCalendarAsync(string id, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetCalendarAsync)}:id={id}");
        if (_calendars.TryGetValue(id, out var c)) return Task.FromResult(c);
        throw new KeyNotFoundException(id);
    }

    public Task<PagedResult<CalendarEvent>> ListEventsAsync(string? calendarId, DateTimeTimeZone start, DateTimeTimeZone end, int top = 50, int skip = 0, string? filter = null, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(ListEventsAsync)}:cal={calendarId},window={start.DateTime}/{end.DateTime},top={top},skip={skip},filter={filter}");
        var list = _events.Values.Skip(skip).Take(top).ToList();
        return Task.FromResult(new PagedResult<CalendarEvent> { Value = list });
    }

    public Task<CalendarEvent> GetEventAsync(string id, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetEventAsync)}:id={id}");
        if (_events.TryGetValue(id, out var e)) return Task.FromResult(e);
        throw new KeyNotFoundException(id);
    }

    public Task<string> CreateEventAsync(CreateEventRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(CreateEventAsync)}:subject={request.Subject},attendeeCount={request.Attendees.Count}");
        return Task.FromResult("EVT-1");
    }

    public Task UpdateEventAsync(string id, EventUpdate update, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(UpdateEventAsync)}:id={id},subject={update.Subject}");
        return Task.CompletedTask;
    }

    public Task DeleteEventAsync(string id, bool sendCancellation = true, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(DeleteEventAsync)}:id={id},sendCancellation={sendCancellation}");
        return Task.CompletedTask;
    }

    public Task RespondToEventAsync(string id, RespondToEventRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(RespondToEventAsync)}:id={id},response={request.Response}");
        return Task.CompletedTask;
    }

    public Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(FindMeetingTimesRequest request, CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(FindMeetingTimesAsync)}:duration={request.DurationMinutes}");
        return Task.FromResult<IReadOnlyList<MeetingTimeCandidate>>(Array.Empty<MeetingTimeCandidate>());
    }

    // ===== Active-Inspector / Selection (Fakes fuer Tests) =====

    public Func<CancellationToken, ActiveItem?>? OnGetActiveItem { get; set; }
    public Func<SelectionScope, int, CancellationToken, IReadOnlyList<ActiveItem>>? OnGetSelectedItems { get; set; }
    private readonly List<ActiveItem> _seededSelectedItems = new();
    private ActiveItem? _seededActiveItem;

    public void SeedActiveItem(ActiveItem item) => _seededActiveItem = item;
    public void SeedSelectedItems(IEnumerable<ActiveItem> items)
    {
        _seededSelectedItems.Clear();
        _seededSelectedItems.AddRange(items);
    }
    public void ClearSelectedItems() => _seededSelectedItems.Clear();

    public Task<ActiveItem?> GetActiveItemAsync(CancellationToken cancellationToken = default)
    {
        Calls.Add(nameof(GetActiveItemAsync));
        ThrowIfInjected();
        if (OnGetActiveItem is not null) return Task.FromResult(OnGetActiveItem(cancellationToken));
        return Task.FromResult<ActiveItem?>(_seededActiveItem);
    }

    public Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top = 50,
        CancellationToken cancellationToken = default)
    {
        Calls.Add($"{nameof(GetSelectedItemsAsync)}:scope={scope},top={top}");
        ThrowIfInjected();

        // Validation (mimics OutlookService-Validation in [1,250])
        ValidationHelpers.ValidateRange(top, 1, 250, nameof(top));

        // Quelle: Callback (Override) oder seeded items
        IReadOnlyList<ActiveItem> source = OnGetSelectedItems is not null
            ? OnGetSelectedItems(scope, top, cancellationToken)
            : _seededSelectedItems.ToList();

        // Scope-Filter (mimics OutlookInteropAdapter Class-Dispatch + Filter)
        IEnumerable<ActiveItem> filtered = scope switch
        {
            SelectionScope.Mail => source.Where(i => i is ActiveMail),
            SelectionScope.Calendar => source.Where(i => i is ActiveEvent),
            _ => source,
        };

        // Top-Cap (mimics OutlookInteropAdapter .Take(top))
        return Task.FromResult<IReadOnlyList<ActiveItem>>(filtered.Take(top).ToList());
    }

    private void ThrowIfInjected()
    {
        if (InjectException is not null)
        {
            var ex = InjectException;
            InjectException = null;
            throw ex;
        }
    }

    private static bool IsHiddenName(string name) =>
        name.StartsWith("Hidden", StringComparison.OrdinalIgnoreCase) || name.Contains("RSS", StringComparison.OrdinalIgnoreCase);
}
