using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Domain.Abstractions;

/// <summary>
/// Internes Adapter-Interface. Implementiert vom COM-Adapter
/// (<c>OutlookMcpServer.Interop.InteropOutlookAdapter</c>) in Karte 3.
/// <c>OutlookService</c> delegiert hierhin. Gleiche Methodensignaturen
/// wie <see cref="IOutlookService"/>.
/// </summary>
public interface IInteropOutlookAdapter
{
    Task<PagedResult<MailFolder>> ListMailFoldersAsync(
        string? parentFolderId = null,
        bool includeHidden = false,
        CancellationToken cancellationToken = default);

    Task<MailFolder> GetMailFolderAsync(
        string folderId,
        CancellationToken cancellationToken = default);

    Task<PagedResult<MailMessage>> ListMailsAsync(
        string folderId,
        int top = 25,
        int skip = 0,
        string? filter = null,
        string? search = null,
        CancellationToken cancellationToken = default);

    Task<MailMessage> GetMailAsync(
        string id,
        bool includeBody = true,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<InternetMessageHeader>> GetMailHeadersAsync(
        string id,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<AttachmentSummary>> ListAttachmentsAsync(
        string mailId,
        CancellationToken cancellationToken = default);

    Task<AttachmentData> GetAttachmentAsync(
        string mailId,
        string attachmentId,
        CancellationToken cancellationToken = default);

    Task<PagedResult<MailMessage>> SearchMailsAsync(
        string query,
        string? folderId = null,
        int top = 25,
        CancellationToken cancellationToken = default);

    Task<SendMailResult> SendMailAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default);

    Task<string> CreateDraftAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default);

    Task UpdateMailAsync(
        string id,
        MailUpdate update,
        CancellationToken cancellationToken = default);

    Task<string> MoveMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default);

    Task<string> CopyMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default);

    Task DeleteMailAsync(
        string id,
        bool permanent = false,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<Calendar>> ListCalendarsAsync(
        CancellationToken cancellationToken = default);

    Task<Calendar> GetCalendarAsync(
        string id,
        CancellationToken cancellationToken = default);

    Task<PagedResult<CalendarEvent>> ListEventsAsync(
        string? calendarId,
        DateTimeTimeZone start,
        DateTimeTimeZone end,
        int top = 50,
        int skip = 0,
        string? filter = null,
        CancellationToken cancellationToken = default);

    Task<CalendarEvent> GetEventAsync(
        string id,
        CancellationToken cancellationToken = default);

    Task<string> CreateEventAsync(
        CreateEventRequest request,
        CancellationToken cancellationToken = default);

    Task UpdateEventAsync(
        string id,
        EventUpdate update,
        CancellationToken cancellationToken = default);

    Task DeleteEventAsync(
        string id,
        bool sendCancellation = true,
        CancellationToken cancellationToken = default);

    Task RespondToEventAsync(
        string id,
        RespondToEventRequest request,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(
        FindMeetingTimesRequest request,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// COM-Mapping fuer <see cref="IOutlookService.GetActiveItemAsync"/>:
    /// <c>Application.ActiveInspector()?.CurrentItem</c> + Type-Dispatch.
    /// </summary>
    Task<ActiveItem?> GetActiveItemAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// COM-Mapping fuer <see cref="IOutlookService.GetSelectedItemsAsync"/>:
    /// <c>Application.ActiveExplorer()?.Selection</c> + Scope-Filter.
    /// ActiveExplorer()==null soll als <c>OutlookInteropException(OutlookNotActive, ...)</c>
    /// hochkommen.
    /// </summary>
    Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top = 50,
        CancellationToken cancellationToken = default);
}