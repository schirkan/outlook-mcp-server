using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Domain.Abstractions;

/// <summary>
/// Oeffentliches Service-Interface. Wird von den MCP-Tools
/// (<c>MailTools</c>, <c>CalendarTools</c>) per DI aufgerufen.
/// Methoden-Signaturen 1:1 zu <c>specs/API-DESIGN.md</c>.
/// </summary>
public interface IOutlookService
{
    // ===== Mail: Ordner =====

    Task<PagedResult<MailFolder>> ListMailFoldersAsync(
        string? parentFolderId = null,
        bool includeHidden = false,
        CancellationToken cancellationToken = default);

    Task<MailFolder> GetMailFolderAsync(
        string folderId,
        CancellationToken cancellationToken = default);

    // ===== Mail: Nachrichten =====

    Task<PagedResult<MailMessage>> ListMailsAsync(
        string folderId,
        int top = 25,
        int skip = 0,
        string? filter = null,
        string? search = null,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default);

    Task<MailMessage> GetMailAsync(
        string id,
        bool includeBody = true,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Liest mehrere Mails in einem Aufruf (Bulk-Variante von <c>get_mail</c>).
    /// Vermeidet N round-trips wenn ein Client eine Liste EntryIDs (typischerweise
    /// von <c>list_mails</c> oder <c>search_mails</c>) hat und zu jeder Mail
    /// Body/Metadata braucht.
    /// <para>
    /// Einzelne ungueltige IDs werden NICHT als Exception propagiert, sondern
    /// in <see cref="BulkMailResult.NotFoundIds"/> gelistet. Valides Empty-Result
    /// wenn keine ID gefunden wurde. IDs werden vor dem Adapter-Aufruf dedupliziert
    /// (Caller-Liste darf Duplikate enthalten).
    /// </para>
    /// </summary>
    Task<BulkMailResult> GetMailsAsync(
        IReadOnlyList<string> ids,
        bool includeBody = false,
        BodyFormat bodyFormat = BodyFormat.Markdown,
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

    /// <summary>
    /// Rekursive Ordner-Suche mit DASL-Property-Filter ueber die in <paramref name="scope"/>
    /// angegebenen Well-Known-Mailordner plus deren Unterordner. Liefert
    /// neueste zuerst, dedupliziert per EntryID. Hard-Cap <paramref name="top"/>
    /// (1-100). Pagination wird nicht unterstuetzt (NextSkip=null).
    /// </summary>
    Task<PagedResult<MailMessage>> ListMailsRecursiveAsync(
        IReadOnlyList<string> scope,
        int top = 25,
        string? filter = null,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default);

    // ===== Mail: Mutationen =====

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

    // ===== Calendar: Kalender =====

    Task<IReadOnlyList<Calendar>> ListCalendarsAsync(
        CancellationToken cancellationToken = default);

    Task<Calendar> GetCalendarAsync(
        string id,
        CancellationToken cancellationToken = default);

    // ===== Calendar: Termine =====

    Task<PagedResult<CalendarEvent>> ListEventsAsync(
        string? calendarId,
        DateTimeTimeZone start,
        DateTimeTimeZone end,
        int top = 50,
        int skip = 0,
        string? filter = null,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default);

    Task<CalendarEvent> GetEventAsync(
        string id,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default);

    // ===== Calendar: Mutationen =====

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

    // ===== Active-Inspector / Selection (COM-only, kein Graph-Aequivalent) =====

    /// <summary>
    /// Liefert das aktuell im aktiven Outlook-Inspector-Fenster offene Item
    /// (MailItem oder AppointmentItem). null wenn kein Inspector offen oder
    /// Item-Typ nicht in v1-Scope (Tasks/Contacts sind v1.1).
    /// <paramref name="bodyFormat"/> steuert nur den Mail-Body von ActiveMail
    /// (AppointmentItem hat keinen Body).
    /// </summary>
    Task<ActiveItem?> GetActiveItemAsync(
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Liefert die im aktiven Outlook-Explorer-Fenster markierten Items,
    /// optional gefiltert nach scope (mail/calendar/any) und gecappt durch top (1-250).
    /// Selection.Count==0 ist valides Empty-Result (kein Fehler).
    /// Wirft OutlookNotActive wenn ActiveExplorer()==null.
    /// <paramref name="bodyFormat"/> steuert nur den Mail-Body der enthaltenen ActiveMails.
    /// </summary>
    Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top = 50,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default);
}