using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Domain.Serialization;

/// <summary>
/// Source-generated <see cref="JsonSerializerContext"/> fuer alle Root-Typen,
/// die ueber MCP-Tool-Aufrufe serialisiert werden. Wird benoetigt, sobald
/// <c>PublishTrimmed=true</c> aktiv ist — der Trimmer kann sonst polymorphe
/// Typen (z. B. <see cref="ActiveItem"/>) oder DTOs, die nur ueber dynamische
/// Aufrufe erreicht werden, nicht konservieren. Ohne explizite Registrierung
/// kommt es zu <c>NotSupportedException: JsonTypeInfo metadata for type ... was not provided</c>.
/// </summary>
/// <remarks>
/// Wird ueber <c>WithToolsFromAssembly(serializerOptions: ...)</c> in
/// <c>Program.cs</c> via <c>TypeInfoResolverChain</c> an den MCP-SDK
/// <c>McpJsonUtilities.DefaultOptions</c> angehaengt.
/// </remarks>
[JsonSourceGenerationOptions(JsonSerializerDefaults.Web,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
// Mail
[JsonSerializable(typeof(MailFolder))]
[JsonSerializable(typeof(MailMessage))]
[JsonSerializable(typeof(Recipient))]
[JsonSerializable(typeof(EmailAddress))]
[JsonSerializable(typeof(ItemBody))]
[JsonSerializable(typeof(InternetMessageHeader))]
[JsonSerializable(typeof(AttachmentSummary))]
[JsonSerializable(typeof(AttachmentData))]
[JsonSerializable(typeof(InlineAttachment))]
[JsonSerializable(typeof(SendMailRequest))]
[JsonSerializable(typeof(SendMailResult))]
[JsonSerializable(typeof(BulkMailResult))]
[JsonSerializable(typeof(MailUpdate))]
// Calendar
[JsonSerializable(typeof(Calendar))]
[JsonSerializable(typeof(CalendarEvent))]
[JsonSerializable(typeof(EventUpdate))]
[JsonSerializable(typeof(CreateEventRequest))]
[JsonSerializable(typeof(RespondToEventRequest))]
[JsonSerializable(typeof(FindMeetingTimesRequest))]
[JsonSerializable(typeof(MeetingTimeCandidate))]
[JsonSerializable(typeof(Location))]
[JsonSerializable(typeof(Attendee))]
[JsonSerializable(typeof(Organizer))]
[JsonSerializable(typeof(EventAttendeeInput))]
[JsonSerializable(typeof(RecurrencePattern))]
[JsonSerializable(typeof(DateTimeTimeZone))]
// Common (Paged, ActiveSelection)
[JsonSerializable(typeof(PagedResult<MailFolder>))]
[JsonSerializable(typeof(PagedResult<MailMessage>))]
[JsonSerializable(typeof(PagedResult<CalendarEvent>))]
[JsonSerializable(typeof(SelectedItemsResult))]
// ReadOnlyList-Varianten — Tool-Methoden geben haeufig IReadOnlyList<T>
// zurueck (ListCalendars, MapRecipients etc.). Source-Gen erzeugt pro
// geschlossener Generic-Instanz exakt eine JsonTypeInfo.
[JsonSerializable(typeof(IReadOnlyList<Calendar>))]
[JsonSerializable(typeof(IReadOnlyList<CalendarEvent>))]
[JsonSerializable(typeof(IReadOnlyList<MailFolder>))]
[JsonSerializable(typeof(IReadOnlyList<MailMessage>))]
[JsonSerializable(typeof(IReadOnlyList<Recipient>))]
[JsonSerializable(typeof(IReadOnlyList<InternetMessageHeader>))]
[JsonSerializable(typeof(IReadOnlyList<AttachmentSummary>))]
[JsonSerializable(typeof(IReadOnlyList<AttachmentData>))]
[JsonSerializable(typeof(IReadOnlyList<InlineAttachment>))]
[JsonSerializable(typeof(IReadOnlyList<EventAttendeeInput>))]
[JsonSerializable(typeof(IReadOnlyList<ActiveItem>))]
[JsonSerializable(typeof(IReadOnlyList<string>))]
[JsonSerializable(typeof(IReadOnlyList<DateTimeTimeZone>))]
[JsonSerializable(typeof(IReadOnlyList<MeetingTimeCandidate>))]
// Konkrete Array-Typen fuer Tool-Parameter (MCP-SDK serialisiert nach
// deklariertem Parametertyp — string[] != IReadOnlyList<string> fuer
// die Source-Gen-Metadata). Beispiel: get_mails_by_ids(string[] ids).
[JsonSerializable(typeof(string[]))]
// ActiveItem ist polymorph — Basis + Sub-Typen explizit registrieren
[JsonSerializable(typeof(ActiveItem))]
[JsonSerializable(typeof(ActiveMail))]
[JsonSerializable(typeof(ActiveEvent))]
// Enums (koennten implizit gehen, aber explizit macht das Manifest selbsterklaerend)
[JsonSerializable(typeof(Importance))]
[JsonSerializable(typeof(Sensitivity))]
[JsonSerializable(typeof(ItemBodyType))]
[JsonSerializable(typeof(SelectionScope))]
[JsonSerializable(typeof(Response))]
[JsonSerializable(typeof(ShowAs))]
[JsonSerializable(typeof(AttendeeType))]
[JsonSerializable(typeof(EventType))]
public sealed partial class OutlookMcpJsonContext : JsonSerializerContext
{
}