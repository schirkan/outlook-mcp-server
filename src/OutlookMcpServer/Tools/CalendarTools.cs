using System.ComponentModel;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Tools;

/// <summary>
/// MCP-Tools fuer Calendar-Operationen. Naming: <c>snake_case</c>
/// (z. B. <c>list_events</c>, <c>get_event</c>, <c>create_event</c>).
/// Spezifikation siehe <c>specs/API-DESIGN.md</c>.
/// Subject + IDs werden geloggt, Body-Inhalt NIE.
/// </summary>
[McpServerToolType]
public sealed class CalendarTools
{
    private readonly IOutlookService _service;
    private readonly ILogger<CalendarTools> _logger;

    public CalendarTools(IOutlookService service, ILogger<CalendarTools> logger)
    {
        _service = service;
        _logger = logger;
    }

    // ===== Kalender =====

    [McpServerTool(Name = "list_calendars")]
    [Description("Listet alle Kalender im aktuellen Outlook-Profil — Id, Name, IsDefaultCalendar, CanEdit, Owner.")]
    public async Task<IReadOnlyList<Calendar>> ListCalendars(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("list_calendars");
        return await _service.ListCalendarsAsync(cancellationToken);
    }

    [McpServerTool(Name = "get_calendar")]
    [Description("Liest einen einzelnen Kalender per ID.")]
    public async Task<Calendar> GetCalendar(
        [Description("Kalender-ID (EntryID).")] string id,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("get_calendar id={Id}", id);
        return await _service.GetCalendarAsync(id, cancellationToken);
    }

    // ===== Termine =====

    [McpServerTool(Name = "list_events")]
    [Description("Listet Termine in einem Zeitfenster. Start/End als ISO-8601 + IANA-Zeitzone.")]
    public async Task<PagedResult<CalendarEvent>> ListEvents(
        [Description("Optional: Kalender-ID (null = Standardkalender).")] string? calendarId,
        [Description("Fenster-Start (ISO-8601, z. B. 2026-07-15T09:00:00).")] string startDateTime,
        [Description("IANA-Zeitzone des Start-Werts (z. B. Europe/Berlin).")] string startTimeZone,
        [Description("Fenster-Ende (ISO-8601).")] string endDateTime,
        [Description("IANA-Zeitzone des End-Werts.")] string endTimeZone,
        [Description("Max Anzahl (1-250, Default 50). Bei mehr Terminen im Fenster: skip = N * top fuer Seite N (Pagination).")] int top = 50,
        [Description("Skip-Count fuer Pagination (Default 0).")] int skip = 0,
        [Description("Optional: Filter-Ausdruck (z. B. 'isAllDay eq false').")] string? filter = null,
        [Description(@"Format des Termin-Body (siehe get_mail). Default 'markdown'.")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        _logger.LogInformation("list_events calendarId={CalendarId} window={Start}/{End} top={Top} skip={Skip} bodyFormat={Bf}",
            calendarId, startDateTime, endDateTime, top, skip, bf);
        var request = new ListEventsArgs(calendarId,
            new DateTimeTimeZone { DateTime = startDateTime, TimeZone = startTimeZone },
            new DateTimeTimeZone { DateTime = endDateTime, TimeZone = endTimeZone },
            top, skip, filter);
        return await _service.ListEventsAsync(request.CalendarId, request.Start, request.End,
            request.Top, request.Skip, request.Filter, bf, cancellationToken);
    }

    [McpServerTool(Name = "get_event")]
    [Description("Liest einen einzelnen Termin inkl. Body, Attendees, Reminder.")]
    public async Task<CalendarEvent> GetEvent(
        [Description("Termin EntryID.")] string id,
        [Description(@"Format des Termin-Body (siehe get_mail). Default 'markdown'.")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        _logger.LogInformation("get_event id={Id} bodyFormat={Bf}", id, bf);
        return await _service.GetEventAsync(id, bf, cancellationToken);
    }

    [McpServerTool(Name = "create_event")]
    [Description("Legt einen neuen Termin an. Erfordert AllowCreate=true. Bei 'sendInvitations' werden Attendees via Outlook-Standard-Meeting-Versand benachrichtigt.")]
    public async Task<string> CreateEvent(
        [Description("Subject (Termin-Titel).")] string subject,
        [Description("Start (ISO-8601, z. B. 2026-07-15T09:00:00).")] string startDateTime,
        [Description("IANA-Zeitzone des Start-Werts (z. B. Europe/Berlin).")] string startTimeZone,
        [Description("Ende (ISO-8601).")] string endDateTime,
        [Description("IANA-Zeitzone des End-Werts.")] string endTimeZone,
        [Description("Ganztaegig (Default false).")] bool isAllDay = false,
        [Description("Optional: Ort-Anzeigename (z. B. 'Konferenzraum 3.14').")] string? location = null,
        [Description("Optional: Body Inhalt.")] string? body = null,
        [Description("Optional: Body content type 'text' oder 'html' (Default text).")] string bodyContentType = "text",
        [Description("Optional: Komma-getrennte Attendees (Format: 'email:Type' oder nur 'email' fuer required). Type = required|optional|resource.")] string? attendees = null,
        [Description("Optional: Erinnerung in Minuten vor Start (null/0 = keine).")] int? reminderMinutesBeforeStart = null,
        [Description("Optional: Komma-getrennte Kategorien (z. B. 'Wichtig,Projekt').")] string? categories = null,
        [Description("showAs: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown' (Default busy).")] string showAs = "busy",
        [Description("Importance: 'low' | 'normal' | 'high' (Default normal).")] string importance = "normal",
        [Description("Sensitivity: 'normal' | 'personal' | 'private' | 'confidential' (Default normal).")] string sensitivity = "normal",
        [Description("Einladungen an Attendees senden (Default true).")] bool sendInvitations = true,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("create_event subject={Subject} start={Start} end={End} isAllDay={IsAllDay} attendeeCount={Count}",
            subject, startDateTime, endDateTime, isAllDay, CountCsv(attendees));
        var request = new CreateEventRequest
        {
            Subject = subject,
            Start = new DateTimeTimeZone { DateTime = startDateTime, TimeZone = startTimeZone },
            End = new DateTimeTimeZone { DateTime = endDateTime, TimeZone = endTimeZone },
            IsAllDay = isAllDay,
            Location = location,
            Body = body is null
                ? null
                : new ItemBody
                {
                    Content = body,
                    ContentType = string.Equals(bodyContentType, "html", StringComparison.OrdinalIgnoreCase) ? ItemBodyType.Html : ItemBodyType.Text,
                },
            Attendees = ParseAttendees(attendees),
            ReminderMinutesBeforeStart = reminderMinutesBeforeStart,
            Categories = SplitCsv(categories),
            ShowAs = ParseShowAs(showAs),
            Importance = ParseImportance(importance),
            Sensitivity = ParseSensitivity(sensitivity),
            SendInvitations = sendInvitations,
        };
        return await _service.CreateEventAsync(request, cancellationToken);
    }

    [McpServerTool(Name = "update_event")]
    [Description("PATCH auf Termin-Felder (Subject, Body, Start/End, Location, Attendees, Reminder, Categories, ShowAs, Importance, Sensitivity). Subject + IDs werden geloggt, Body-Inhalt NIE.")]
    public async Task UpdateEvent(
        [Description("Termin EntryID.")] string id,
        [Description("Optional: Neuer Subject.")] string? subject = null,
        [Description("Optional: Neuer Body Inhalt.")] string? body = null,
        [Description("Optional: Body content type 'text' oder 'html'.")] string bodyContentType = "text",
        [Description("Optional: Neuer Start (ISO-8601).")] string? startDateTime = null,
        [Description("Optional: IANA-Zeitzone des neuen Starts.")] string? startTimeZone = null,
        [Description("Optional: Neues Ende (ISO-8601).")] string? endDateTime = null,
        [Description("Optional: IANA-Zeitzone des neuen Endes.")] string? endTimeZone = null,
        [Description("Optional: Neuer Ort-Anzeigename.")] string? location = null,
        [Description("Optional: Neue Attendees als Komma-getrennte Liste (Format wie create_event). null = unveraendert, leerer String = leeren.")] string? attendees = null,
        [Description("Optional: Neue Erinnerung in Minuten (null = unveraendert).")] int? reminderMinutesBeforeStart = null,
        [Description("Optional: Neue Kategorien als Komma-getrennte Liste.")] string? categories = null,
        [Description("Optional: showAs-Wert.")] string? showAs = null,
        [Description("Optional: Importance-Wert.")] string? importance = null,
        [Description("Optional: Sensitivity-Wert.")] string? sensitivity = null,
        [Description("Update-Notification an Attendees senden (Default true).")] bool sendUpdate = true,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("update_event id={Id} subject={Subject}", id, subject);
        var update = new EventUpdate
        {
            Subject = subject,
            Body = body is null
                ? null
                : new ItemBody
                {
                    Content = body,
                    ContentType = string.Equals(bodyContentType, "html", StringComparison.OrdinalIgnoreCase) ? ItemBodyType.Html : ItemBodyType.Text,
                },
            Start = (startDateTime is null || startTimeZone is null)
                ? null
                : new DateTimeTimeZone { DateTime = startDateTime, TimeZone = startTimeZone },
            End = (endDateTime is null || endTimeZone is null)
                ? null
                : new DateTimeTimeZone { DateTime = endDateTime, TimeZone = endTimeZone },
            Location = location,
            Attendees = attendees is null ? null : ParseAttendees(attendees),
            ReminderMinutesBeforeStart = reminderMinutesBeforeStart,
            Categories = attendees is null && categories is null ? null : SplitCsv(categories),
            ShowAs = showAs is null ? null : ParseShowAs(showAs),
            Importance = importance is null ? null : ParseImportance(importance),
            Sensitivity = sensitivity is null ? null : ParseSensitivity(sensitivity),
            SendUpdate = sendUpdate,
        };
        await _service.UpdateEventAsync(id, update, cancellationToken);
    }

    [McpServerTool(Name = "delete_event")]
    [Description("Loescht einen Termin. sendCancellation=true → Cancel-Nachricht an Attendees (Default true). Erfordert AllowDelete=true.")]
    public async Task DeleteEvent(
        [Description("Termin EntryID.")] string id,
        [Description("Cancel-Notification an Attendees senden (Default true).")] bool sendCancellation = true,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("delete_event id={Id} sendCancellation={Send}", id, sendCancellation);
        await _service.DeleteEventAsync(id, sendCancellation, cancellationToken);
    }

    [McpServerTool(Name = "respond_to_event")]
    [Description("Antwortet auf eine Einladung (Accept / Tentatively / Decline). Beim Organizer selbst ohne Effekt.")]
    public async Task RespondToEvent(
        [Description("Termin EntryID.")] string id,
        [Description("Antwort: 'accepted' | 'tentativelyAccepted' | 'declined' | 'none'.")] string response,
        [Description("Optional: Antwort-Kommentar (wird in die Antwort-Email uebernommen).")] string? comment = null,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("respond_to_event id={Id} response={Response}", id, response);
        var request = new RespondToEventRequest
        {
            Response = ParseResponse(response),
            Comment = comment,
        };
        await _service.RespondToEventAsync(id, request, cancellationToken);
    }

    [McpServerTool(Name = "find_meeting_times")]
    [Description("Sucht freie Zeitfenster fuer ein Meeting (Outlook Best-Slot-Heuristik). Gibt maximal 'maxCandidates' Vorschlaege zurueck.")]
    public async Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimes(
        [Description("Gewuenschte Meeting-Dauer in Minuten.")] int durationMinutes,
        [Description("Suchfenster-Start (ISO-8601).")] string timeWindowStartDateTime,
        [Description("IANA-Zeitzone des Suchfenster-Starts.")] string timeWindowStartTimeZone,
        [Description("Suchfenster-Ende (ISO-8601).")] string timeWindowEndDateTime,
        [Description("IANA-Zeitzone des Suchfenster-Endes.")] string timeWindowEndTimeZone,
        [Description("Max Anzahl Kandidaten (Default 10).")] int maxCandidates = 10,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("find_meeting_times duration={Duration} window={Start}/{End} maxCandidates={Max}",
            durationMinutes, timeWindowStartDateTime, timeWindowEndDateTime, maxCandidates);
        var request = new FindMeetingTimesRequest
        {
            DurationMinutes = durationMinutes,
            TimeWindowStart = new DateTimeTimeZone { DateTime = timeWindowStartDateTime, TimeZone = timeWindowStartTimeZone },
            TimeWindowEnd = new DateTimeTimeZone { DateTime = timeWindowEndDateTime, TimeZone = timeWindowEndTimeZone },
            MaxCandidates = maxCandidates,
        };
        return await _service.FindMeetingTimesAsync(request, cancellationToken);
    }

    // ===== Helpers (private) =====

    private sealed record ListEventsArgs(string? CalendarId, DateTimeTimeZone Start, DateTimeTimeZone End, int Top, int Skip, string? Filter);

    private static IReadOnlyList<string> SplitCsv(string? value) =>
        string.IsNullOrWhiteSpace(value)
            ? Array.Empty<string>()
            : value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private static int CountCsv(string? value) =>
        string.IsNullOrWhiteSpace(value) ? 0 : value.Split(',', StringSplitOptions.RemoveEmptyEntries).Length;

    private static IReadOnlyList<EventAttendeeInput> ParseAttendees(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return Array.Empty<EventAttendeeInput>();
        var parts = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var result = new List<EventAttendeeInput>(parts.Length);
        foreach (var raw in parts)
        {
            var segments = raw.Split(':', 2, StringSplitOptions.TrimEntries);
            if (segments.Length == 0 || string.IsNullOrWhiteSpace(segments[0])) continue;
            var email = segments[0];
            var typeToken = segments.Length > 1 ? segments[1] : null;
            AttendeeType type = typeToken?.ToLowerInvariant() switch
            {
                "optional" => AttendeeType.Optional,
                "resource" => AttendeeType.Resource,
                _ => AttendeeType.Required,
            };
            string? name = null;
            var atIdx = email.IndexOf('<');
            if (atIdx > 0 && email.EndsWith('>'))
            {
                name = email[..atIdx].Trim();
                email = email[(atIdx + 1)..^1];
            }
            result.Add(new EventAttendeeInput { Email = email, Name = name, Type = type });
        }
        return result;
    }

    private static Importance ParseImportance(string? value) => (value ?? "normal").Trim().ToLowerInvariant() switch
    {
        "low" => Importance.Low,
        "high" => Importance.High,
        _ => Importance.Normal,
    };

    private static Sensitivity ParseSensitivity(string? value) => (value ?? "normal").Trim().ToLowerInvariant() switch
    {
        "personal" => Sensitivity.Personal,
        "private" => Sensitivity.Private,
        "confidential" => Sensitivity.Confidential,
        _ => Sensitivity.Normal,
    };

    private static ShowAs ParseShowAs(string? value) => (value ?? "busy").Trim().ToLowerInvariant() switch
    {
        "free" => ShowAs.Free,
        "tentative" => ShowAs.Tentative,
        "oof" => ShowAs.OutOfOffice,
        "workingelsewhere" => ShowAs.WorkingElsewhere,
        "unknown" => ShowAs.Unknown,
        _ => ShowAs.Busy,
    };

    private static Response ParseResponse(string? value) => (value ?? "none").Trim().ToLowerInvariant() switch
    {
        "organizer" => Response.Organizer,
        "tentativelyaccepted" => Response.TentativelyAccepted,
        "accepted" => Response.Accepted,
        "declined" => Response.Declined,
        "notresponded" => Response.NotResponded,
        _ => Response.None,
    };
}
