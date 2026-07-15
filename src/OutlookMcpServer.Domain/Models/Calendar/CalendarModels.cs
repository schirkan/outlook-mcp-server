using System.Text.Json.Serialization;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Domain.Models.Calendar;

public enum AttendeeType
{
    [JsonPropertyName("required")]
    Required,

    [JsonPropertyName("optional")]
    Optional,

    [JsonPropertyName("resource")]
    Resource,
}

public enum ShowAs
{
    [JsonPropertyName("free")]
    Free,

    [JsonPropertyName("tentative")]
    Tentative,

    [JsonPropertyName("busy")]
    Busy,

    [JsonPropertyName("oof")]
    OutOfOffice,

    [JsonPropertyName("workingElsewhere")]
    WorkingElsewhere,

    [JsonPropertyName("unknown")]
    Unknown,
}

public enum EventType
{
    [JsonPropertyName("singleInstance")]
    SingleInstance,

    [JsonPropertyName("occurrence")]
    Occurrence,

    [JsonPropertyName("exception")]
    Exception,

    [JsonPropertyName("seriesMaster")]
    SeriesMaster,
}

public enum Response
{
    [JsonPropertyName("none")]
    None,

    [JsonPropertyName("organizer")]
    Organizer,

    [JsonPropertyName("tentativelyAccepted")]
    TentativelyAccepted,

    [JsonPropertyName("accepted")]
    Accepted,

    [JsonPropertyName("declined")]
    Declined,

    [JsonPropertyName("notResponded")]
    NotResponded,
}

/// <summary>
/// Kalender-Eintrag. 1:1 zu Microsoft Graph <c>calendar</c>.
/// </summary>
public sealed record Calendar
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    [JsonPropertyName("isDefaultCalendar")]
    public bool IsDefaultCalendar { get; init; }

    [JsonPropertyName("canEdit")]
    public bool CanEdit { get; init; }

    [JsonPropertyName("owner")]
    public string? Owner { get; init; }
}

/// <summary>
/// Antwort-Status auf eine Einladung.
/// </summary>
public sealed record ResponseStatus
{
    [JsonPropertyName("response")]
    public Response Response { get; init; } = Response.NotResponded;

    [JsonPropertyName("time")]
    public DateTimeOffset? Time { get; init; }
}

/// <summary>
/// Ort eines Termins (Anzeige + optional Anschrift).
/// </summary>
public sealed record Location
{
    [JsonPropertyName("displayName")]
    public string DisplayName { get; init; } = string.Empty;

    [JsonPropertyName("address")]
    public string? Address { get; init; }
}

/// <summary>
/// Teilnehmer an einem Termin. <c>Status</c> enthaelt Antwort-Status.
/// </summary>
public sealed record Attendee
{
    [JsonPropertyName("emailAddress")]
    public EmailAddress EmailAddress { get; init; } = new();

    [JsonPropertyName("type")]
    public AttendeeType Type { get; init; } = AttendeeType.Required;

    [JsonPropertyName("status")]
    public ResponseStatus Status { get; init; } = new();
}

/// <summary>
/// Veranstalter eines Termins. 1:1 zu Microsoft Graph <c>Recipient</c>.
/// </summary>
public sealed record Organizer
{
    [JsonPropertyName("emailAddress")]
    public EmailAddress EmailAddress { get; init; } = new();
}

/// <summary>
/// Eingeladener Teilnehmer fuer <c>createEvent</c> (vereinfacht: nur E-Mail + Typ).
/// </summary>
public sealed record EventAttendeeInput
{
    [JsonPropertyName("email")]
    public string Email { get; init; } = string.Empty;

    [JsonPropertyName("name")]
    public string? Name { get; init; }

    [JsonPropertyName("type")]
    public AttendeeType Type { get; init; } = AttendeeType.Required;
}

/// <summary>
/// Wiederholungsmuster (vereinfacht fuer v1).
/// </summary>
public sealed record RecurrencePattern
{
    [JsonPropertyName("type")]
    public string Type { get; init; } = string.Empty; // daily, weekly, monthly, yearly

    [JsonPropertyName("interval")]
    public int Interval { get; init; } = 1;

    [JsonPropertyName("daysOfWeek")]
    public IReadOnlyList<string>? DaysOfWeek { get; init; }

    [JsonPropertyName("endDate")]
    public DateTimeOffset? EndDate { get; init; }

    [JsonPropertyName("occurrences")]
    public int? Occurrences { get; init; }
}

/// <summary>
/// Vollstaendiges Calendar-Event. Properties 1:1 zu Microsoft Graph <c>event</c>.
/// </summary>
public sealed record CalendarEvent
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("subject")]
    public string? Subject { get; init; }

    [JsonPropertyName("bodyPreview")]
    public string? BodyPreview { get; init; }

    [JsonPropertyName("body")]
    public ItemBody? Body { get; init; }

    [JsonPropertyName("start")]
    public DateTimeTimeZone? Start { get; init; }

    [JsonPropertyName("end")]
    public DateTimeTimeZone? End { get; init; }

    [JsonPropertyName("isAllDay")]
    public bool IsAllDay { get; init; }

    [JsonPropertyName("location")]
    public Location? Location { get; init; }

    [JsonPropertyName("locations")]
    public IReadOnlyList<Location> Locations { get; init; } = Array.Empty<Location>();

    [JsonPropertyName("organizer")]
    public Organizer? Organizer { get; init; }

    [JsonPropertyName("attendees")]
    public IReadOnlyList<Attendee> Attendees { get; init; } = Array.Empty<Attendee>();

    [JsonPropertyName("importance")]
    public Importance Importance { get; init; } = Importance.Normal;

    [JsonPropertyName("sensitivity")]
    public Sensitivity Sensitivity { get; init; } = Sensitivity.Normal;

    [JsonPropertyName("showAs")]
    public ShowAs ShowAs { get; init; } = ShowAs.Busy;

    [JsonPropertyName("isCancelled")]
    public bool IsCancelled { get; init; }

    [JsonPropertyName("isReminderOn")]
    public bool IsReminderOn { get; init; }

    [JsonPropertyName("reminderMinutesBeforeStart")]
    public int ReminderMinutesBeforeStart { get; init; }

    [JsonPropertyName("categories")]
    public IReadOnlyList<string> Categories { get; init; } = Array.Empty<string>();

    [JsonPropertyName("hasAttachments")]
    public bool HasAttachments { get; init; }

    [JsonPropertyName("recurrence")]
    public RecurrencePattern? Recurrence { get; init; }

    [JsonPropertyName("seriesMasterId")]
    public string? SeriesMasterId { get; init; }

    [JsonPropertyName("iCalUId")]
    public string? ICalUId { get; init; }

    [JsonPropertyName("webLink")]
    public string? WebLink { get; init; }

    [JsonPropertyName("createdDateTime")]
    public DateTimeOffset? CreatedDateTime { get; init; }

    [JsonPropertyName("lastModifiedDateTime")]
    public DateTimeOffset? LastModifiedDateTime { get; init; }

    [JsonPropertyName("changeKey")]
    public string? ChangeKey { get; init; }

    [JsonPropertyName("type")]
    public EventType Type { get; init; } = EventType.SingleInstance;
}

/// <summary>
/// Input fuer <c>createEvent</c>.
/// </summary>
public sealed record CreateEventRequest
{
    [JsonPropertyName("subject")]
    public string Subject { get; init; } = string.Empty;

    [JsonPropertyName("body")]
    public ItemBody? Body { get; init; }

    [JsonPropertyName("start")]
    public DateTimeTimeZone Start { get; init; } = new();

    [JsonPropertyName("end")]
    public DateTimeTimeZone End { get; init; } = new();

    [JsonPropertyName("isAllDay")]
    public bool IsAllDay { get; init; }

    [JsonPropertyName("location")]
    public string? Location { get; init; }

    [JsonPropertyName("attendees")]
    public IReadOnlyList<EventAttendeeInput> Attendees { get; init; } = Array.Empty<EventAttendeeInput>();

    [JsonPropertyName("reminderMinutesBeforeStart")]
    public int? ReminderMinutesBeforeStart { get; init; }

    [JsonPropertyName("categories")]
    public IReadOnlyList<string> Categories { get; init; } = Array.Empty<string>();

    [JsonPropertyName("showAs")]
    public ShowAs ShowAs { get; init; } = ShowAs.Busy;

    [JsonPropertyName("importance")]
    public Importance Importance { get; init; } = Importance.Normal;

    [JsonPropertyName("sensitivity")]
    public Sensitivity Sensitivity { get; init; } = Sensitivity.Normal;

    [JsonPropertyName("sendInvitations")]
    public bool SendInvitations { get; init; } = true;
}

/// <summary>
/// Input fuer <c>updateEvent</c> (PATCH-Semantik).
/// </summary>
public sealed record EventUpdate
{
    [JsonPropertyName("subject")]
    public string? Subject { get; init; }

    [JsonPropertyName("body")]
    public ItemBody? Body { get; init; }

    [JsonPropertyName("start")]
    public DateTimeTimeZone? Start { get; init; }

    [JsonPropertyName("end")]
    public DateTimeTimeZone? End { get; init; }

    [JsonPropertyName("location")]
    public string? Location { get; init; }

    [JsonPropertyName("attendees")]
    public IReadOnlyList<EventAttendeeInput>? Attendees { get; init; }

    [JsonPropertyName("reminderMinutesBeforeStart")]
    public int? ReminderMinutesBeforeStart { get; init; }

    [JsonPropertyName("categories")]
    public IReadOnlyList<string>? Categories { get; init; }

    [JsonPropertyName("showAs")]
    public ShowAs? ShowAs { get; init; }

    [JsonPropertyName("importance")]
    public Importance? Importance { get; init; }

    [JsonPropertyName("sensitivity")]
    public Sensitivity? Sensitivity { get; init; }

    [JsonPropertyName("sendUpdate")]
    public bool SendUpdate { get; init; } = true;
}

/// <summary>
/// Input fuer <c>respondToEvent</c>.
/// </summary>
public sealed record RespondToEventRequest
{
    [JsonPropertyName("response")]
    public Response Response { get; init; }

    [JsonPropertyName("comment")]
    public string? Comment { get; init; }
}

/// <summary>
/// Input fuer <c>findMeetingTimes</c>.
/// </summary>
public sealed record FindMeetingTimesRequest
{
    [JsonPropertyName("durationMinutes")]
    public int DurationMinutes { get; init; }

    [JsonPropertyName("timeWindowStart")]
    public DateTimeTimeZone TimeWindowStart { get; init; } = new();

    [JsonPropertyName("timeWindowEnd")]
    public DateTimeTimeZone TimeWindowEnd { get; init; } = new();

    [JsonPropertyName("maxCandidates")]
    public int MaxCandidates { get; init; } = 10;
}

/// <summary>
/// Kandidat-Fenster fuer <c>findMeetingTimes</c>.
/// </summary>
public sealed record MeetingTimeCandidate
{
    [JsonPropertyName("start")]
    public DateTimeTimeZone Start { get; init; } = new();

    [JsonPropertyName("end")]
    public DateTimeTimeZone End { get; init; } = new();

    [JsonPropertyName("confidence")]
    public int Confidence { get; init; }

    [JsonPropertyName("conflictsAttendeeCount")]
    public int ConflictsAttendeeCount { get; init; }
}