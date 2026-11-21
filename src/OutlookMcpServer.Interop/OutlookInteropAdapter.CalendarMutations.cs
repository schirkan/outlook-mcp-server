using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    public async Task<string> CreateEventAsync(
        CreateEventRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidateAppointmentBody(request.Body);
        var window = GetCalendarWindow(request.Start, request.End);
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic calendar = GetMapiNamespace().GetDefaultFolder(9);
            dynamic? items = null;
            dynamic? appointment = null;
            try
            {
                items = calendar.Items;
                appointment = items.Add(1); // olAppointmentItem
                RequireAppointment((object?)appointment);
                FillAppointmentFromCreateRequest(appointment, request, window.Start, window.End);
                if (request.Attendees.Count > 0)
                {
                    appointment.MeetingStatus = OlMeeting;
                    ReplaceAttendees((object)appointment, request.Attendees);
                }

                // EntryID is only guaranteed after Save; Send may invalidate the Outlook item reference.
                appointment.Save();
                string id = (string)appointment.EntryID;
                if (request.SendInvitations && request.Attendees.Count > 0)
                    appointment.Send();
                return id;
            }
            finally
            {
                ReleaseComObject(appointment);
                ReleaseComObject(items);
                ReleaseComObject(calendar);
            }
        }, nameof(CreateEventAsync));
    }

    public async Task UpdateEventAsync(
        string id,
        EventUpdate update,
        CancellationToken cancellationToken = default)
    {
        ValidateAppointmentBody(update.Body);
        DateTime? startUtc = update.Start is null ? null : ToCalendarUtc(update.Start);
        DateTime? endUtc = update.End is null ? null : ToCalendarUtc(update.End);
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic appointment = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                RequireAppointment((object?)appointment);
                var status = (int)appointment.MeetingStatus;
                var willSend = update.SendUpdate && (update.Attendees?.Count ?? CountAppointmentAttendees((object)appointment)) > 0;
                if ((willSend || update.Attendees is not null) && status is not (OlMeeting or OlNonMeeting))
                    throw new OutlookServiceException(ErrorCode.InvalidInput, "Nur der Organisator eines aktiven Meetings kann Teilnehmer aendern oder Updates senden.");
                if (update.SendUpdate && update.Attendees is { Count: 0 } && status == OlMeeting)
                    throw new OutlookServiceException(ErrorCode.InvalidInput, "Outlook kann kein Meeting-Update ohne Teilnehmer senden.");

                if (startUtc.HasValue || endUtc.HasValue)
                {
                    SetAppointmentWindow((object)appointment,
                        startUtc ?? (DateTime)appointment.StartUTC,
                        endUtc ?? (DateTime)appointment.EndUTC);
                }
                if (update.Body is { } body)
                    appointment.Body = body.Content;
                if (update.Subject is { } subject)
                    appointment.Subject = subject;
                if (update.Location is { } location)
                    appointment.Location = location;
                if (update.ReminderMinutesBeforeStart is int rem)
                {
                    appointment.ReminderSet = true;
                    appointment.ReminderMinutesBeforeStart = rem;
                }
                if (update.Categories is not null)
                    appointment.Categories = string.Join(", ", update.Categories);
                if (update.ShowAs is { } showAs)
                    appointment.BusyStatus = OlEnumMappings.ToOlBusyStatus(showAs);
                if (update.Importance is { } importance)
                    appointment.Importance = OlEnumMappings.ToOlImportance(importance);
                if (update.Sensitivity is { } sensitivity)
                    appointment.Sensitivity = OlEnumMappings.ToOlSensitivity(sensitivity);
                if (update.Attendees is { } attendees)
                {
                    if (attendees.Count > 0 && status == OlNonMeeting)
                        appointment.MeetingStatus = OlMeeting;
                    ReplaceAttendees((object)appointment, attendees);
                }

                if (willSend)
                {
                    if ((int)appointment.MeetingStatus == OlNonMeeting)
                        appointment.MeetingStatus = OlMeeting;
                    appointment.Send();
                }
                else
                {
                    appointment.Save();
                }
                return true;
            }
            finally
            {
                ReleaseComObject(appointment);
            }
        }, nameof(UpdateEventAsync));
    }

    public async Task DeleteEventAsync(
        string id,
        bool sendCancellation = true,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic appointment = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                DeleteAppointment((object?)appointment, sendCancellation);
                return true;
            }
            finally
            {
                ReleaseComObject(appointment);
            }
        }, nameof(DeleteEventAsync));
    }

    internal static void DeleteAppointment(object? item, bool sendCancellation)
    {
        RequireAppointment(item);
        dynamic appointment = item!;
        var status = (int)appointment.MeetingStatus;
        if (status is not (OlNonMeeting or OlMeetingReceived or OlMeetingCanceled or OlMeetingReceivedAndCanceled))
        {
            // AppointmentItem.Delete has no cancellation control, and OOM exposes no CancelMeeting method.
            throw new OutlookServiceException(ErrorCode.InvalidInput, sendCancellation
                ? "Outlook COM kann keine Meeting-Absage senden. Bitte das Meeting in Outlook absagen."
                : "Outlook COM garantiert keine stille Loeschung eines eigenen Meetings. Bitte Outlook verwenden.");
        }
        appointment.Delete();
    }

    public async Task RespondToEventAsync(
        string id,
        RespondToEventRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic appointment = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                SendAppointmentResponse((object?)appointment, request);
                return true;
            }
            finally
            {
                ReleaseComObject(appointment);
            }
        }, nameof(RespondToEventAsync));
    }

    internal static void SendAppointmentResponse(object? item, RespondToEventRequest request)
    {
        RequireAppointment(item);
        dynamic appointment = item!;
        if ((int)appointment.MeetingStatus != OlMeetingReceived)
            throw new OutlookServiceException(ErrorCode.InvalidInput, "Nur auf eine aktive empfangene Einladung kann geantwortet werden.");
        var responseClass = request.Response switch
        {
            Response.Accepted => 56,            // olMeetingResponsePositive
            Response.TentativelyAccepted => 57, // olMeetingResponseTentative
            Response.Declined => 55,            // olMeetingResponseNegative
            _ => throw new OutlookServiceException(ErrorCode.InvalidInput, "Nur Zusage, vorlaeufige Zusage oder Ablehnung sind erlaubt."),
        };

        dynamic? response = appointment.Respond(OlEnumMappings.ToOlResponseStatus(request.Response), true, false);
        try
        {
            if (response is null || (int)response.Class != responseClass)
                throw new OutlookServiceException(ErrorCode.InternalError, "Outlook hat keine passende Meeting-Antwort erzeugt.");
            if (!string.IsNullOrEmpty(request.Comment))
                response.Body = request.Comment;
            response.Send();
        }
        finally { ReleaseComObject(response); }
    }

    public async Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(
        FindMeetingTimesRequest request,
        CancellationToken cancellationToken = default)
    {
        var window = GetCalendarWindow(request.TimeWindowStart, request.TimeWindowEnd);
        var duration = TimeSpan.FromMinutes(request.DurationMinutes);
        if (duration <= TimeSpan.Zero)
            throw new OutlookServiceException(ErrorCode.InvalidInput, "Die Meeting-Dauer muss positiv sein.");
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic calendar = GetMapiNamespace().GetDefaultFolder(9);
            try
            {
                var busy = new List<(DateTime Start, DateTime End)>();
                foreach (var (item, start, end) in EnumerateCalendarWindow((object)calendar, window.Start, window.End, null, cancellationToken))
                {
                    dynamic appointment = item;
                    if (IsBlockingAppointment((int)appointment.BusyStatus, (int)appointment.MeetingStatus))
                        busy.Add((start, end)); // Non-free all-day items block time just like timed items.
                }
                return BuildMeetingCandidates(busy, window.Start, window.End, duration, request.MaxCandidates);
            }
            finally
            {
                ReleaseComObject(calendar);
            }
        }, nameof(FindMeetingTimesAsync));
    }

    internal static bool IsBlockingAppointment(int busyStatus, int meetingStatus) =>
        busyStatus != 0 && !IsCancelledMeeting(meetingStatus); // olFree

    internal static IReadOnlyList<MeetingTimeCandidate> BuildMeetingCandidates(
        IEnumerable<(DateTime Start, DateTime End)> busy, DateTime windowStart, DateTime windowEnd,
        TimeSpan duration, int maxCandidates)
    {
        var candidates = new List<MeetingTimeCandidate>();
        var cursor = windowStart;
        foreach (var (start, end) in busy.OrderBy(b => b.Start))
        {
            if (!OverlapsCalendarWindow(start, end, windowStart, windowEnd)) continue;
            if (start - cursor >= duration)
                AddCandidate(cursor);
            if (end > cursor) cursor = end;
            if (cursor >= windowEnd || candidates.Count == maxCandidates) return candidates;
        }
        if (windowEnd - cursor >= duration && candidates.Count < maxCandidates)
            AddCandidate(cursor);
        return candidates;

        void AddCandidate(DateTime start) => candidates.Add(new MeetingTimeCandidate
        {
            Start = FormatCalendarUtc(start),
            End = FormatCalendarUtc(start + duration),
            Confidence = 100,
            ConflictsAttendeeCount = 0,
        });
    }

    private static void FillAppointmentFromCreateRequest(dynamic appointment, CreateEventRequest request, DateTime start, DateTime end)
    {
        appointment.Subject = request.Subject;
        if (request.Body is not null)
            appointment.Body = request.Body.Content;
        appointment.AllDayEvent = request.IsAllDay;
        SetAppointmentWindow((object)appointment, start, end);
        if (!string.IsNullOrEmpty(request.Location))
            appointment.Location = request.Location;
        if (request.ReminderMinutesBeforeStart is int rem)
        {
            appointment.ReminderSet = true;
            appointment.ReminderMinutesBeforeStart = rem;
        }
        if (request.Categories.Count > 0)
            appointment.Categories = string.Join(", ", request.Categories);
        appointment.BusyStatus = OlEnumMappings.ToOlBusyStatus(request.ShowAs);
        appointment.Importance = OlEnumMappings.ToOlImportance(request.Importance);
        appointment.Sensitivity = OlEnumMappings.ToOlSensitivity(request.Sensitivity);
    }

    private static void SetAppointmentWindow(dynamic appointment, DateTime start, DateTime end)
    {
        if (end <= start)
            throw new OutlookServiceException(ErrorCode.InvalidInput, "Das Terminende muss nach dem Start liegen.");
        var allDay = (bool)appointment.AllDayEvent;
        appointment.StartUTC = start;
        appointment.EndUTC = end;
        if (allDay != (bool)appointment.AllDayEvent || (DateTime)appointment.StartUTC != start || (DateTime)appointment.EndUTC != end)
            throw new OutlookServiceException(ErrorCode.InvalidInput, "Outlook kann das Terminzeitfenster nicht unveraendert abbilden.");
    }

    private static void ValidateAppointmentBody(ItemBody? body)
    {
        if (body is not null && body.ContentType != ItemBodyType.Text)
            throw new OutlookServiceException(ErrorCode.InvalidInput, "Outlook-Kalendertermine unterstuetzen nur Text als Body.");
    }

    private static int CountAppointmentAttendees(dynamic appointment)
    {
        dynamic recipients = appointment.Recipients;
        try
        {
            int attendees = 0;
            for (var i = (int)recipients.Count; i > 0; i--)
            {
                dynamic recipient = recipients.Item(i);
                try { if ((int)recipient.Type is 1 or 2 or 3) attendees++; }
                finally { ReleaseComObject(recipient); }
            }
            return attendees;
        }
        finally { ReleaseComObject(recipients); }
    }

    private static void ReplaceAttendees(dynamic appointment, IReadOnlyList<EventAttendeeInput> attendees)
    {
        dynamic recipients = appointment.Recipients;
        try
        {
            for (var i = (int)recipients.Count; i > 0; i--)
            {
                dynamic recipient = recipients.Item(i);
                try { if ((int)recipient.Type is 1 or 2 or 3) recipients.Remove(i); }
                finally { ReleaseComObject(recipient); }
            }
            foreach (var attendee in attendees)
            {
                dynamic recipient = recipients.Add(attendee.Email);
                try
                {
                    recipient.Type = attendee.Type switch
                    {
                        AttendeeType.Optional => 2,
                        AttendeeType.Resource => 3,
                        _ => 1,
                    };
                }
                finally { ReleaseComObject(recipient); }
            }
            if (attendees.Count != 0 && !(bool)recipients.ResolveAll())
                throw new OutlookServiceException(ErrorCode.InvalidInput, "Outlook konnte mindestens einen Meeting-Teilnehmer nicht aufloesen.");
        }
        finally { ReleaseComObject(recipients); }
    }
}
