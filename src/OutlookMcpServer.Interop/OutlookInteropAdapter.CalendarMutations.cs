using System.Runtime.InteropServices;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    public async Task<string> CreateEventAsync(
        CreateEventRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic calendar = GetMapiNamespace().GetDefaultFolder(9);
            dynamic appointment = calendar.Items.Add(1);
            try
            {
                FillAppointmentFromCreateRequest(appointment, request);
                AddAttendees(appointment, request.Attendees);

                if (request.SendInvitations && request.Attendees.Count > 0)
                {
                    appointment.Send();
                }
                else
                {
                    appointment.Save();
                }

                return (string)appointment.EntryID;
            }
            finally
            {
                ReleaseComObject(appointment);
                ReleaseComObject(calendar);
            }
        }, nameof(CreateEventAsync));
    }

    public async Task UpdateEventAsync(
        string id,
        EventUpdate update,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic appointment = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                if (update.Body is { } body)
                {
                    var bodyFormat = OlEnumMappings.ToOlBodyFormat(body.ContentType);
                    appointment.BodyFormat = bodyFormat;
                    if (bodyFormat == 2)
                        appointment.HTMLBody = body.Content;
                    else
                        appointment.Body = body.Content;
                }
                if (update.Subject is { } subject)
                    appointment.Subject = subject;
                if (update.Start is { } start)
                    appointment.Start = DateTime.Parse(start.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                if (update.End is { } end)
                    appointment.End = DateTime.Parse(end.DateTime, System.Globalization.CultureInfo.InvariantCulture);
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
                if (update.Attendees is not null)
                {
                    dynamic recipients = appointment.Recipients;
                    try
                    {
                        while ((int)recipients.Count > 0)
                        {
                            ReleaseComObject(recipients.Item(1));
                            recipients.Remove(1);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(recipients);
                    }
                    AddAttendees(appointment, update.Attendees);
                }

                if (update.SendUpdate && (int)appointment.Recipients.Count > 0)
                {
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
                appointment.Delete();
                return true;
            }
            finally
            {
                ReleaseComObject(appointment);
            }
        }, nameof(DeleteEventAsync));
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
                var olResponse = OlEnumMappings.ToOlResponseStatus(request.Response);
                bool hasComment = !string.IsNullOrEmpty(request.Comment);
                appointment.Respond(olResponse, hasComment, request.Comment ?? string.Empty);
                return true;
            }
            finally
            {
                ReleaseComObject(appointment);
            }
        }, nameof(RespondToEventAsync));
    }

    public async Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(
        FindMeetingTimesRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var candidates = new List<MeetingTimeCandidate>();
            dynamic calendar = GetMapiNamespace().GetDefaultFolder(9);
            try
            {
                var windowStart = DateTime.Parse(request.TimeWindowStart.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                var windowEnd = DateTime.Parse(request.TimeWindowEnd.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                var duration = TimeSpan.FromMinutes(request.DurationMinutes);
                if (windowEnd <= windowStart || duration <= TimeSpan.Zero) return (IReadOnlyList<MeetingTimeCandidate>)candidates;

                dynamic items = calendar.Items;
                try
                {
                    items.IncludeRecurrences = true;
                    var startStr = windowStart.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    var endStr = windowEnd.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    items = items.Restrict($"[Start] >= '{startStr}' AND [End] <= '{endStr}'");
                    items.Sort("[Start]");

                    var busyList = new List<(DateTime Start, DateTime End)>();
                    for (int i = 1; i <= (int)items.Count; i++)
                    {
                        dynamic item = items.Item(i);
                        try
                        {
                            bool isAllDay = false;
                            try { isAllDay = (bool)item.AllDayEvent; } catch { }
                            if (isAllDay) continue;

                            int busyStatus = 0;
                            try { busyStatus = (int)item.BusyStatus; } catch { }
                            if (busyStatus == 0) continue;

                            var s = (DateTime)item.Start;
                            var e = (DateTime)item.End;
                            if (e <= s) continue;
                            busyList.Add((s, e));
                        }
                        finally
                        {
                            ReleaseComObject(item);
                        }
                    }

                    var sortedBusy = busyList.OrderBy(b => b.Start).ToList();

                    DateTime cursor = windowStart;
                    foreach (var (bs, be) in sortedBusy)
                    {
                        if (bs >= windowEnd) break;
                        if (bs > cursor)
                        {
                            AddCandidateIfFits(candidates, cursor, bs < windowEnd ? bs : windowEnd, duration);
                        }
                        if (be > cursor) cursor = be;
                        if (cursor >= windowEnd) break;
                    }
                    if (cursor < windowEnd)
                    {
                        AddCandidateIfFits(candidates, cursor, windowEnd, duration);
                    }
                }
                finally
                {
                    ReleaseComObject(items);
                }
            }
            finally
            {
                ReleaseComObject(calendar);
            }

            var ranked = candidates
                .OrderByDescending(c => (DateTime.Parse(c.End.DateTime, System.Globalization.CultureInfo.InvariantCulture)
                                       - DateTime.Parse(c.Start.DateTime, System.Globalization.CultureInfo.InvariantCulture)))
                .ThenBy(c => c.Start.DateTime)
                .Take(request.MaxCandidates)
                .ToList();
            return (IReadOnlyList<MeetingTimeCandidate>)ranked;
        }, nameof(FindMeetingTimesAsync));
    }

    private static void FillAppointmentFromCreateRequest(dynamic appointment, CreateEventRequest request)
    {
        appointment.Subject = request.Subject;
        if (request.Body is not null)
        {
            var bodyFormat = OlEnumMappings.ToOlBodyFormat(request.Body.ContentType);
            appointment.BodyFormat = bodyFormat;
            if (bodyFormat == 2)
                appointment.HTMLBody = request.Body.Content;
            else
                appointment.Body = request.Body.Content;
        }
        appointment.Start = DateTime.Parse(request.Start.DateTime, System.Globalization.CultureInfo.InvariantCulture);
        appointment.End = DateTime.Parse(request.End.DateTime, System.Globalization.CultureInfo.InvariantCulture);
        appointment.AllDayEvent = request.IsAllDay;
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

    private static void AddAttendees(dynamic appointment, IReadOnlyList<EventAttendeeInput> attendees)
    {
        if (attendees.Count == 0) return;
        foreach (var att in attendees)
        {
            dynamic recipient = appointment.Recipients.Add(att.Email);
            try
            {
                if (!string.IsNullOrEmpty(att.Name))
                {
                    try { recipient.Name = att.Name; } catch { }
                }
                int olType = att.Type switch
                {
                    AttendeeType.Optional => 2,
                    AttendeeType.Resource => 3,
                    _ => 1,
                };
                recipient.Type = olType;
                try { recipient.Resolve(); } catch { }
            }
            finally
            {
                ReleaseComObject(recipient);
            }
        }
        try { appointment.Recipients.ResolveAll(); } catch { }
    }

    private static void AddCandidateIfFits(
        List<MeetingTimeCandidate> candidates,
        DateTime start,
        DateTime end,
        TimeSpan duration)
    {
        if (end - start < duration) return;
        candidates.Add(new MeetingTimeCandidate
        {
            Start = new DateTimeTimeZone
            {
                DateTime = start.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            End = new DateTimeTimeZone
            {
                DateTime = (start + duration).ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            Confidence = 100,
            ConflictsAttendeeCount = 0,
        });
    }
}
