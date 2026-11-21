using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    private record AppointmentStartEnd(DateTimeTimeZone? Start, DateTimeTimeZone? End);
    private record AppointmentLocation(Location? Location, List<Location> Locations);
    private record AppointmentFlags(Importance Importance, Sensitivity Sensitivity, ShowAs ShowAs, bool IsCancelled);
    private record AppointmentReminder(bool IsOn, int MinutesBeforeStart);
    private record AppointmentMetadata(bool HasAttachments, string? ICalUId, DateTimeOffset? Created, DateTimeOffset? Modified, string? ChangeKey, EventType Type);

    /// <summary>
    /// Mappt einen Outlook-Kalenderordner auf das Calendar-DTO.
    /// </summary>
    private static Calendar MapCalendar(dynamic folder, bool isDefault)
    {
        var entryId = (string)folder.EntryID;
        var name = (string)folder.Name;
        bool canEdit = true;
        try { canEdit = (int)folder.Permission != 0; } catch { }
        string? owner = null;
        try
        {
            dynamic parentStore = folder.Parent;
            try
            {
                if (parentStore is not null)
                {
                    dynamic propAcc = parentStore.PropertyAccessor;
                    try
                    {
                        owner = (string?)propAcc.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A1F001F");
                    }
                    finally { ReleaseComObject(propAcc); }
                }
            }
            finally { ReleaseComObject(parentStore); }
        }
        catch { }

        return new Calendar
        {
            Id = entryId,
            Name = name,
            IsDefaultCalendar = isDefault,
            CanEdit = canEdit,
            Owner = owner,
        };
    }

    /// <summary>
    /// Mappt ein Outlook-Appointment auf das CalendarEvent-DTO.
    /// </summary>
    private static CalendarEvent MapAppointmentItem(dynamic item, BodyFormat? bodyFormat)
    {
        var entryId = (string)item.EntryID;

        var body = TryMapAppointmentBody(item, bodyFormat);
        var startEnd = TryMapAppointmentStartEnd(item);
        var loc = TryMapAppointmentLocation(item);
        var organizer = TryMapAppointmentOrganizer(item);
        var attendees = TryMapAppointmentAttendees(item);
        var flags = TryMapAppointmentFlags(item);
        var reminder = TryMapAppointmentReminder(item);
        var categories = TryMapAppointmentCategories(item);
        var meta = TryMapAppointmentMetadata(item);

        return new CalendarEvent
        {
            Id = entryId,
            Subject = TryGetString(item, "Subject"),
            BodyPreview = TryGetString(item, "BodyPreview"),
            Body = body,
            Start = startEnd.Start,
            End = startEnd.End,
            IsAllDay = TryGetAppointmentAllDay(item),
            Location = loc.Location,
            Locations = loc.Locations,
            Organizer = organizer,
            Attendees = attendees,
            Importance = flags.Importance,
            Sensitivity = flags.Sensitivity,
            ShowAs = flags.ShowAs,
            IsCancelled = flags.IsCancelled,
            IsReminderOn = reminder.IsOn,
            ReminderMinutesBeforeStart = reminder.MinutesBeforeStart,
            Categories = categories,
            HasAttachments = meta.HasAttachments,
            Recurrence = meta.Type == EventType.SeriesMaster ? TryMapRecurrencePattern(item) : null,
            ICalUId = meta.ICalUId,
            CreatedDateTime = meta.Created,
            LastModifiedDateTime = meta.Modified,
            ChangeKey = meta.ChangeKey,
            Type = meta.Type,
        };
    }

    private static RecurrencePattern? TryMapRecurrencePattern(dynamic appointment)
    {
        dynamic? pattern = null;
        try
        {
            pattern = appointment.GetRecurrencePattern();
            int recurrenceType = (int)pattern.RecurrenceType;
            bool ends = !(bool)pattern.NoEndDate;
            return MapRecurrencePattern(recurrenceType, (int)pattern.Interval,
                recurrenceType is 1 or 3 or 6 ? (int)pattern.DayOfWeekMask : 0,
                ends ? (DateTime)pattern.PatternEndDate : null,
                ends ? (int)pattern.Occurrences : null);
        }
        catch (System.Runtime.InteropServices.COMException) { return null; }
        finally { ReleaseComObject(pattern); }
    }

    internal static RecurrencePattern? MapRecurrencePattern(int type, int interval, int dayMask, DateTime? end, int? count)
    {
        string? name = type switch
        {
            0 => "daily", 1 => "weekly", 2 => "absoluteMonthly", 3 => "relativeMonthly",
            5 => "absoluteYearly", 6 => "relativeYearly", _ => null,
        };
        if (name is null) return null;
        string[] weekdays = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
        return new RecurrencePattern
        {
            Type = name,
            Interval = interval,
            DaysOfWeek = dayMask == 0 ? null : weekdays.Where((_, i) => (dayMask & (1 << i)) != 0).ToArray(),
            EndDate = end is { } date ? new DateTimeOffset(DateTime.SpecifyKind(date, DateTimeKind.Unspecified), TimeSpan.Zero) : null,
            Occurrences = count is > 0 ? count : null,
        };
    }

    /// <summary>
    /// Liest und konvertiert den Appointment-Body in das angeforderte Format.
    /// </summary>
    private static ItemBody? TryMapAppointmentBody(dynamic item, BodyFormat? bodyFormat)
    {
        if (bodyFormat is null) return null;
        string? text = TryGetString(item, "Body");
        if (string.IsNullOrEmpty(text)) return null;
        if (bodyFormat == BodyFormat.Html)
            return new ItemBody { ContentType = ItemBodyType.Html, Content = $"<pre>{System.Net.WebUtility.HtmlEncode(text)}</pre>" };
        return new ItemBody { ContentType = ItemBodyType.Text, Content = text };
    }

    /// <summary>
    /// Liest Start- und Endzeit eines Appointments als DateTimeTimeZone.
    /// </summary>
    private static AppointmentStartEnd TryMapAppointmentStartEnd(dynamic item)
    {
        DateTimeTimeZone? startDttz = null;
        DateTimeTimeZone? endDttz = null;
        try { startDttz = FormatCalendarUtc((DateTime)item.StartUTC); } catch { }
        try { endDttz = FormatCalendarUtc((DateTime)item.EndUTC); } catch { }
        return new AppointmentStartEnd(startDttz, endDttz);
    }

    /// <summary>
    /// Liest das All-Day-Flag eines Appointments defensiv aus.
    /// </summary>
    private static bool TryGetAppointmentAllDay(dynamic item)
    {
        try { return (bool)item.AllDayEvent; } catch { return false; }
    }

    /// <summary>
    /// Mappt den Outlook-Ort auf Single- und Collection-Location-Felder.
    /// </summary>
    private static AppointmentLocation TryMapAppointmentLocation(dynamic item)
    {
        var locations = new List<Location>();
        try
        {
            var locName = TryGetString(item, "Location");
            if (!string.IsNullOrEmpty(locName))
            {
                var loc = new Location { DisplayName = locName };
                locations.Add(loc);
                return new AppointmentLocation(loc, locations);
            }
        }
        catch { }
        return new AppointmentLocation(null, locations);
    }

    /// <summary>
    /// Liest den Organizer eines Appointments, sofern vorhanden.
    /// </summary>
    private static Organizer? TryMapAppointmentOrganizer(dynamic item)
    {
        try
        {
            dynamic org = item.Organizer;
            if (org is not null)
            {
                try
                {
                    var orgName = TryGetString(org, "Name");
                    var orgAddr = TryGetString(org, "Address");
                    if (!string.IsNullOrEmpty(orgAddr))
                    {
                        return new Organizer
                        {
                            EmailAddress = new EmailAddress { Name = orgName, Address = orgAddr },
                        };
                    }
                }
                finally { ReleaseComObject(org); }
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Mappt die Outlook-Teilnehmer eines Appointments auf Domain-Attendees.
    /// </summary>
    private static List<Attendee> TryMapAppointmentAttendees(dynamic item)
    {
        var attendees = new List<Attendee>();
        try
        {
            foreach (var att in item.Recipients)
            {
                try
                {
                    var attName = TryGetString(att, "Name");
                    var attAddr = TryGetString(att, "Address");
                    if (string.IsNullOrEmpty(attAddr)) continue;

                    int responseStatus = 0;
                    try { responseStatus = (int)att.MeetingResponseStatus; } catch { }
                    int typeRaw = 0;
                    try { typeRaw = (int)att.Type; } catch { }
                    var attType = typeRaw switch
                    {
                        2 => AttendeeType.Optional,
                        3 => AttendeeType.Resource,
                        _ => AttendeeType.Required,
                    };

                    attendees.Add(new Attendee
                    {
                        EmailAddress = new EmailAddress { Name = attName, Address = attAddr },
                        Type = attType,
                        Status = new ResponseStatus { Response = OlEnumMappings.ToResponse(responseStatus) },
                    });
                }
                finally { ReleaseComObject(att); }
            }
        }
        catch { }
        return attendees;
    }

    /// <summary>
    /// Liest Status-Flags wie Importance, Sensitivity, ShowAs und Cancelled.
    /// </summary>
    private static AppointmentFlags TryMapAppointmentFlags(dynamic item)
    {
        var importance = Importance.Normal;
        try { importance = OlEnumMappings.ToImportance((int)item.Importance); } catch { }
        var sensitivity = Sensitivity.Normal;
        try { sensitivity = OlEnumMappings.ToSensitivity((int)item.Sensitivity); } catch { }
        var showAs = ShowAs.Busy;
        try { showAs = OlEnumMappings.ToShowAs((int)item.BusyStatus); } catch { }

        bool isCancelled = false;
        try
        {
            isCancelled = IsCancelledMeeting((int)item.MeetingStatus);
        }
        catch { }

        return new AppointmentFlags(importance, sensitivity, showAs, isCancelled);
    }

    /// <summary>
    /// Liest Reminder-Einstellungen eines Appointments defensiv aus.
    /// </summary>
    private static AppointmentReminder TryMapAppointmentReminder(dynamic item)
    {
        bool isOn = false;
        try { isOn = (bool)item.ReminderSet; } catch { }
        int minutes = 0;
        try { minutes = (int)item.ReminderMinutesBeforeStart; } catch { }
        return new AppointmentReminder(isOn, minutes);
    }

    /// <summary>
    /// Zerlegt Outlook-Kategorien in eine bereinigte String-Liste.
    /// </summary>
    private static List<string> TryMapAppointmentCategories(dynamic item)
    {
        var categories = new List<string>();
        try
        {
            var catStr = TryGetString(item, "Categories") ?? string.Empty;
            if (!string.IsNullOrEmpty(catStr))
            {
                foreach (var c in catStr.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    categories.Add(c);
                }
            }
        }
        catch { }
        return categories;
    }

    /// <summary>
    /// Liest Metadaten wie Attachment-Flag, UIDs und Zeitstempel eines Appointments.
    /// </summary>
    private static AppointmentMetadata TryMapAppointmentMetadata(dynamic item)
    {
        bool hasAttachments = false;
        try { hasAttachments = (int)item.Attachments.Count > 0; } catch { }

        string? iCalUId = null;
        try { iCalUId = TryGetString(item, "GlobalAppointmentID"); } catch { }

        DateTimeOffset? created = null, modified = null;
        try { created = ((DateTime)item.CreationTime).ToUniversalTime(); } catch { }
        try { modified = ((DateTime)item.LastModificationTime).ToUniversalTime(); } catch { }

        string? changeKey = null;
        try { changeKey = TryGetString(item, "EntryID"); } catch { }

        var type = EventType.SingleInstance;
        try { type = MapRecurrenceState((int)item.RecurrenceState); } catch { }

        return new AppointmentMetadata(hasAttachments, iCalUId, created, modified, changeKey, type);
    }

}
