using System.Runtime.InteropServices;
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
            Recurrence = null,
            ICalUId = meta.ICalUId,
            CreatedDateTime = meta.Created,
            LastModifiedDateTime = meta.Modified,
            ChangeKey = meta.ChangeKey,
            Type = meta.Type,
        };
    }

    /// <summary>
    /// Liest und konvertiert den Appointment-Body in das angeforderte Format.
    /// </summary>
    private static ItemBody? TryMapAppointmentBody(dynamic item, BodyFormat? bodyFormat)
    {
        if (bodyFormat is null) return null;
        try
        {
            string bodyText = TryGetString(item, "Body") ?? string.Empty;
            string bodyHtml = TryGetString(item, "HTMLBody") ?? string.Empty;
            string? nativeSource = null;
            if (!string.IsNullOrEmpty(bodyHtml))
            {
                nativeSource = bodyHtml;
            }
            else if (!string.IsNullOrEmpty(bodyText))
            {
                nativeSource = bodyText;
            }
            if (nativeSource is null || nativeSource.Length == 0) return null;
            string htmlForConversion = nativeSource == bodyText
                ? $"<html><body><pre>{System.Net.WebUtility.HtmlEncode(bodyText)}</pre></body></html>"
                : nativeSource;
            var (ct, content) = HtmlBodyConverter.Convert(htmlForConversion, bodyFormat.Value);
            return new ItemBody { ContentType = ct, Content = content };
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Liest Start- und Endzeit eines Appointments als DateTimeTimeZone.
    /// </summary>
    private static AppointmentStartEnd TryMapAppointmentStartEnd(dynamic item)
    {
        DateTimeTimeZone? startDttz = null;
        DateTimeTimeZone? endDttz = null;
        try
        {
            var startDt = (DateTime)item.Start;
            var startTz = TryGetWindowsTimeZoneId(item, "StartTimeZone") ?? "UTC";
            startDttz = new DateTimeTimeZone
            {
                DateTime = startDt.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = startTz,
            };
        }
        catch { }
        try
        {
            var endDt = (DateTime)item.End;
            var endTz = TryGetWindowsTimeZoneId(item, "EndTimeZone") ?? "UTC";
            endDttz = new DateTimeTimeZone
            {
                DateTime = endDt.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = endTz,
            };
        }
        catch { }
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
            var status = (int)item.MeetingStatus;
            isCancelled = status == 2 || status == 3;
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
        try { created = (DateTimeOffset)item.CreationTime; } catch { }
        try { modified = (DateTimeOffset)item.LastModificationTime; } catch { }

        string? changeKey = null;
        try { changeKey = TryGetString(item, "EntryID"); } catch { }

        var type = EventType.SingleInstance;
        try
        {
            if ((bool)item.IsRecurring)
            {
                type = EventType.SeriesMaster;
            }
        }
        catch { }

        return new AppointmentMetadata(hasAttachments, iCalUId, created, modified, changeKey, type);
    }

    /// <summary>
    /// Liest die Windows-TimeZone-ID aus einem Outlook-TimeZone-Objekt.
    /// </summary>
    private static string? TryGetWindowsTimeZoneId(dynamic item, string propertyName)
    {
        try
        {
            var tz = item.GetType().GetProperty(propertyName)?.GetValue(item);
            if (tz is null) return null;
            try
            {
                return (string?)tz.GetType().GetProperty("ID")?.GetValue(tz);
            }
            finally
            {
                ReleaseComObject(tz);
            }
        }
        catch { return null; }
    }
}
