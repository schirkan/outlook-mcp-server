using System.Globalization;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    private const int OlAppointment = 26;
    private const int OlNonMeeting = 0;
    private const int OlMeeting = 1;
    private const int OlMeetingReceived = 3;
    private const int OlMeetingCanceled = 5;
    private const int OlMeetingReceivedAndCanceled = 7;
    private const int MaxCalendarScan = 100_000;

    private static bool IsAppointment(dynamic? item) => item is not null && (int)item.Class == OlAppointment;

    private static void RequireAppointment(object? item)
    {
        if (!IsAppointment(item))
            throw new OutlookServiceException(ErrorCode.EventNotFound, "Die EntryID bezeichnet keinen Outlook-Termin.");
    }

    internal static bool IsCancelledMeeting(int status) => status is OlMeetingCanceled or OlMeetingReceivedAndCanceled;

    internal static EventType MapRecurrenceState(int state) => state switch
    {
        1 => EventType.SeriesMaster,  // olApptMaster
        2 => EventType.Occurrence,    // olApptOccurrence
        3 => EventType.Exception,     // olApptException
        _ => EventType.SingleInstance,
    };

    internal static IEnumerable<(object Item, DateTime Start, DateTime End)> EnumerateCalendarWindow(
        object folder, DateTime startUtc, DateTime endUtc, string? filter, CancellationToken cancellationToken)
    {
        dynamic calendar = folder;
        if ((int)calendar.DefaultItemType != 1) // olAppointmentItem
            throw new OutlookServiceException(ErrorCode.CalendarNotFound, "Der Ordner ist kein Outlook-Kalender.");

        dynamic items = calendar.Items;
        dynamic? restricted = null;
        try
        {
            items.Sort("[Start]", false);
            items.IncludeRecurrences = true;
            restricted = items.Restrict(BuildCalendarRestriction(startUtc, endUtc, filter));

            // Count is undefined after expanding rolling recurrences; only walk the restricted cursor.
            var scanned = 0;
            dynamic? item = restricted.GetFirst();
            while (item is not null)
            {
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (++scanned > MaxCalendarScan)
                        throw new OutlookServiceException(ErrorCode.InvalidInput,
                            "Zu viele Kalender-Eintraege fuer eine sichere Abfrage. Bitte das Zeitfenster verkleinern.");

                    if (IsAppointment((object)item))
                    {
                        var start = DateTime.SpecifyKind((DateTime)item.StartUTC, DateTimeKind.Utc);
                        var end = DateTime.SpecifyKind((DateTime)item.EndUTC, DateTimeKind.Utc);
                        if (OverlapsCalendarWindow(start, end, startUtc, endUtc))
                            yield return ((object)item, start, end);
                    }
                }
                finally { ReleaseComObject(item); }

                item = restricted.GetNext();
            }
        }
        finally
        {
            if (!ReferenceEquals(restricted, items)) ReleaseComObject(restricted);
            ReleaseComObject(items);
        }
    }

    private static string BuildCalendarRestriction(DateTime startUtc, DateTime endUtc, string? filter)
    {
        // Jet dates are local, with Windows regional formatting and no seconds. Allow for
        // DST/rounding and for Outlook's display zone differing from the Windows zone.
        // The iterator checks exact UTC overlap, so this padding cannot add results.
        var localStart = startUtc.ToLocalTime();
        var localEnd = endUtc.ToLocalTime();
        const long padding = 2 * TimeSpan.TicksPerDay;
        var start = localStart.AddTicks(-Math.Min(padding, localStart.Ticks));
        var end = localEnd.AddTicks(Math.Min(padding, DateTime.MaxValue.Ticks - localEnd.Ticks));
        var restriction = $"[Start] < '{FormatDate(end)}' AND [End] > '{FormatDate(start)}'";
        return string.IsNullOrWhiteSpace(filter) ? restriction : $"({restriction}) AND ({filter})";

        static string FormatDate(DateTime value) => value.ToString("g", CultureInfo.CurrentCulture).Replace("'", "''");
    }

    internal static bool OverlapsCalendarWindow(DateTime start, DateTime end, DateTime windowStart, DateTime windowEnd) =>
        end > start && start < windowEnd && end > windowStart;

    private static (DateTime Start, DateTime End) GetCalendarWindow(DateTimeTimeZone start, DateTimeTimeZone end)
    {
        var startUtc = ToCalendarUtc(start);
        var endUtc = ToCalendarUtc(end);
        if (endUtc <= startUtc)
            throw new OutlookServiceException(ErrorCode.InvalidInput, "Das Ende muss nach dem Start liegen.");
        return (startUtc, endUtc);
    }

    internal static DateTime ToCalendarUtc(DateTimeTimeZone value)
    {
        string[] localFormats = ["yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF"];
        try
        {
            var zone = string.IsNullOrWhiteSpace(value.TimeZone) ? TimeZoneInfo.Utc : TimeZoneInfo.FindSystemTimeZoneById(value.TimeZone);
            var text = value.DateTime.Trim();
            if (DateTime.TryParseExact(text, localFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var local))
            {
                if (zone.IsAmbiguousTime(local))
                    throw new OutlookServiceException(ErrorCode.InvalidInput, "Mehrdeutige lokale Uhrzeit: Bitte einen expliziten UTC-Offset angeben.");
                return TimeZoneInfo.ConvertTimeToUtc(local, zone);
            }

            var offsetFormats = Array.ConvertAll(localFormats, f => f + "zzz");
            if (text.EndsWith('Z')) text = text[..^1] + "+00:00";
            if (DateTimeOffset.TryParseExact(text, offsetFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var instant))
                return instant.UtcDateTime;
        }
        catch (Exception ex) when (ex is ArgumentException or TimeZoneNotFoundException or InvalidTimeZoneException)
        {
            throw new OutlookServiceException(ErrorCode.InvalidInput, "Ungueltige Datumszeit oder Zeitzone.", ex);
        }
        throw new OutlookServiceException(ErrorCode.InvalidInput, "Datumszeit muss ISO-8601 entsprechen.");
    }

    internal static DateTimeTimeZone FormatCalendarUtc(DateTime utc) => new()
    {
        DateTime = utc.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture),
        TimeZone = "UTC",
    };
}
