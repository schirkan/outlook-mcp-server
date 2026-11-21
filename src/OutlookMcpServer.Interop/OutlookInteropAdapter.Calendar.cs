using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    public async Task<IReadOnlyList<Calendar>> ListCalendarsAsync(
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var result = new List<Calendar>();
            dynamic session = GetMapiNamespace();
            try
            {
                string? defaultCalendarEntryId = null;
                try
                {
                    dynamic defaultCal = session.GetDefaultFolder(9);
                    try { defaultCalendarEntryId = (string)defaultCal.EntryID; } catch { }
                    ReleaseComObject(defaultCal);
                }
                catch { }

                foreach (var store in session.Stores)
                {
                    dynamic? storeCal = null;
                    try
                    {
                        try
                        {
                            storeCal = store.GetDefaultFolder(9);
                        }
                        catch
                        {
                            continue;
                        }

                        var isDefault = false;
                        try { isDefault = (string)storeCal.EntryID == defaultCalendarEntryId; } catch { }
                        var entryId = (string)storeCal.EntryID;
                        if (!result.Any(c => c.Id == entryId))
                        {
                            result.Add(MapCalendar(storeCal, isDefault));
                        }

                        foreach (var sub in storeCal.Folders)
                        {
                            try
                            {
                                int defaultItemType = 0;
                                try { defaultItemType = (int)sub.DefaultItemType; } catch { }
                                if (defaultItemType == 1)
                                {
                                    var subEntryId = (string)sub.EntryID;
                                    if (!result.Any(c => c.Id == subEntryId))
                                    {
                                        result.Add(MapCalendar(sub, false));
                                    }
                                }
                            }
                            finally
                            {
                                ReleaseComObject(sub);
                            }
                        }
                    }
                    finally
                    {
                        ReleaseComObject(storeCal);
                        ReleaseComObject(store);
                    }
                }
            }
            finally
            {
                ReleaseComObject(session);
            }
            return (IReadOnlyList<Calendar>)result;
        }, nameof(ListCalendarsAsync));
    }

    public async Task<Calendar> GetCalendarAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic folder = GetMapiNamespace().GetFolderFromID(id, Type.Missing);
            try
            {
                int defaultItemType = 0;
                try { defaultItemType = (int)folder.DefaultItemType; } catch { }
                if (defaultItemType != 1)
                {
                    throw new OutlookServiceException(
                        ErrorCode.CalendarNotFound,
                        $"Ordner mit EntryID '{id}' ist kein Kalender-Ordner (DefaultItemType={defaultItemType})");
                }
                return MapCalendar(folder, isDefault: false);
            }
            finally
            {
                ReleaseComObject(folder);
            }
        }, nameof(GetCalendarAsync));
    }

    public async Task<PagedResult<CalendarEvent>> ListEventsAsync(
        string? calendarId,
        DateTimeTimeZone start,
        DateTimeTimeZone end,
        int top = 50,
        int skip = 0,
        string? filter = null,
        CancellationToken cancellationToken = default)
    {
        var window = GetCalendarWindow(start, end);
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic calendar = calendarId is null
                ? GetMapiNamespace().GetDefaultFolder(9)
                : GetMapiNamespace().GetFolderFromID(calendarId, Type.Missing);
            try
            {
                var slice = new List<CalendarEvent>();
                var skipped = 0;
                foreach (var (item, _, _) in EnumerateCalendarWindow((object)calendar, window.Start, window.End, filter, cancellationToken))
                {
                    if (skipped < skip)
                    {
                        skipped++;
                        continue;
                    }
                    if (slice.Count == top)
                        return new PagedResult<CalendarEvent> { Value = slice, NextSkip = skip + top };
                    slice.Add(MapAppointmentItem(item, null));
                }
                return new PagedResult<CalendarEvent> { Value = slice };
            }
            finally
            {
                ReleaseComObject(calendar);
            }
        }, nameof(ListEventsAsync));
    }

    public async Task<CalendarEvent> GetEventAsync(
        string id,
        BodyFormat? bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic item = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                RequireAppointment((object?)item);
                return MapAppointmentItem((object)item, bodyFormat);
            }
            finally
            {
                ReleaseComObject(item);
            }
        }, nameof(GetEventAsync));
    }
}
