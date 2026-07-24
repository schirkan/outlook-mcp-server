using System.Runtime.InteropServices;
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
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic calendar = calendarId is null
                ? GetMapiNamespace().GetDefaultFolder(9)
                : GetMapiNamespace().GetFolderFromID(calendarId, Type.Missing);
            try
            {
                dynamic items = calendar.Items;
                try
                {
                    items.IncludeRecurrences = true;

                    var startDt = DateTime.Parse(start.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                    var endDt = DateTime.Parse(end.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                    var startStr = startDt.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    var endStr = endDt.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    var combined = $"[Start] >= '{startStr}' AND [End] <= '{endStr}'";
                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        combined = $"({combined}) AND ({filter})";
                    }
                    items = items.Restrict(combined);
                    items.Sort("[Start]");

                    int total = (int)items.Count;
                    int startIdx = Math.Min(skip, total);
                    int endIdx = Math.Min(skip + top, total);
                    var slice = new List<CalendarEvent>(endIdx - startIdx);
                    for (int i = startIdx; i < endIdx; i++)
                    {
                        dynamic item = items.Item(i + 1);
                        try
                        {
                            slice.Add(MapAppointmentItem(item, null));
                        }
                        finally
                        {
                            ReleaseComObject(item);
                        }
                    }
                    return new PagedResult<CalendarEvent>
                    {
                        Value = slice,
                        NextSkip = endIdx < total ? endIdx : (int?)null,
                    };
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
                return MapAppointmentItem(item, bodyFormat);
            }
            finally
            {
                ReleaseComObject(item);
            }
        }, nameof(GetEventAsync));
    }
}
