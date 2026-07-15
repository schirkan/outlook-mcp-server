using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;
using OutlookMcpServer.Domain.Tests.Fakes;
using OutlookMcpServer.Tools;
using Xunit;

namespace OutlookMcpServer.Domain.Tests.Tools;

/// <summary>
/// Tests fuer <see cref="CalendarTools"/>: CSV-Attendee-Parsing
/// (verschiedene Formate), ISO-DateTime+TZ-Mapping, Response-String->Enum,
/// ShowAs/Importance/Sensitivity-Mappings.
/// </summary>
public sealed class CalendarToolsTests
{
    private static (CalendarTools, FakeOutlookService) Create()
    {
        var svc = new FakeOutlookService();
        var tools = new CalendarTools(svc, NullLogger<CalendarTools>.Instance);
        return (tools, svc);
    }

    [Fact]
    public async Task CreateEvent_ParsesAttendeesAllForms()
    {
        var (tools, svc) = Create();
        await tools.CreateEvent(
            subject: "Workshop",
            startDateTime: "2026-07-15T09:00:00",
            startTimeZone: "Europe/Berlin",
            endDateTime: "2026-07-15T17:00:00",
            endTimeZone: "Europe/Berlin",
            attendees: "alice@example.com, bob@example.com:optional, Carol <carol@example.com>, room@example.com:resource",
            reminderMinutesBeforeStart: 15,
            categories: "Workshop, Q3",
            showAs: "tentative",
            importance: "high",
            sensitivity: "confidential",
            sendInvitations: false);

        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.CreateEventAsync)));
        Assert.Contains("attendeeCount=4", entry);
        Assert.Contains("subject=Workshop", entry);
    }

    [Fact]
    public async Task RespondToEvent_MapsAccepted()
    {
        var (tools, svc) = Create();
        await tools.RespondToEvent(id: "EVT-1", response: "accepted", comment: "Sounds good");
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.RespondToEventAsync)));
        Assert.Contains("response=Accepted", entry);
    }

    [Fact]
    public async Task RespondToEvent_MapsDeclined()
    {
        var (tools, svc) = Create();
        await tools.RespondToEvent(id: "EVT-1", response: "declined");
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.RespondToEventAsync)));
        Assert.Contains("response=Declined", entry);
    }

    [Fact]
    public async Task RespondToEvent_UnknownResponseMapsToNone()
    {
        var (tools, svc) = Create();
        await tools.RespondToEvent(id: "EVT-1", response: "garbage");
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.RespondToEventAsync)));
        Assert.Contains("response=None", entry);
    }

    [Fact]
    public async Task FindMeetingTimes_ForwardsRequest()
    {
        var (tools, svc) = Create();
        var result = await tools.FindMeetingTimes(
            durationMinutes: 60,
            timeWindowStartDateTime: "2026-07-15T09:00:00",
            timeWindowStartTimeZone: "Europe/Berlin",
            timeWindowEndDateTime: "2026-07-15T17:00:00",
            timeWindowEndTimeZone: "Europe/Berlin",
            maxCandidates: 5);
        Assert.NotNull(result);
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.FindMeetingTimesAsync)));
        Assert.Contains("duration=60", entry);
    }

    [Fact]
    public async Task ListEvents_ForwardsIsoDateTime()
    {
        var (tools, svc) = Create();
        var page = await tools.ListEvents(
            calendarId: "CAL-1",
            startDateTime: "2026-07-15T09:00:00",
            startTimeZone: "Europe/Berlin",
            endDateTime: "2026-07-15T17:00:00",
            endTimeZone: "Europe/Berlin",
            top: 10,
            skip: 5,
            filter: "isAllDay eq false");
        Assert.NotNull(page);
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.ListEventsAsync)));
        Assert.Contains("cal=CAL-1", entry);
        Assert.Contains("top=10", entry);
        Assert.Contains("skip=5", entry);
        Assert.Contains("filter=isAllDay eq false", entry);
    }

    [Fact]
    public async Task ListCalendars_ReturnsSeeded()
    {
        var (tools, svc) = Create();
        svc.SeedCalendar(SampleData.DefaultCalendar());
        var cals = await tools.ListCalendars();
        Assert.Single(cals);
        Assert.True(cals[0].IsDefaultCalendar);
    }

    [Fact]
    public async Task GetEvent_ReturnsSeeded()
    {
        var (tools, svc) = Create();
        svc.SeedEvent(SampleData.Event1());
        var evt = await tools.GetEvent(id: "EVT-1");
        Assert.Equal("Daily Standup", evt.Subject);
        Assert.Equal(ShowAs.Busy, evt.ShowAs);
    }

    [Fact]
    public async Task DeleteEvent_ForwardsSendCancellation()
    {
        var (tools, svc) = Create();
        await tools.DeleteEvent(id: "EVT-1", sendCancellation: false);
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.DeleteEventAsync)));
        Assert.Contains("sendCancellation=False", entry);
    }
}
