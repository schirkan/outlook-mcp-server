using System.Reflection;
using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Exceptions;
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

    [Theory]
    [InlineData("garbage")]
    [InlineData("none")]
    [InlineData("organizer")]
    [InlineData("notResponded")]
    public async Task RespondToEvent_InvalidResponseIsRejected(string response)
    {
        var (tools, svc) = Create();
        var error = await Assert.ThrowsAsync<OutlookServiceException>(() => tools.RespondToEvent("EVT-1", response));
        Assert.Equal(ErrorCode.InvalidInput, error.Code);
        Assert.Empty(svc.Calls);
    }

    [Theory]
    [InlineData(null, null)]
    [InlineData("", "")]
    [InlineData("Wichtig, Projekt", "Wichtig|Projekt")]
    public async Task UpdateEvent_AttendeePatchDoesNotSetCategoriesUnlessExplicit(string? categories, string? expected)
    {
        var service = DispatchProxy.Create<IOutlookService, UpdateCapture>();
        var tools = new CalendarTools(service, NullLogger<CalendarTools>.Instance);
        await tools.UpdateEvent("EVT-1", attendees: "bob@example.com:optional", categories: categories, sendUpdate: false);

        var update = Assert.IsType<EventUpdate>(((UpdateCapture)(object)service).Update);
        Assert.Equal(AttendeeType.Optional, Assert.Single(update.Attendees!).Type);
        Assert.Equal(expected, update.Categories is null ? null : string.Join("|", update.Categories));
        Assert.False(update.SendUpdate);
    }

    [Fact]
    public async Task UpdateEvent_HtmlBodyOrIncompleteTimeIsRejectedBeforeService()
    {
        var (tools, svc) = Create();
        var html = await Assert.ThrowsAsync<OutlookServiceException>(() => tools.UpdateEvent("EVT-1", body: "<b>Hi</b>", bodyContentType: "html"));
        var time = await Assert.ThrowsAsync<OutlookServiceException>(() => tools.UpdateEvent("EVT-1", startDateTime: "2026-07-15T09:00:00"));
        Assert.Equal(ErrorCode.InvalidInput, html.Code);
        Assert.Equal(ErrorCode.InvalidInput, time.Code);
        Assert.Empty(svc.Calls);
    }

    public class UpdateCapture : DispatchProxy
    {
        public EventUpdate? Update { get; private set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            Assert.Equal(nameof(IOutlookService.UpdateEventAsync), targetMethod?.Name);
            Update = Assert.IsType<EventUpdate>(args![1]);
            return Task.CompletedTask;
        }
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
            filter: "[AllDayEvent] = False");
        Assert.NotNull(page);
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.ListEventsAsync)));
        Assert.Contains("cal=CAL-1", entry);
        Assert.Contains("top=10", entry);
        Assert.Contains("skip=5", entry);
        Assert.Contains("filter=[AllDayEvent] = False", entry);
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
