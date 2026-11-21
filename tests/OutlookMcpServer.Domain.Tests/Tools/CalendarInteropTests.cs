using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Interop;
using Xunit;

namespace OutlookMcpServer.Domain.Tests.Tools;

public sealed class CalendarInteropTests
{
    [Fact]
    public void RecurrencePattern_PreservesIntervalAndWeekdayMask()
    {
        var rule = OutlookInteropAdapter.MapRecurrencePattern(1, 2, 2 | 16, null, null);

        Assert.NotNull(rule);
        Assert.Equal("weekly", rule.Type);
        Assert.Equal(2, rule.Interval);
        Assert.Equal(new[] { "monday", "thursday" }, rule.DaysOfWeek);
        Assert.Null(rule.EndDate);
        Assert.Null(rule.Occurrences);
    }

    [Theory]
    [InlineData("2026-07-15T09:00:00", "Europe/Berlin", 7)]
    [InlineData("2026-01-15T09:00:00", "Europe/Berlin", 8)]
    [InlineData("2026-07-15T09:00:00", "W. Europe Standard Time", 7)]
    [InlineData("2026-07-15T09:00:00-04:00", "America/New_York", 13)]
    [InlineData("2026-07-15T09:00:00Z", "Europe/Berlin", 9)]
    public void InputTime_UsesZoneForWallTimeAndOffsetForExplicitInstant(string time, string zone, int utcHour)
    {
        var utc = OutlookInteropAdapter.ToCalendarUtc(new DateTimeTimeZone { DateTime = time, TimeZone = zone });
        Assert.Equal(DateTimeKind.Utc, utc.Kind);
        Assert.Equal(utcHour, utc.Hour);
        Assert.Equal("UTC", OutlookInteropAdapter.FormatCalendarUtc(utc).TimeZone);
    }

    [Theory]
    [InlineData("2026-03-29T02:30:00", "Europe/Berlin")]
    [InlineData("2026-10-25T02:30:00", "Europe/Berlin")]
    [InlineData("2026-07-15T09:00:00", "No/SuchZone")]
    [InlineData("tomorrow", "UTC")]
    public void InputTime_RejectsInvalidOrAmbiguousWallTime(string time, string zone)
    {
        var error = Assert.Throws<OutlookServiceException>(() =>
            OutlookInteropAdapter.ToCalendarUtc(new DateTimeTimeZone { DateTime = time, TimeZone = zone }));
        Assert.Equal(ErrorCode.InvalidInput, error.Code);
    }

    [Fact]
    public void CalendarCursor_SortsThenExpandsAndRestricts_AndReturnsOnlyOverlappingAppointments()
    {
        var spanning = new Appointment(26, At(8), At(10));
        var touching = new Appointment(26, At(12), At(14));
        var nonAppointment = new Appointment(43, At(9), At(11));
        var endingLate = new Appointment(26, At(11), At(14));
        var folder = new CalendarFolder(new CalendarItems(spanning, nonAppointment, endingLate, touching));

        var actual = OutlookInteropAdapter.EnumerateCalendarWindow(folder, At(9), At(12), "[AllDayEvent] = False", default).ToList();

        Assert.Equal(new object[] { spanning, endingLate }, actual.Select(row => row.Item));
        Assert.Equal(DateTimeKind.Utc, actual[0].Start.Kind);
        Assert.Equal(new[] { "sort", "expand", "restrict", "first", "next", "next", "next", "next" }, folder.Items.Calls);
        Assert.Contains("[Start] <", folder.Items.Filter);
        Assert.Contains("[End] >", folder.Items.Filter);
        Assert.Contains("[AllDayEvent] = False", folder.Items.Filter);
    }

    [Theory]
    [InlineData(0, EventType.SingleInstance)]
    [InlineData(1, EventType.SeriesMaster)]
    [InlineData(2, EventType.Occurrence)]
    [InlineData(3, EventType.Exception)]
    public void RecurrenceState_DistinguishesMasterOccurrencesAndExceptions(int value, EventType type) =>
        Assert.Equal(type, OutlookInteropAdapter.MapRecurrenceState(value));

    [Fact]
    public void Availability_BusyAllDayAndBoundarySpanningEventsBlockSlots()
    {
        var busyAllDay = (At(0), At(0).AddDays(1));
        Assert.True(OutlookInteropAdapter.IsBlockingAppointment(2, 3)); // received is not cancelled
        Assert.False(OutlookInteropAdapter.IsBlockingAppointment(0, 0));
        Assert.False(OutlookInteropAdapter.IsBlockingAppointment(2, 5));
        Assert.False(OutlookInteropAdapter.IsBlockingAppointment(2, 7));
        Assert.Empty(OutlookInteropAdapter.BuildMeetingCandidates(new[] { busyAllDay }, At(9), At(17), TimeSpan.FromMinutes(30), 5));

        var candidate = Assert.Single(OutlookInteropAdapter.BuildMeetingCandidates(
            new[] { (At(15), At(22)), (At(8), At(11)), (At(10), At(12)) }, At(9), At(17), TimeSpan.FromMinutes(60), 5));
        Assert.Equal("2026-07-15T12:00:00", candidate.Start.DateTime);
        Assert.Equal("2026-07-15T13:00:00", candidate.End.DateTime);
        Assert.Equal("UTC", candidate.Start.TimeZone);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Delete_RejectsUnsupportedOrganizerCancellationModesWithoutDeleting(bool sendCancellation)
    {
        var activeMeeting = new Appointment(26, At(9), At(10)) { MeetingStatus = 1 };
        var error = Assert.Throws<OutlookServiceException>(() => OutlookInteropAdapter.DeleteAppointment(activeMeeting, sendCancellation));
        Assert.Equal(ErrorCode.InvalidInput, error.Code);
        Assert.False(activeMeeting.Deleted);
    }

    [Fact]
    public void Delete_DoesNotDeleteMailByCalendarId()
    {
        var mail = new Appointment(43, At(9), At(10));
        var error = Assert.Throws<OutlookServiceException>(() => OutlookInteropAdapter.DeleteAppointment(mail, false));
        Assert.Equal(ErrorCode.EventNotFound, error.Code);
        Assert.False(mail.Deleted);
    }

    [Fact]
    public void Delete_DeletesLocalAppointmentWithoutSending()
    {
        var appointment = new Appointment(26, At(9), At(10));
        OutlookInteropAdapter.DeleteAppointment(appointment, false);
        Assert.True(appointment.Deleted);
    }

    [Theory]
    [InlineData(Response.TentativelyAccepted, 2, 57)]
    [InlineData(Response.Accepted, 3, 56)]
    [InlineData(Response.Declined, 4, 55)]
    public void Respond_UsesNoUIAndSendsReturnedMeetingItemWithComment(Response response, int responseValue, int responseClass)
    {
        var appointment = new Appointment(26, At(9), At(10)) { MeetingStatus = 3 };
        appointment.Reply.Class = responseClass;
        OutlookInteropAdapter.SendAppointmentResponse(appointment,
            new RespondToEventRequest { Response = response, Comment = "Got the invite" });
        Assert.Equal((responseValue, true, false), appointment.ResponseArgs);
        Assert.True(appointment.Reply.Sent);
        Assert.Equal("Got the invite", appointment.Reply.Body);
    }

    [Fact]
    public void Respond_RejectsUnexpectedComItemWithoutSending()
    {
        var appointment = new Appointment(26, At(9), At(10)) { MeetingStatus = 3 };
        appointment.Reply.Class = 43;
        var error = Assert.Throws<OutlookServiceException>(() => OutlookInteropAdapter.SendAppointmentResponse(appointment,
            new RespondToEventRequest { Response = Response.Accepted }));
        Assert.Equal(ErrorCode.InternalError, error.Code);
        Assert.False(appointment.Reply.Sent);
    }

    private static DateTime At(int hour) => new(2026, 7, 15, hour, 0, 0, DateTimeKind.Utc);

    public sealed record CalendarFolder(CalendarItems Items)
    {
        public int DefaultItemType => 1;
    }

    public sealed class CalendarItems(params Appointment[] items)
    {
        public List<string> Calls { get; } = [];
        public string? Filter { get; private set; }
        private int _index;
        public void Sort(string field, bool descending)
        {
            Assert.Equal("[Start]", field);
            Assert.False(descending);
            Calls.Add("sort");
        }
        public bool IncludeRecurrences
        {
            set
            {
                Assert.True(value);
                Calls.Add("expand");
            }
        }
        public CalendarItems Restrict(string filter)
        {
            Calls.Add("restrict");
            Filter = filter;
            return this;
        }
        public Appointment? GetFirst()
        {
            Calls.Add("first");
            _index = 0;
            return items.FirstOrDefault();
        }
        public Appointment? GetNext()
        {
            Calls.Add("next");
            return ++_index < items.Length ? items[_index] : null;
        }
    }

    public sealed record Appointment(int Class, DateTime StartUTC, DateTime EndUTC)
    {
        public int MeetingStatus { get; init; }
        public bool Deleted { get; private set; }
        public (int Response, bool NoUi, bool AdditionalTextDialog) ResponseArgs { get; private set; }
        public MeetingReply Reply { get; } = new();
        public void Delete() => Deleted = true;
        public MeetingReply Respond(int response, bool noUi, bool additionalTextDialog)
        {
            ResponseArgs = (response, noUi, additionalTextDialog);
            return Reply;
        }
    }

    public sealed class MeetingReply
    {
        public int Class { get; set; } = 56;
        public string? Body { get; set; }
        public bool Sent { get; private set; }
        public void Send() => Sent = true;
    }
}
