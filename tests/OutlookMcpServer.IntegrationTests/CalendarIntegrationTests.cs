namespace OutlookMcpServer.IntegrationTests;

/// <summary>
/// Integration-Tests fuer Calendar-Operationen gegen ein laufendes
/// klassisches Outlook-Profil.
/// </summary>
public sealed class CalendarIntegrationTests : OutlookIntegrationTestBase
{
    [SkippableFact]
    public async Task ListCalendars_ReturnsAtLeastOneCalendar()
    {
        SkipIfOutlookNotAvailable();

        // Act
        var result = await Adapter.ListCalendarsAsync(cancellationToken: CancellationToken.None);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        // Default-Calendar (Posteingang des Kalenders) muss vorhanden sein
        Assert.Contains(result, c => c.IsDefaultCalendar);
    }

    [SkippableFact]
    public async Task ListEvents_Next7Days_DoesNotThrow()
    {
        SkipIfOutlookNotAvailable();

        // Fenster: jetzt bis jetzt+7 Tage
        var start = DateTimeOffset.Now;
        var end = start.AddDays(7);

        // Act
        var result = await Adapter.ListEventsAsync(
            calendarId: null,
            start: new DateTimeTimeZone
            {
                DateTime = start.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            end: new DateTimeTimeZone
            {
                DateTime = end.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            top: 50,
            cancellationToken: CancellationToken.None);

        // Assert: kein Throw, Value ist nicht null
        Assert.NotNull(result);
        Assert.NotNull(result.Value);
        // 0 Events ist ein valides Resultat (Outlook kann leer sein)
    }

    [SkippableFact]
    public async Task ListEvents_EmptyWindow_ReturnsEmpty()
    {
        SkipIfOutlookNotAvailable();

        // Fenster: 2000 (vor Outlook) bis 2001
        var start = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
        var end = new DateTimeOffset(2001, 1, 1, 0, 0, 0, TimeSpan.Zero);

        // Act
        var result = await Adapter.ListEventsAsync(
            calendarId: null,
            start: new DateTimeTimeZone
            {
                DateTime = start.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            end: new DateTimeTimeZone
            {
                DateTime = end.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            top: 50,
            cancellationToken: CancellationToken.None);

        // Assert: leeres Resultat, kein Fehler
        Assert.NotNull(result);
        Assert.Empty(result.Value);
        Assert.Null(result.NextSkip);
    }
}