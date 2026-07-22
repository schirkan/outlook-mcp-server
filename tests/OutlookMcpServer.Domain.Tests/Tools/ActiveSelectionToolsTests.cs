using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;
using OutlookMcpServer.Tools;
using OutlookMcpServer.Domain.Tests.Fakes;
using Xunit;

namespace OutlookMcpServer.Domain.Tests.Tools;

/// <summary>
/// Tests fuer <see cref="ActiveSelectionTools"/> (Phase 3h Acceptance Criteria):
/// Polymorphie (kind=mail vs kind=event), Scope-Filter (mail/calendar/any),
/// Top-Cap, Selection.Count==0 (= Empty-Result ohne Fehler), ParseSelectionScope.
/// </summary>
public sealed class ActiveSelectionToolsTests
{
    private static ActiveSelectionTools CreateTools(FakeOutlookService fake)
        => new(fake, NullLogger<ActiveSelectionTools>.Instance);

    private static MailMessage SampleMail(string id) => new()
    {
        Id = id,
        Subject = $"Subject-{id}",
        From = new Recipient { EmailAddress = new EmailAddress { Address = "alice@example.com" } },
    };

    private static CalendarEvent SampleEvent(string id) => new()
    {
        Id = id,
        Subject = $"Subject-{id}",
        Start = new DateTimeTimeZone { DateTime = "2026-07-16T09:00:00", TimeZone = "Europe/Berlin" },
        End = new DateTimeTimeZone { DateTime = "2026-07-16T10:00:00", TimeZone = "Europe/Berlin" },
    };

    // ===== GetActiveItem =====

    [Fact]
    public async Task GetActiveItem_NullWhenNoInspector()
    {
        // Default-Fake gibt null zurueck
        var fake = new FakeOutlookService();
        var tools = CreateTools(fake);

        var result = await tools.GetActiveItem();

        Assert.Null(result);
        Assert.Contains($"{nameof(FakeOutlookService.GetActiveItemAsync)}:bodyFormat=Markdown", fake.Calls);
    }

    [Fact]
    public async Task GetActiveItem_ReturnsActiveMail()
    {
        var mail = SampleMail("M-1");
        var fake = new FakeOutlookService
        {
            OnGetActiveItem = (_, _) => new ActiveMail { Item = mail },
        };
        var tools = CreateTools(fake);

        var result = await tools.GetActiveItem();

        Assert.NotNull(result);
        var activeMail = Assert.IsType<ActiveMail>(result);
        Assert.Equal("mail", activeMail.Kind);
        Assert.Equal("M-1", activeMail.Item.Id);
    }

    [Fact]
    public async Task GetActiveItem_ReturnsActiveEvent()
    {
        var evt = SampleEvent("E-1");
        var fake = new FakeOutlookService
        {
            OnGetActiveItem = (_, _) => new ActiveEvent { Item = evt },
        };
        var tools = CreateTools(fake);

        var result = await tools.GetActiveItem();

        Assert.NotNull(result);
        var activeEvent = Assert.IsType<ActiveEvent>(result);
        Assert.Equal("event", activeEvent.Kind);
        Assert.Equal("E-1", activeEvent.Item.Id);
    }

    // ===== GetSelectedItems - Empty =====

    [Fact]
    public async Task GetSelectedItems_EmptyByDefault()
    {
        var fake = new FakeOutlookService();
        var tools = CreateTools(fake);

        var result = await tools.GetSelectedItems();

        Assert.NotNull(result);
        Assert.Empty(result);
        Assert.Contains("scope=Any,top=50", string.Join(",", fake.Calls));
    }

    // ===== GetSelectedItems - Scope-Filter =====

    [Fact]
    public async Task GetSelectedItems_ScopeAny_ReturnsAllKinds()
    {
        var fake = new FakeOutlookService
        {
            OnGetSelectedItems = (_, _, _, _) => new ActiveItem[]
            {
                new ActiveMail { Item = SampleMail("M-1") },
                new ActiveEvent { Item = SampleEvent("E-1") },
                new ActiveMail { Item = SampleMail("M-2") },
            },
        };
        var tools = CreateTools(fake);

        var result = await tools.GetSelectedItems(scope: "any", top: 50);

        Assert.Equal(3, result.Count);
        Assert.Equal(2, result.OfType<ActiveMail>().Count());
        Assert.Single(result.OfType<ActiveEvent>());
    }

    [Fact]
    public async Task GetSelectedItems_ScopeMail_OnlyMailItems()
    {
        var fake = new FakeOutlookService
        {
            OnGetSelectedItems = (_, _, _, _) => new ActiveItem[]
            {
                new ActiveMail { Item = SampleMail("M-1") },
                new ActiveEvent { Item = SampleEvent("E-1") },
                new ActiveMail { Item = SampleMail("M-2") },
            },
        };
        var tools = CreateTools(fake);

        var result = await tools.GetSelectedItems(scope: "mail", top: 50);

        Assert.Equal(2, result.Count);
        Assert.All(result, item => Assert.IsType<ActiveMail>(item));
    }

    [Fact]
    public async Task GetSelectedItems_ScopeCalendar_OnlyEventItems()
    {
        var fake = new FakeOutlookService
        {
            OnGetSelectedItems = (_, _, _, _) => new ActiveItem[]
            {
                new ActiveMail { Item = SampleMail("M-1") },
                new ActiveEvent { Item = SampleEvent("E-1") },
                new ActiveEvent { Item = SampleEvent("E-2") },
            },
        };
        var tools = CreateTools(fake);

        var result = await tools.GetSelectedItems(scope: "calendar", top: 50);

        Assert.Equal(2, result.Count);
        Assert.All(result, item => Assert.IsType<ActiveEvent>(item));
    }

    [Fact]
    public async Task GetSelectedItems_ScopeDefaultsToAny_WhenInvalidOrEmpty()
    {
        var fake = new FakeOutlookService
        {
            OnGetSelectedItems = (_, _, _, _) => new ActiveItem[]
            {
                new ActiveMail { Item = SampleMail("M-1") },
                new ActiveEvent { Item = SampleEvent("E-1") },
            },
        };
        var tools = CreateTools(fake);

        // Leere / unbekannte Werte -> Fallback any
        var result = await tools.GetSelectedItems(scope: "unknown-garbage", top: 50);

        Assert.Equal(2, result.Count);
        // Service wurde mit scope=Any aufgerufen
        Assert.Contains("scope=Any,top=50", string.Join(",", fake.Calls));
    }

    // ===== GetSelectedItems - Top-Cap =====

    [Fact]
    public async Task GetSelectedItems_TopCap_LimitsResults()
    {
        var fake = new FakeOutlookService
        {
            OnGetSelectedItems = (_, _, _, _) => new ActiveItem[]
            {
                new ActiveMail { Item = SampleMail("M-1") },
                new ActiveMail { Item = SampleMail("M-2") },
                new ActiveMail { Item = SampleMail("M-3") },
                new ActiveMail { Item = SampleMail("M-4") },
                new ActiveMail { Item = SampleMail("M-5") },
            },
        };
        var tools = CreateTools(fake);

        var result = await tools.GetSelectedItems(scope: "any", top: 2);

        Assert.Equal(2, result.Count);
        // Service wurde mit top=2 aufgerufen
        Assert.Contains("scope=Any,top=2", string.Join(",", fake.Calls));
    }

    // ===== GetSelectedItems - Validation (via OutlookService) =====

    [Fact]
    public async Task GetSelectedItems_ZeroTop_ThrowsInvalidInput()
    {
        // OutlookService validiert top in [1,250] (siehe OutlookService.GetSelectedItemsAsync).
        // ActiveSelectionTools leitet 1:1 an den Service -> Validation-Fehler propagiert.
        var fake = new FakeOutlookService();
        var tools = CreateTools(fake);

        var ex = await Assert.ThrowsAsync<OutlookServiceException>(
            () => tools.GetSelectedItems(scope: "any", top: 0));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task GetSelectedItems_TopAbove250_ThrowsInvalidInput()
    {
        var fake = new FakeOutlookService();
        var tools = CreateTools(fake);

        var ex = await Assert.ThrowsAsync<OutlookServiceException>(
            () => tools.GetSelectedItems(scope: "any", top: 251));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    // ===== GetSelectedItems - Service-Failure-Propagation =====

    [Fact]
    public async Task GetSelectedItems_OutlookNotActive_PropagatesException()
    {
        // Wenn der Adapter (Service) OutlookNotActive wirft, muss der Tool-Aufruf
        // die Exception unveraendert durchreichen (kein Swallowing).
        var fake = new FakeOutlookService
        {
            OnGetSelectedItems = (_, _, _, _) => throw new OutlookServiceException(
                ErrorCode.OutlookNotActive,
                "Outlook ActiveExplorer ist null"),
        };
        var tools = CreateTools(fake);

        var ex = await Assert.ThrowsAsync<OutlookServiceException>(
            () => tools.GetSelectedItems(scope: "any", top: 50));
        Assert.Equal(ErrorCode.OutlookNotActive, ex.Code);
    }
}