using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using OutlookMcpServer.Domain.Configuration;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;
using OutlookMcpServer.Domain.Services;
using OutlookMcpServer.Domain.Tests.Fakes;
using Xunit;

namespace OutlookMcpServer.Domain.Tests;

/// <summary>
/// Tests fuer <see cref="OutlookService"/>: Validation-Pfade (InvalidInput,
/// AttachmentTooLarge) und Policy-Pfade (AllowSend/AllowDelete/AllowCreate).
/// Adapter wird via <see cref="FakeInteropAdapter"/> aufgerufen.
/// </summary>
public sealed class OutlookServiceTests
{
    private static OutlookService CreateService(OutlookOptions options, FakeInteropAdapter adapter)
    {
        var opts = Options.Create(new OutlookMcpServerOptions { Outlook = options });
        return new OutlookService(adapter, opts, NullLogger<OutlookService>.Instance);
    }

    // ===== Mail-Validation =====

    [Fact]
    public async Task SendMailAsync_EmptyTo_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions { AllowSend = true }, new FakeInteropAdapter());
        var req = new SendMailRequest { To = Array.Empty<string>(), Subject = "X", Body = new ItemBody { Content = "Y" } };
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.SendMailAsync(req));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task SendMailAsync_InvalidEmailFormat_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions { AllowSend = true }, new FakeInteropAdapter());
        var req = new SendMailRequest { To = new[] { "not-an-email" }, Subject = "X", Body = new ItemBody { Content = "Y" } };
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.SendMailAsync(req));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task SendMailAsync_WithReplyToId_NoToRequired_DoesNotThrow()
    {
        var svc = CreateService(new OutlookOptions { AllowSend = true }, new FakeInteropAdapter());
        var req = new SendMailRequest { Subject = "Re: Hi", Body = new ItemBody { Content = "X" }, ReplyToId = "ORIG-1" };
        var result = await svc.SendMailAsync(req);
        Assert.True(result.Sent);
    }

    [Fact]
    public async Task SendMailAsync_EmptyBody_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions { AllowSend = true }, new FakeInteropAdapter());
        var req = new SendMailRequest { To = new[] { "a@b.com" }, Subject = "X", Body = new ItemBody { Content = "" } };
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.SendMailAsync(req));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task SendMailAsync_AttachmentTooLarge_ThrowsAttachmentTooLarge()
    {
        // Base64-Laenge ~ 8 Zeichen -> ~6 Bytes; MaxAttachmentBytes = 5 -> fail
        var svc = CreateService(new OutlookOptions { AllowSend = true, MaxAttachmentBytes = 5 }, new FakeInteropAdapter());
        var req = new SendMailRequest
        {
            To = new[] { "recipient@example.com" },
            Subject = "Hello",
            Body = new ItemBody { ContentType = ItemBodyType.Text, Content = "Body-Text" },
            Attachments = new[]
            {
                new InlineAttachment { Name = "big.bin", ContentType = "application/octet-stream", ContentBase64 = new string('A', 8) },
            },
        };
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.SendMailAsync(req));
        Assert.Equal(ErrorCode.AttachmentTooLarge, ex.Code);
    }

    [Fact]
    public async Task ListMailsAsync_TopAboveRange_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() =>
            svc.ListMailsAsync(SampleData.InboxEntryId, top: 200));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task GetMailAsync_EmptyId_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.GetMailAsync(""));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    // ===== Policy: AllowSend / AllowDelete / AllowCreate =====

    [Fact]
    public async Task SendMailAsync_AllowSendFalse_ThrowsSendDisabled()
    {
        var svc = CreateService(new OutlookOptions { AllowSend = false }, new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.SendMailAsync(SampleData.ValidSendRequest()));
        Assert.Equal(ErrorCode.SendDisabled, ex.Code);
    }

    [Fact]
    public async Task DeleteMailAsync_AllowDeleteFalse_ThrowsDeleteDisabled()
    {
        var svc = CreateService(new OutlookOptions { AllowDelete = false }, new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.DeleteMailAsync("X"));
        Assert.Equal(ErrorCode.DeleteDisabled, ex.Code);
    }

    [Fact]
    public async Task DeleteEventAsync_AllowDeleteFalse_ThrowsDeleteDisabled()
    {
        var svc = CreateService(new OutlookOptions { AllowDelete = false }, new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.DeleteEventAsync("X"));
        Assert.Equal(ErrorCode.DeleteDisabled, ex.Code);
    }

    [Fact]
    public async Task CreateDraftAsync_AllowCreateFalse_ThrowsDisabled()
    {
        var svc = CreateService(new OutlookOptions { AllowCreate = false }, new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.CreateDraftAsync(SampleData.ValidSendRequest()));
        // OutlookService nutzt aktuell SendDisabled-Code auch fuer CreateDisabled (Spec-Bug,
        // wird in Karte 8 (Doku) als Known-Issue dokumentiert).
        Assert.Equal(ErrorCode.SendDisabled, ex.Code);
    }

    [Fact]
    public async Task CreateEventAsync_AllowCreateFalse_ThrowsDisabled()
    {
        var svc = CreateService(new OutlookOptions { AllowCreate = false }, new FakeInteropAdapter());
        var req = new CreateEventRequest
        {
            Subject = "x",
            Start = new DateTimeTimeZone { DateTime = "2026-07-15T09:00:00", TimeZone = "Europe/Berlin" },
            End = new DateTimeTimeZone { DateTime = "2026-07-15T10:00:00", TimeZone = "Europe/Berlin" },
        };
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.CreateEventAsync(req));
        Assert.Equal(ErrorCode.SendDisabled, ex.Code);
    }

    // ===== Calendar-Validation =====

    [Fact]
    public async Task ListEventsAsync_InvalidTimeZone_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() =>
            svc.ListEventsAsync(
                calendarId: null,
                start: new DateTimeTimeZone { DateTime = "2026-07-15T09:00:00", TimeZone = "Atlantis/Lemuria" },
                end: new DateTimeTimeZone { DateTime = "2026-07-15T17:00:00", TimeZone = "Europe/Berlin" }));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task FindMeetingTimesAsync_DurationTooLarge_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var req = new FindMeetingTimesRequest
        {
            DurationMinutes = 25 * 60,
            TimeWindowStart = new DateTimeTimeZone { DateTime = "2026-07-15T00:00:00", TimeZone = "Europe/Berlin" },
            TimeWindowEnd = new DateTimeTimeZone { DateTime = "2026-07-16T00:00:00", TimeZone = "Europe/Berlin" },
        };
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.FindMeetingTimesAsync(req));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task RespondToEventAsync_ResponseNone_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var req = new RespondToEventRequest { Response = Response.None };
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() => svc.RespondToEventAsync("X", req));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    // ===== Happy-Path: Adapter wird gerufen =====

    [Fact]
    public async Task SendMailAsync_HappyPath_CallsAdapter()
    {
        var adapter = new FakeInteropAdapter();
        var svc = CreateService(new OutlookOptions { AllowSend = true }, adapter);
        await svc.SendMailAsync(SampleData.ValidSendRequest());
        Assert.Contains(adapter.Calls, c => c.StartsWith(nameof(OutlookMcpServer.Domain.Abstractions.IInteropOutlookAdapter.SendMailAsync)));
    }

    // ===== ListMailsRecursiveAsync: Validation + Passthrough =====

    [Fact]
    public async Task ListMailsRecursiveAsync_TopAboveRange_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() =>
            svc.ListMailsRecursiveAsync(new[] { "inbox" }, top: 200));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task ListMailsRecursiveAsync_InvalidScopeEntry_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() =>
            svc.ListMailsRecursiveAsync(new[] { "inbox", "not-a-real-folder" }));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
        Assert.Contains("not-a-real-folder", ex.Message);
    }

    [Fact]
    public async Task ListMailsRecursiveAsync_EmptyScopeEntry_ThrowsInvalidInput()
    {
        var svc = CreateService(new OutlookOptions(), new FakeInteropAdapter());
        var ex = await Assert.ThrowsAsync<OutlookServiceException>(() =>
            svc.ListMailsRecursiveAsync(new[] { "inbox", "" }));
        Assert.Equal(ErrorCode.InvalidInput, ex.Code);
    }

    [Fact]
    public async Task ListMailsRecursiveAsync_ValidScopeAndFilter_PassesToAdapter()
    {
        var adapter = new FakeInteropAdapter();
        var svc = CreateService(new OutlookOptions(), adapter);
        var result = await svc.ListMailsRecursiveAsync(
            new[] { "inbox", "archive" }, top: 10, filter: "[UnRead] = true");
        Assert.NotNull(result);
        Assert.Contains(
            adapter.Calls,
            c => c.Contains("scope=inbox,archive") && c.Contains("top=10") && c.Contains("[UnRead] = true"));
    }

    [Fact]
    public async Task ListMailsRecursiveAsync_NullScope_PassesEmptyArrayToAdapter()
    {
        var adapter = new FakeInteropAdapter();
        var svc = CreateService(new OutlookOptions(), adapter);
        await svc.ListMailsRecursiveAsync(null!, top: 5);
        Assert.Contains(adapter.Calls, c => c.Contains("ListMailsRecursiveAsync") && c.Contains("scope="));
    }
}
