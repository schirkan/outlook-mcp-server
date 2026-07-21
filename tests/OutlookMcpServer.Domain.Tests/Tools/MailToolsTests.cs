using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;
using OutlookMcpServer.Domain.Tests.Fakes;
using OutlookMcpServer.Tools;
using Xunit;

namespace OutlookMcpServer.Domain.Tests.Tools;

/// <summary>
/// Tests fuer <see cref="MailTools"/>: prueft, dass die Tool-Methoden
/// korrekt an <see cref="OutlookMcpServer.Domain.Abstractions.IOutlookService"/>
/// delegieren und die richtigen Datentypen mappen (CSV-Recipients,
/// Body-ContentType, Categories, Permanent-Flag).
/// </summary>
public sealed class MailToolsTests
{
    private static (MailTools, FakeOutlookService) Create()
    {
        var svc = new FakeOutlookService();
        var tools = new MailTools(svc, NullLogger<MailTools>.Instance);
        return (tools, svc);
    }

    [Fact]
    public async Task ListMailFolders_FiltersHidden()
    {
        var (tools, svc) = Create();
        svc.SeedFolder(SampleData.Inbox());
        svc.SeedFolder(SampleData.HiddenFolder());

        var pageHidden = await tools.ListMailFolders(includeHidden: true);
        var pageVisible = await tools.ListMailFolders(includeHidden: false);

        Assert.Equal(2, pageHidden.Value.Count);
        Assert.Single(pageVisible.Value);
        Assert.Equal(SampleData.InboxEntryId, pageVisible.Value[0].Id);
    }

    [Fact]
    public async Task SendMail_ParsesCsvRecipients()
    {
        var (tools, svc) = Create();
        svc.OnSendMail = req =>
        {
            Assert.Equal(new[] { "a@x.com", "b@x.com", "c@x.com" }, req.To);
            Assert.Equal(new[] { "cc@x.com" }, req.Cc);
            Assert.Equal(new[] { "bcc@x.com" }, req.Bcc);
            Assert.Equal(Importance.High, req.Importance);
            Assert.True(req.SaveToSentItems);
            return new SendMailResult { Sent = true, Id = "X" };
        };

        var result = await tools.SendMail(
            subject: "Hi",
            body: "Hello",
            bodyContentType: "html",
            to: "a@x.com, b@x.com ,c@x.com",
            cc: "cc@x.com",
            bcc: "bcc@x.com",
            importance: "high");

        Assert.True(result.Sent);
        Assert.Contains(svc.Calls, c => c.StartsWith(nameof(FakeOutlookService.SendMailAsync)));
    }

    [Fact]
    public async Task SendMail_DefaultsToTextContentType()
    {
        var (tools, svc) = Create();
        ItemBody? capturedBody = null;
        svc.OnSendMail = req =>
        {
            capturedBody = req.Body;
            return new SendMailResult { Sent = true };
        };

        await tools.SendMail(subject: "Hi", body: "Hello", to: "a@x.com");

        Assert.NotNull(capturedBody);
        Assert.Equal(ItemBodyType.Text, capturedBody!.ContentType);
        Assert.Equal("Hello", capturedBody.Content);
    }

    [Fact]
    public async Task SendMail_BodyContentTypeHtml_MapsToEnum()
    {
        var (tools, svc) = Create();
        ItemBody? capturedBody = null;
        svc.OnSendMail = req =>
        {
            capturedBody = req.Body;
            return new SendMailResult { Sent = true };
        };

        await tools.SendMail(subject: "Hi", body: "<p>Hi</p>", bodyContentType: "HTML", to: "a@x.com");
        Assert.Equal(ItemBodyType.Html, capturedBody!.ContentType);
    }

    [Fact]
    public async Task SendMail_ImportanceDefaultsToNormal_WhenInvalidString()
    {
        var (tools, svc) = Create();
        Importance? captured = null;
        svc.OnSendMail = req =>
        {
            captured = req.Importance;
            return new SendMailResult { Sent = true };
        };

        await tools.SendMail(subject: "x", body: "y", to: "a@x.com", importance: "garbage");
        Assert.Equal(Importance.Normal, captured);
    }

    [Fact]
    public async Task UpdateMail_ParsesCategories()
    {
        var (tools, svc) = Create();
        await tools.UpdateMail(
            id: "M-1",
            isRead: true,
            categories: "Wichtig, Projekt Alpha , RoadMap",
            importance: "low");

        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.UpdateMailAsync)));
        Assert.Contains("Wichtig", entry);
        Assert.Contains("Projekt Alpha", entry);
    }

    [Fact]
    public async Task DeleteMail_ForwardsPermanent()
    {
        var (tools, svc) = Create();
        await tools.DeleteMail(id: "M-1", permanent: true);
        var entry = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.DeleteMailAsync)));
        Assert.Contains("permanent=True", entry);
    }

    [Fact]
    public async Task MoveMail_ForwardsDestination()
    {
        var (tools, svc) = Create();
        var result = await tools.MoveMail(id: "M-1", destinationFolderId: "inbox");
        Assert.StartsWith("MOVED-", result);
    }

    [Fact]
    public async Task ListMails_PaginationReturnsCorrectSlice()
    {
        var (tools, svc) = Create();
        svc.SeedMail(SampleData.Mail1());
        svc.SeedMail(SampleData.Mail2());
        svc.SeedMail(SampleData.Mail3());

        var first = await tools.ListMails(folderId: "inbox", top: 2, skip: 0);
        var second = await tools.ListMails(folderId: "inbox", top: 2, skip: 2);

        Assert.Equal(2, first.Value.Count);
        Assert.Equal(2, first.NextSkip);
        Assert.Single(second.Value);
        Assert.Null(second.NextSkip);
    }

    [Fact]
    public async Task CreateDraft_ForwardsRequest()
    {
        var (tools, svc) = Create();
        var id = await tools.CreateDraft(subject: "D", body: "x", to: "a@x.com", importance: "low");
        Assert.Equal("DRAFT-1", id);
        Assert.Contains(svc.Calls, c => c.StartsWith(nameof(FakeOutlookService.CreateDraftAsync)));
    }

    // ===== Bulk-Get: get_mails =====

    [Fact]
    public async Task GetMails_DelegatesToService_WithIds()
    {
        var (tools, svc) = Create();
        svc.SeedBulkMails(new[] { SampleData.Mail1(), SampleData.Mail2() });

        var result = await tools.GetMails(
            ids: new[] { "MAIL-1", "MAIL-2", "MAIL-MISSING" },
            includeBody: true);

        Assert.Equal(2, result.Value.Count);
        Assert.Single(result.NotFoundIds);
        Assert.Equal("MAIL-MISSING", result.NotFoundIds[0]);
        Assert.Contains(svc.Calls, c => c.StartsWith(nameof(FakeOutlookService.GetMailsAsync)));
    }

    [Fact]
    public async Task GetMails_DefaultIncludeBody_False()
    {
        var (tools, svc) = Create();
        svc.SeedBulkMails(new[] { SampleData.Mail1() });

        var result = await tools.GetMails(ids: new[] { "MAIL-1" });

        Assert.Single(result.Value);
        var call = svc.Calls.Single(c => c.StartsWith(nameof(FakeOutlookService.GetMailsAsync)));
        Assert.Contains("includeBody=False", call);
    }

    [Fact]
    public async Task GetMails_AllIdsMissing_ReturnsEmptyValueWithNotFound()
    {
        var (tools, svc) = Create();
        svc.SeedBulkMails(Array.Empty<MailMessage>());

        var result = await tools.GetMails(ids: new[] { "X-1", "X-2" });

        Assert.Empty(result.Value);
        Assert.Equal(2, result.NotFoundIds.Count);
        Assert.Contains("X-1", result.NotFoundIds);
        Assert.Contains("X-2", result.NotFoundIds);
    }
}
