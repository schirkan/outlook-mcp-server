using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Domain.Tests.Fakes;

/// <summary>
/// Statische Builder fuer deterministische Test-Daten.
/// Jeder Builder hat einen optionalen <c>id</c>-Override; default-Werte
/// sind entryIDs wie aus Outlook (echte Hex-Strings).
/// </summary>
public static class SampleData
{
    public const string InboxEntryId = "00000000-0000-0000-0000-000000000001";
    public const string DraftsEntryId = "00000000-0000-0000-0000-000000000002";
    public const string SentEntryId = "00000000-0000-0000-0000-000000000003";

    public static MailFolder Inbox(string? id = null) => new()
    {
        Id = id ?? InboxEntryId,
        DisplayName = "Posteingang",
        WellKnownName = WellKnownFolder.Inbox,
        TotalItemCount = 10,
        UnreadItemCount = 3,
    };

    public static MailFolder HiddenFolder() => new()
    {
        Id = "00000000-0000-0000-0000-000000000099",
        DisplayName = "Hidden Sync Issues",
        TotalItemCount = 0,
    };

    public static MailMessage Mail1() => new()
    {
        Id = "MAIL-1",
        Subject = "Projektstatus Q2",
        BodyPreview = "Kurze Vorschau...",
        From = new Recipient { EmailAddress = new EmailAddress { Name = "Alice", Address = "alice@example.com" } },
        ToRecipients = new[] { new Recipient { EmailAddress = new EmailAddress { Address = "bob@example.com" } } },
        SentDateTime = DateTimeOffset.Parse("2026-07-01T10:00:00+02:00"),
        ReceivedDateTime = DateTimeOffset.Parse("2026-07-01T10:00:05+02:00"),
        Importance = Importance.High,
        IsRead = false,
        HasAttachments = true,
    };

    public static MailMessage Mail2() => new()
    {
        Id = "MAIL-2",
        Subject = "Weekly Sync",
        From = new Recipient { EmailAddress = new EmailAddress { Address = "carol@example.com" } },
        SentDateTime = DateTimeOffset.Parse("2026-07-08T09:00:00+02:00"),
        ReceivedDateTime = DateTimeOffset.Parse("2026-07-08T09:00:03+02:00"),
        Importance = Importance.Normal,
        IsRead = true,
    };

    public static MailMessage Mail3() => new()
    {
        Id = "MAIL-3",
        Subject = "Random",
        From = new Recipient { EmailAddress = new EmailAddress { Address = "dan@example.com" } },
        SentDateTime = DateTimeOffset.Parse("2026-07-10T15:30:00+02:00"),
        ReceivedDateTime = DateTimeOffset.Parse("2026-07-10T15:30:01+02:00"),
        Importance = Importance.Low,
        IsRead = true,
    };

    public static Calendar DefaultCalendar() => new()
    {
        Id = "CAL-1",
        Name = "Kalender",
        IsDefaultCalendar = true,
        CanEdit = true,
        Owner = "owner@example.com",
    };

    public static CalendarEvent Event1() => new()
    {
        Id = "EVT-1",
        Subject = "Daily Standup",
        Start = new DateTimeTimeZone { DateTime = "2026-07-15T09:00:00", TimeZone = "Europe/Berlin" },
        End = new DateTimeTimeZone { DateTime = "2026-07-15T09:15:00", TimeZone = "Europe/Berlin" },
        Importance = Importance.Normal,
        Sensitivity = Sensitivity.Normal,
        ShowAs = ShowAs.Busy,
        IsAllDay = false,
    };

    public static SendMailRequest ValidSendRequest() => new()
    {
        To = new[] { "recipient@example.com" },
        Subject = "Hello",
        Body = new ItemBody { ContentType = ItemBodyType.Text, Content = "Body-Text" },
    };
}
