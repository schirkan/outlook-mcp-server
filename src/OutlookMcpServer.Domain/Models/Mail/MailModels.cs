using System.Text.Json.Serialization;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Domain.Models.Mail;

public enum Importance
{
    [JsonPropertyName("low")]
    Low,

    [JsonPropertyName("normal")]
    Normal,

    [JsonPropertyName("high")]
    High,
}

public enum Sensitivity
{
    [JsonPropertyName("normal")]
    Normal,

    [JsonPropertyName("personal")]
    Personal,

    [JsonPropertyName("private")]
    Private,

    [JsonPropertyName("confidential")]
    Confidential,
}

/// <summary>
/// Well-known folder names laut Microsoft Graph. Wird beim Lookup
/// <c>folderId</c> in MailTools akzeptiert (statt EntryID).
/// </summary>
public static class WellKnownFolder
{
    public const string Inbox = "inbox";
    public const string Drafts = "drafts";
    public const string SentItems = "sentItems";
    public const string DeletedItems = "deletedItems";
    public const string JunkEmail = "junkEmail";
    public const string Archive = "archive";
    public const string Outbox = "outbox";
}

/// <summary>
/// Mail-Ordner. <c>WellKnownName</c> ist null fuer Custom-Ordner.
/// </summary>
public sealed record MailFolder
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; init; } = string.Empty;

    [JsonPropertyName("wellKnownName")]
    public string? WellKnownName { get; init; }

    [JsonPropertyName("parentFolderId")]
    public string? ParentFolderId { get; init; }

    [JsonPropertyName("childFolderCount")]
    public int ChildFolderCount { get; init; }

    [JsonPropertyName("totalItemCount")]
    public int TotalItemCount { get; init; }

    [JsonPropertyName("unreadItemCount")]
    public int UnreadItemCount { get; init; }
}

/// <summary>
/// Empfaenger (To/Cc/Bcc). 1:1 zu Microsoft Graph <c>Recipient</c>.
/// </summary>
public sealed record Recipient
{
    [JsonPropertyName("emailAddress")]
    public EmailAddress EmailAddress { get; init; } = new();
}

/// <summary>
/// Vollstaendige Mail-Nachricht. Properties 1:1 zu Microsoft Graph <c>message</c>.
/// </summary>
public sealed record MailMessage
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("conversationId")]
    public string? ConversationId { get; init; }

    [JsonPropertyName("subject")]
    public string? Subject { get; init; }

    [JsonPropertyName("bodyPreview")]
    public string? BodyPreview { get; init; }

    [JsonPropertyName("body")]
    public ItemBody? Body { get; init; }

    [JsonPropertyName("from")]
    public Recipient? From { get; init; }

    [JsonPropertyName("toRecipients")]
    public IReadOnlyList<Recipient> ToRecipients { get; init; } = Array.Empty<Recipient>();

    [JsonPropertyName("ccRecipients")]
    public IReadOnlyList<Recipient> CcRecipients { get; init; } = Array.Empty<Recipient>();

    [JsonPropertyName("bccRecipients")]
    public IReadOnlyList<Recipient> BccRecipients { get; init; } = Array.Empty<Recipient>();

    [JsonPropertyName("sentDateTime")]
    public DateTimeOffset? SentDateTime { get; init; }

    [JsonPropertyName("receivedDateTime")]
    public DateTimeOffset? ReceivedDateTime { get; init; }

    [JsonPropertyName("hasAttachments")]
    public bool HasAttachments { get; init; }

    [JsonPropertyName("importance")]
    public Importance Importance { get; init; } = Importance.Normal;

    [JsonPropertyName("isRead")]
    public bool IsRead { get; init; }

    [JsonPropertyName("categories")]
    public IReadOnlyList<string> Categories { get; init; } = Array.Empty<string>();

    [JsonPropertyName("internetMessageHeaders")]
    public IReadOnlyList<InternetMessageHeader>? InternetMessageHeaders { get; init; }
}

/// <summary>
/// Zusammenfassung einer Anlage (ohne Inhalt).
/// </summary>
public sealed record AttachmentSummary
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    [JsonPropertyName("contentType")]
    public string ContentType { get; init; } = string.Empty;

    [JsonPropertyName("size")]
    public long Size { get; init; }

    [JsonPropertyName("isInline")]
    public bool IsInline { get; init; }
}

/// <summary>
/// Vollstaendige Anlage mit Base64-codiertem Inhalt.
/// </summary>
public sealed record AttachmentData
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    [JsonPropertyName("contentType")]
    public string ContentType { get; init; } = string.Empty;

    [JsonPropertyName("size")]
    public long Size { get; init; }

    [JsonPropertyName("contentBase64")]
    public string ContentBase64 { get; init; } = string.Empty;
}

/// <summary>
/// Inline-Anlage fuer <c>sendMail</c>/<c>createDraft</c>.
/// </summary>
public sealed record InlineAttachment
{
    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    [JsonPropertyName("contentType")]
    public string ContentType { get; init; } = string.Empty;

    [JsonPropertyName("contentBase64")]
    public string ContentBase64 { get; init; } = string.Empty;
}

/// <summary>
/// Input fuer <c>sendMail</c> und <c>createDraft</c>.
/// </summary>
public sealed record SendMailRequest
{
    [JsonPropertyName("to")]
    public IReadOnlyList<string> To { get; init; } = Array.Empty<string>();

    [JsonPropertyName("cc")]
    public IReadOnlyList<string> Cc { get; init; } = Array.Empty<string>();

    [JsonPropertyName("bcc")]
    public IReadOnlyList<string> Bcc { get; init; } = Array.Empty<string>();

    [JsonPropertyName("subject")]
    public string Subject { get; init; } = string.Empty;

    [JsonPropertyName("body")]
    public ItemBody Body { get; init; } = new();

    [JsonPropertyName("importance")]
    public Importance Importance { get; init; } = Importance.Normal;

    [JsonPropertyName("attachments")]
    public IReadOnlyList<InlineAttachment> Attachments { get; init; } = Array.Empty<InlineAttachment>();

    [JsonPropertyName("replyToId")]
    public string? ReplyToId { get; init; }

    [JsonPropertyName("replyAll")]
    public bool ReplyAll { get; init; }

    [JsonPropertyName("forwardFromId")]
    public string? ForwardFromId { get; init; }

    [JsonPropertyName("sendAt")]
    public DateTimeOffset? SendAt { get; init; }

    [JsonPropertyName("saveToSentItems")]
    public bool SaveToSentItems { get; init; } = true;
}

/// <summary>
/// Input fuer <c>updateMail</c>. Alle Felder optional (PATCH-Semantik).
/// </summary>
public sealed record MailUpdate
{
    [JsonPropertyName("isRead")]
    public bool? IsRead { get; init; }

    [JsonPropertyName("categories")]
    public IReadOnlyList<string>? Categories { get; init; }

    [JsonPropertyName("importance")]
    public Importance? Importance { get; init; }
}

/// <summary>
/// Ergebnis von <c>sendMail</c>: Sent-Mail-EntryID in SentItems.
/// </summary>
public sealed record SendMailResult
{
    [JsonPropertyName("sent")]
    public bool Sent { get; init; }

    [JsonPropertyName("id")]
    public string? Id { get; init; }
}