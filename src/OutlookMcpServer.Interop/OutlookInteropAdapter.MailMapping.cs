using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    private record MailRecipients(IReadOnlyList<Recipient> To, IReadOnlyList<Recipient> Cc, IReadOnlyList<Recipient> Bcc);
    private record MailDates(DateTimeOffset? Sent, DateTimeOffset? Received);
    private record MailMetadata(bool HasAttachments, bool IsUnread, Importance Importance, List<string> Categories);

    private const int OlRecipientTo = 2;
    private const int OlRecipientCC = 3;
    private const int OlRecipientBCC = 4;

    /// <summary>
    /// Mappt einen Outlook-Mailordner auf das Domain-DTO.
    /// </summary>
    /// <param name="folder">Outlook-MAPIFolder-COM-Referenz.</param>
    /// <param name="parentId">Bekannte Parent-EntryID oder null (dann wird folder.Parent ausgelesen).</param>
    /// <param name="wellKnownName">
    /// Optionaler Well-Known-Name, falls der Folder einem der Standard-Ordner
    /// (Inbox/Drafts/SentItems/...) entspricht. Wird vom Caller gesetzt, wenn
    /// die EntryID mit einem GetDefaultFolder()-Treffer uebereinstimmt.
    /// </param>
    private static MailFolder MapMailFolder(dynamic folder, string? parentId, string? wellKnownName = null)
    {
        var entryId = (string)folder.EntryID;
        var name = (string)folder.Name;
        int childCount = 0;
        int totalCount = 0;
        int unreadCount = 0;
        try { childCount = (int)folder.Folders.Count; } catch { }
        try { totalCount = (int)folder.Items.Count; } catch { }
        try { unreadCount = (int)folder.UnReadItemCount; } catch { }

        string? actualParentId = parentId;
        if (actualParentId is null)
        {
            try
            {
                var parent = folder.Parent;
                if (parent is not null)
                {
                    actualParentId = (string)parent.EntryID;
                    ReleaseComObject(parent);
                }
            }
            catch { }
        }

        return new MailFolder
        {
            Id = entryId,
            DisplayName = name,
            WellKnownName = wellKnownName,
            ParentFolderId = actualParentId,
            ChildFolderCount = childCount,
            TotalItemCount = totalCount,
            UnreadItemCount = unreadCount,
        };
    }

    /// <summary>
    /// Mappt ein Outlook-MailItem inkl. optionalem Body auf das Domain-DTO.
    /// </summary>
    private static MailMessage MapMailItem(dynamic mail, BodyFormat? bodyFormat, ILogger? logger = null)
    {
        var entryId = (string)mail.EntryID;

        var body = TryMapMailBody(mail, bodyFormat, logger);
        var from = TryMapMailFrom(mail);
        var recipients = TryMapMailRecipients(mail);
        var dates = TryMapMailDates(mail);
        var meta = TryMapMailMetadata(mail);

        return new MailMessage
        {
            Id = entryId,
            ConversationId = TryGetString(mail, "ConversationID"),
            Subject = TryGetString(mail, "Subject"),
            BodyPreview = TryGetString(mail, "BodyPreview"),
            Body = body,
            From = from,
            ToRecipients = recipients.To,
            CcRecipients = recipients.Cc,
            BccRecipients = recipients.Bcc,
            SentDateTime = dates.Sent,
            ReceivedDateTime = dates.Received,
            HasAttachments = meta.HasAttachments,
            Importance = meta.Importance,
            IsRead = !meta.IsUnread,
            Categories = meta.Categories,
        };
    }

    /// <summary>
    /// Liest und konvertiert den Mail-Body in das angeforderte Zielformat.
    /// </summary>
    private static ItemBody? TryMapMailBody(dynamic mail, BodyFormat? bodyFormat, ILogger? logger = null)
    {
        if (bodyFormat is null) return null;
        try
        {
            string text = string.Empty;
            string html = string.Empty;
            try { text = (string)mail.Body ?? string.Empty; }
            catch (COMException bodyEx)
            {
                logger?.LogWarning(bodyEx,
                    "TryMapMailBody: Mail.Body nicht lesbar (HResult=0x{HResult:X8})",
                    bodyEx.HResult);
            }
            catch (Exception bodyEx)
            {
                logger?.LogWarning(bodyEx, "TryMapMailBody: Unerwarteter Fehler beim Lesen von Mail.Body");
            }
            try { html = (string)mail.HTMLBody ?? string.Empty; }
            catch (COMException htmlEx)
            {
                logger?.LogWarning(htmlEx,
                    "TryMapMailBody: Mail.HTMLBody nicht lesbar (HResult=0x{HResult:X8})",
                    htmlEx.HResult);
            }
            catch (Exception htmlEx)
            {
                logger?.LogWarning(htmlEx, "TryMapMailBody: Unerwarteter Fehler beim Lesen von Mail.HTMLBody");
            }

            string? nativeSource = null;
            if (!string.IsNullOrEmpty(html))
            {
                nativeSource = html;
            }
            else if (!string.IsNullOrEmpty(text))
            {
                nativeSource = text;
            }

            if (nativeSource is null || nativeSource.Length == 0)
            {
                try
                {
                    var preview = (string?)mail.BodyPreview;
                    if (!string.IsNullOrEmpty(preview))
                    {
                        logger?.LogInformation(
                            "TryMapMailBody: Body leer, liefere BodyPreview als Fallback (len={Length})",
                            preview.Length);
                        return new ItemBody { ContentType = ItemBodyType.Text, Content = preview };
                    }
                }
                catch (Exception previewEx)
                {
                    logger?.LogWarning(previewEx, "TryMapMailBody: BodyPreview ebenfalls nicht lesbar");
                }
                logger?.LogInformation("TryMapMailBody: Mail hat weder Body, HTMLBody noch BodyPreview");
                return null;
            }

            string htmlForConversion = nativeSource == text
                ? $"<html><body><pre>{System.Net.WebUtility.HtmlEncode(text)}</pre></body></html>"
                : nativeSource;

            var (ct, content) = HtmlBodyConverter.Convert(htmlForConversion, bodyFormat.Value);
            return new ItemBody { ContentType = ct, Content = content };
        }
        catch (Exception ex)
        {
            logger?.LogWarning(ex, "TryMapMailBody: Outer try fehlgeschlagen");
        }
        return null;
    }

    /// <summary>
    /// Extrahiert den Absender einer Mail, sofern Outlook ihn lesbar liefert.
    /// </summary>
    private static Recipient? TryMapMailFrom(dynamic mail)
    {
        try
        {
            var senderName = TryGetString(mail, "SenderName");
            var senderAddress = TryGetString(mail, "SenderEmailAddress");
            if (mail.Sender is not null && !string.IsNullOrEmpty(senderAddress))
            {
                return new Recipient
                {
                    EmailAddress = new EmailAddress { Name = senderName, Address = senderAddress },
                };
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Liest Outlook-Recipients und gruppiert sie nach To, Cc und Bcc.
    /// </summary>
    private static MailRecipients TryMapMailRecipients(dynamic mail)
    {
        dynamic? recipients = TryGetDynamic(mail, "Recipients");
        if (recipients is null)
        {
            return new MailRecipients(
                Array.Empty<Recipient>(),
                Array.Empty<Recipient>(),
                Array.Empty<Recipient>());
        }
        var to = new List<Recipient>();
        var cc = new List<Recipient>();
        var bcc = new List<Recipient>();
        try
        {
            foreach (var rec in recipients)
            {
                try
                {
                    if (rec is null) continue;
                    int type = 0;
                    try { type = (int)rec.Type; } catch { }
                    var name = TryGetString(rec, "Name");
                    var address = TryGetString(rec, "Address");
                    if (string.IsNullOrEmpty(address)) continue;
                    var recipient = new Recipient
                    {
                        EmailAddress = new EmailAddress { Name = name, Address = address },
                    };
                    switch (type)
                    {
                        case OlRecipientTo: to.Add(recipient); break;
                        case OlRecipientCC: cc.Add(recipient); break;
                        case OlRecipientBCC: bcc.Add(recipient); break;
                    }
                }
                finally
                {
                    ReleaseComObject(rec);
                }
            }
        }
        catch { }
        finally
        {
            ReleaseComObject(recipients);
        }
        return new MailRecipients(to, cc, bcc);
    }

    /// <summary>
    /// Liest Versand- und Empfangszeit einer Mail defensiv aus.
    /// </summary>
    private static MailDates TryMapMailDates(dynamic mail)
    {
        DateTimeOffset? sent = null;
        DateTimeOffset? received = null;
        try { sent = (DateTimeOffset)mail.SentOn; } catch { }
        try { received = (DateTimeOffset)mail.ReceivedTime; } catch { }
        return new MailDates(sent, received);
    }

    /// <summary>
    /// Liest Metadaten wie Attachments, Read-Status, Importance und Kategorien.
    /// </summary>
    private static MailMetadata TryMapMailMetadata(dynamic mail)
    {
        bool hasAttachments = false;
        try { hasAttachments = (int)mail.Attachments.Count > 0; } catch { }

        bool isUnread = false;
        try { isUnread = (bool)mail.UnRead; } catch { }

        var importanceRaw = 1;
        try { importanceRaw = (int)mail.Importance; } catch { }
        var importance = OlEnumMappings.ToImportance(importanceRaw);

        var categories = new List<string>();
        try
        {
            foreach (var cat in mail.Categories)
            {
                var s = (string)cat;
                if (!string.IsNullOrEmpty(s)) categories.Add(s);
            }
        }
        catch { }

        return new MailMetadata(hasAttachments, isUnread, importance, categories);
    }

    /// <summary>
    /// Mappt eine Outlook-Recipients-Collection auf Domain-Recipients.
    /// </summary>
    private static IReadOnlyList<Recipient> MapRecipients(dynamic recipients)
    {
        var list = new List<Recipient>();
        if (recipients is null) return list;
        try
        {
            foreach (var rec in recipients)
            {
                try
                {
                    if (rec is null) continue;
                    var name = TryGetString(rec, "Name");
                    var address = TryGetString(rec, "Address");
                    if (string.IsNullOrEmpty(address)) continue;
                    list.Add(new Recipient
                    {
                        EmailAddress = new EmailAddress { Name = name, Address = address },
                    });
                }
                finally
                {
                    ReleaseComObject(rec);
                }
            }
        }
        catch { }
        return list;
    }
}
