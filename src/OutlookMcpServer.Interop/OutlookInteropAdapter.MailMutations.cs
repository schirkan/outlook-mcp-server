using System.Runtime.InteropServices;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    public async Task<SendMailResult> SendMailAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic mail = BuildMailItemFromRequest(request);
            try
            {
                mail.DeleteAfterSubmit = !request.SaveToSentItems;
                if (request.SendAt is { } sendAt)
                {
                    mail.DeferredDeliveryTime = sendAt.UtcDateTime;
                }
                mail.Send();
                return new SendMailResult { Sent = true, Id = null };
            }
            finally
            {
                ReleaseComObject(mail);
            }
        }, nameof(SendMailAsync));
    }

    public async Task<string> CreateDraftAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic mail = BuildMailItemFromRequest(request);
            try
            {
                mail.Save();
                return (string)mail.EntryID;
            }
            finally
            {
                ReleaseComObject(mail);
            }
        }, nameof(CreateDraftAsync));
    }

    private dynamic BuildMailItemFromRequest(SendMailRequest request)
    {
        EnsureOutlookInitialized();

        dynamic? sourceItem = null;
        dynamic? mail = null;
        try
        {
            mail = CreateMailItemForRequest(request, out sourceItem);
            ApplyMailProperties(mail, request, sourceItem);
            return mail;
        }
        catch
        {
            ReleaseComObject(mail);
            throw;
        }
        finally
        {
            if (sourceItem is not null)
            {
                ReleaseComObject(sourceItem);
            }
        }
    }

    private void EnsureOutlookInitialized()
    {
        if (_outlookApp is null || _mapiNamespace is null)
        {
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                "Outlook-Application nicht initialisiert (GetOutlookApplicationAsync zuerst aufrufen)");
        }
    }

    private dynamic CreateMailItemForRequest(SendMailRequest request, out dynamic? sourceItem)
    {
        sourceItem = null;
        if (!string.IsNullOrEmpty(request.ForwardFromId))
        {
            sourceItem = _mapiNamespace!.GetItemFromID(request.ForwardFromId, Type.Missing);
            return sourceItem!.Forward();
        }
        if (!string.IsNullOrEmpty(request.ReplyToId))
        {
            sourceItem = _mapiNamespace!.GetItemFromID(request.ReplyToId, Type.Missing);
            return request.ReplyAll ? sourceItem!.ReplyAll() : sourceItem!.Reply();
        }
        return _outlookApp!.CreateItem(0);
    }

    private void ApplyMailProperties(dynamic mail, SendMailRequest request, dynamic? sourceItem)
    {
        if (sourceItem is null && !string.IsNullOrEmpty(request.Subject))
        {
            mail.Subject = request.Subject;
        }

        if (request.To.Count > 0) mail.To = string.Join("; ", request.To);
        if (request.Cc.Count > 0) mail.Cc = string.Join("; ", request.Cc);
        if (request.Bcc.Count > 0) mail.Bcc = string.Join("; ", request.Bcc);

        var bodyFormat = OlEnumMappings.ToOlBodyFormat(request.Body.ContentType);
        mail.BodyFormat = bodyFormat;
        if (bodyFormat == 2)
        {
            mail.HTMLBody = request.Body.Content;
        }
        else
        {
            mail.Body = request.Body.Content;
        }

        mail.Importance = OlEnumMappings.ToOlImportance(request.Importance);

        AddAttachments(mail, request.Attachments);
    }

    private static void AddAttachments(dynamic mail, IReadOnlyList<InlineAttachment> attachments)
    {
        if (attachments.Count == 0) return;
        foreach (var att in attachments)
        {
            var tempPath = Path.Combine(Path.GetTempPath(), $"outlook-mcp-attach-{Guid.NewGuid():N}.bin");
            try
            {
                var bytes = Convert.FromBase64String(att.ContentBase64);
                File.WriteAllBytes(tempPath, bytes);
                mail.Attachments.Add(tempPath, 1, Type.Missing, att.Name);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }
    }

    public async Task UpdateMailAsync(
        string id,
        MailUpdate update,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic mail = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                if (update.IsRead is { } isRead)
                {
                    mail.UnRead = !isRead;
                }
                if (update.Importance is { } importance)
                {
                    mail.Importance = OlEnumMappings.ToOlImportance(importance);
                }
                if (update.Categories is not null)
                {
                    mail.Categories = string.Join(", ", update.Categories);
                }
                mail.Save();
                return true;
            }
            finally
            {
                ReleaseComObject(mail);
            }
        }, nameof(UpdateMailAsync));
    }

    public async Task<string> MoveMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic mail = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                dynamic destFolder = GetFolderByIdOrWellKnownName(destinationFolderId)
                    ?? throw new OutlookServiceException(
                        ErrorCode.FolderNotFound,
                        $"Zielordner nicht gefunden: '{destinationFolderId}'.");
                try
                {
                    dynamic moved = mail.Move(destFolder);
                    try
                    {
                        return (string)moved.EntryID;
                    }
                    finally
                    {
                        ReleaseComObject(moved);
                    }
                }
                finally
                {
                    ReleaseComObject(destFolder);
                }
            }
            finally
            {
                ReleaseComObject(mail);
            }
        }, nameof(MoveMailAsync));
    }

    public async Task<string> CopyMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic mail = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                dynamic copy = mail.Copy();
                try
                {
                    dynamic destFolder = GetFolderByIdOrWellKnownName(destinationFolderId)
                        ?? throw new OutlookServiceException(
                            ErrorCode.FolderNotFound,
                            $"Zielordner nicht gefunden: '{destinationFolderId}'.");
                    try
                    {
                        dynamic moved = copy.Move(destFolder);
                        try
                        {
                            return (string)moved.EntryID;
                        }
                        finally
                        {
                            ReleaseComObject(moved);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(destFolder);
                    }
                }
                finally
                {
                    ReleaseComObject(copy);
                }
            }
            finally
            {
                    ReleaseComObject(mail);
            }
        }, nameof(CopyMailAsync));
    }

    public async Task DeleteMailAsync(
        string id,
        bool permanent = false,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic mail = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                if (permanent)
                {
                    dynamic deletedItems = GetMapiNamespace().GetDefaultFolder(5);
                    try
                    {
                        dynamic inDeleted = mail.Move(deletedItems);
                        try
                        {
                            inDeleted.Delete();
                        }
                        finally
                        {
                            ReleaseComObject(inDeleted);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(deletedItems);
                    }
                }
                else
                {
                    dynamic deletedItems = GetMapiNamespace().GetDefaultFolder(5);
                    try
                    {
                        mail.Move(deletedItems);
                    }
                    finally
                    {
                        ReleaseComObject(deletedItems);
                    }
                }
                return true;
            }
            finally
            {
                ReleaseComObject(mail);
            }
        }, nameof(DeleteMailAsync));
    }
}
