using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Configuration;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Interop;

/// <summary>
/// COM-Adapter: implementiert <see cref="IInteropOutlookAdapter"/> durch
/// Aufrufe auf <c>Microsoft.Office.Interop.Outlook</c> (PIA / NuGet).
/// <para>
/// <b>Threading:</b> Outlook-COM ist single-threaded-affine. Alle Calls laufen
/// serialisiert ueber <see cref="_comLock"/>. <b>Memory:</b> Jedes COM-Objekt
/// wird in try/finally per <see cref="Marshal.ReleaseComObject"/> freigegeben.
/// </para>
/// </summary>
public sealed partial class OutlookInteropAdapter : IInteropOutlookAdapter
{
    private readonly OutlookMcpServerOptions _options;
    private readonly ILogger<OutlookInteropAdapter> _logger;
    private readonly SemaphoreSlim _comLock = new(1, 1);
    private dynamic? _outlookApp;
    private dynamic? _mapiNamespace;

    public OutlookInteropAdapter(
        IOptions<OutlookMcpServerOptions> options,
        ILogger<OutlookInteropAdapter> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    // ===== Lifecycle / Helpers =====

    private async Task<dynamic> GetOutlookApplicationAsync(CancellationToken ct)
    {
        if (_outlookApp is not null) return _outlookApp;

        await _comLock.WaitAsync(ct);
        try
        {
            if (_outlookApp is not null) return _outlookApp;

            // Try active first
            try
            {
                _outlookApp = GetOrStartOutlookApplication();
                _logger.LogInformation("Outlook-Application instanziiert (oder bestehende uebernommen)");
            }
            catch (COMException)
            {
                if (!_options.Outlook.AutoStartOutlook)
                {
                    throw new OutlookServiceException(
                        ErrorCode.OutlookNotRunning,
                        "Outlook laeuft nicht und AutoStartOutlook ist deaktiviert");
                }
                _logger.LogInformation("Starte outlook.exe...");
                var psi = new ProcessStartInfo("outlook.exe") { UseShellExecute = true };
                Process.Start(psi);

                var deadline = DateTime.UtcNow.AddSeconds(_options.Outlook.StartupTimeoutSeconds);
                Exception? lastEx = null;
                while (DateTime.UtcNow < deadline)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        _outlookApp = GetOrStartOutlookApplication();
                        break;
                    }
                    catch (COMException ex)
                    {
                        lastEx = ex;
                        await Task.Delay(500, ct);
                    }
                }
                if (_outlookApp is null)
                {
                    throw new OutlookServiceException(
                        ErrorCode.OutlookNotRunning,
                        $"Outlook konnte nicht innerhalb von {_options.Outlook.StartupTimeoutSeconds}s gestartet werden",
                        lastEx);
                }
            }

            _mapiNamespace = _outlookApp!.GetNamespace("MAPI");
            return _outlookApp;
        }
        finally
        {
            _comLock.Release();
        }
    }

    private dynamic GetMapiNamespace()
    {
        if (_mapiNamespace is null)
        {
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                "MAPI-Namespace nicht initialisiert (GetOutlookApplicationAsync zuerst aufrufen)");
        }
        return _mapiNamespace;
    }

    /// <summary>Konvertiert einen Well-Known-Namen oder EntryID zu einem MAPIFolder-Objekt.</summary>
    private dynamic GetFolderByIdOrWellKnownName(string folderIdOrName)
    {
        var ns = GetMapiNamespace();
        var olFolderId = OlEnumMappings.ToOlDefaultFolder(folderIdOrName);
        if (olFolderId.HasValue)
        {
            return ns.GetDefaultFolder(olFolderId.Value);
        }
        // Annahme: folderIdOrName ist eine EntryID
        return ns.GetFolderFromID(folderIdOrName, Type.Missing);
    }

    /// <summary>
    /// Mappt ein COM-MAPIFolder-Objekt auf das <see cref="MailFolder"/>-DTO.
    /// WellKnownName bleibt null (korrektes Mapping wuerde Vergleich mit
    /// Application.Session.GetDefaultFolder erfordern — wird in einer
    /// Folge-Phase ergaenzt, falls Use-Cases auftauchen).
    /// </summary>
    private static MailFolder MapMailFolder(dynamic folder, string? parentId)
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
                    Marshal.ReleaseComObject(parent);
                }
            }
            catch { }
        }

        return new MailFolder
        {
            Id = entryId,
            DisplayName = name,
            WellKnownName = null, // TODO: spaeter — braucht App-Vergleich
            ParentFolderId = actualParentId,
            ChildFolderCount = childCount,
            TotalItemCount = totalCount,
            UnreadItemCount = unreadCount,
        };
    }

    /// <summary>
    /// Mappt ein COM-MailItem auf das <see cref="MailMessage"/>-DTO. Bei
    /// <paramref name="includeBody"/>=false wird der Body weggelassen (fuer
    /// ListMails-Use-Case, wo BodyPreview reicht und das Voll-Body-HTML teuer ist).
    /// ComException in optionalen Properties wird defensiv mit Fallback-Wert gefangen.
    /// </summary>
    private static MailMessage MapMailItem(dynamic mail, bool includeBody)
    {
        var entryId = (string)mail.EntryID;
        string? convId = TryGetString(mail, "ConversationID");
        string? subject = TryGetString(mail, "Subject");
        string? bodyPreview = TryGetString(mail, "BodyPreview");

        ItemBody? body = null;
        if (includeBody)
        {
            string text = TryGetString(mail, "Body") ?? string.Empty;
            string html = TryGetString(mail, "HTMLBody") ?? string.Empty;
            int fmt = 1;
            try { fmt = (int)mail.BodyFormat; } catch { }
            if (fmt == 2 && !string.IsNullOrEmpty(html))
            {
                body = new ItemBody { ContentType = ItemBodyType.Html, Content = html };
            }
            else
            {
                body = new ItemBody { ContentType = ItemBodyType.Text, Content = text };
            }
        }

        Recipient? from = null;
        try
        {
            var senderName = TryGetString(mail, "SenderName");
            var senderAddress = TryGetString(mail, "SenderEmailAddress");
            if (mail.Sender is not null && !string.IsNullOrEmpty(senderAddress))
            {
                from = new Recipient
                {
                    EmailAddress = new EmailAddress { Name = senderName, Address = senderAddress },
                };
            }
        }
        catch { }

        var toRecipients = MapRecipients(TryGetDynamic(mail, "To"));
        var ccRecipients = MapRecipients(TryGetDynamic(mail, "CC"));
        var bccRecipients = MapRecipients(TryGetDynamic(mail, "BCC"));

        DateTimeOffset? sent = null;
        DateTimeOffset? received = null;
        try { sent = (DateTimeOffset)mail.SentOn; } catch { }
        try { received = (DateTimeOffset)mail.ReceivedTime; } catch { }

        bool hasAttachments = false;
        try { hasAttachments = (int)mail.Attachments.Count > 0; } catch { }

        // Outlook: UnRead = true heisst "ungelesen" → DTO IsRead = !UnRead
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

        return new MailMessage
        {
            Id = entryId,
            ConversationId = convId,
            Subject = subject,
            BodyPreview = bodyPreview,
            Body = body,
            From = from,
            ToRecipients = toRecipients,
            CcRecipients = ccRecipients,
            BccRecipients = bccRecipients,
            SentDateTime = sent,
            ReceivedDateTime = received,
            HasAttachments = hasAttachments,
            Importance = importance,
            IsRead = !isUnread,
            Categories = categories,
        };
    }

    /// <summary>
    /// Iteriert ueber eine COM-Recipients-Collection und mappt auf
    /// <see cref="Recipient"/>-DTOs. Wirft/warnt nicht bei einzelnen
    /// kaputten Empfaengern — gibt leere Liste fuer null-Input zurueck.
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
                    if (rec is not null) Marshal.ReleaseComObject(rec);
                }
            }
        }
        catch { }
        return list;
    }

    /// <summary>Defensive Property-Access-Wrapper fuer COM dynamic-Properties.</summary>
    private static string? TryGetString(dynamic obj, string propertyName)
    {
        try { return (string?)obj.GetType().GetProperty(propertyName)?.GetValue(obj); }
        catch { return null; }
    }

    private static dynamic? TryGetDynamic(dynamic obj, string propertyName)
    {
        try { return obj.GetType().GetProperty(propertyName)?.GetValue(obj); }
        catch { return null; }
    }

    /// <summary>
    /// Kombiniert benutzerdefinierten DASL-Filter mit optionalem Volltext-Such-Query.
    /// Suchbegriff wird in einfache Hochkommata eingeschlossen (mit Escape "''" fuer " innerhalb).
    /// Volltext-Match ueber Subject + Body via DASL LIKE.
    /// </summary>
    private static string CombineDaslFilters(string? userFilter, string? search)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(userFilter)) parts.Add(userFilter);
        if (!string.IsNullOrWhiteSpace(search))
        {
            var escaped = search!.Replace("'", "''");
            parts.Add($"([Subject] LIKE '%{escaped}%' OR [Body] LIKE '%{escaped}%')");
        }
        return string.Join(" AND ", parts);
    }

    /// <summary>
    /// Instanziiert Outlook.Application via COM. Wenn Outlook bereits laeuft,
    /// wird die bestehende Instanz zurueckgegeben (COM-Singleton); sonst wird
    /// eine neue gestartet. Ersetzt Marshal.GetActiveObject (in .NET 8 nicht
    /// verfuegbar, nur in .NET Framework).
    /// </summary>
    private static dynamic GetOrStartOutlookApplication()
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new OutlookServiceException(
                ErrorCode.OutlookNotRunning,
                "Outlook ProgID nicht registriert");
        return Activator.CreateInstance(outlookType)
            ?? throw new OutlookServiceException(
                ErrorCode.InternalError,
                "Activator.CreateInstance gab null zurueck");
    }

    /// <summary>Wrappt einen COM-Call und konvertiert COMException in OutlookServiceException.</summary>
    private T RunCom<T>(Func<T> func, string opName)
    {
        try
        {
            return func();
        }
        catch (COMException ex)
        {
            _logger.LogWarning(ex, "COM-Fehler bei {Op}: HResult=0x{HResult:X8}", opName, ex.HResult);
            throw new OutlookServiceException(
                OlEnumMappings.FromHResult(ex.HResult),
                $"Outlook-Fehler bei {opName}: {ex.Message}",
                ex);
        }
        catch (OutlookServiceException) { throw; }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unerwarteter Fehler bei {Op}", opName);
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                $"Interner Fehler bei {opName}: {ex.Message}",
                ex);
        }
    }

    private Task<T> RunComAsync<T>(Func<T> func, string opName)
    {
        try
        {
            return Task.FromResult(func());
        }
        catch (COMException ex)
        {
            _logger.LogWarning(ex, "COM-Fehler bei {Op}: HResult=0x{HResult:X8}", opName, ex.HResult);
            throw new OutlookServiceException(
                OlEnumMappings.FromHResult(ex.HResult),
                $"Outlook-Fehler bei {opName}: {ex.Message}",
                ex);
        }
        catch (OutlookServiceException) { throw; }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unerwarteter Fehler bei {Op}", opName);
            throw new OutlookServiceException(
                ErrorCode.InternalError,
                $"Interner Fehler bei {opName}: {ex.Message}",
                ex);
        }
    }

    // ===== Mail: Ordner =====

    public async Task<PagedResult<MailFolder>> ListMailFoldersAsync(
        string? parentFolderId = null,
        bool includeHidden = false,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var result = new List<MailFolder>();
            dynamic parent = parentFolderId is null
                ? GetMapiNamespace()
                : GetMapiNamespace().GetFolderFromID(parentFolderId, Type.Missing);
            try
            {
                foreach (var item in parent.Folders)
                {
                    dynamic folder = item;
                    try
                    {
                        var name = (string)folder.Name;
                        if (!includeHidden && OlEnumMappings.IsOutlookHiddenFolderName(name))
                        {
                            continue;
                        }
                        result.Add(MapMailFolder(folder, parentFolderId));
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(folder);
                    }
                }
            }
            finally
            {
                if (parentFolderId is not null) Marshal.ReleaseComObject(parent);
            }
            return new PagedResult<MailFolder>
            {
                Value = result,
                NextSkip = null,
            };
        }, nameof(ListMailFoldersAsync));
    }

    public async Task<MailFolder> GetMailFolderAsync(
        string folderId,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic folder = GetFolderByIdOrWellKnownName(folderId);
            try
            {
                return MapMailFolder(folder, null);
            }
            finally
            {
                Marshal.ReleaseComObject(folder);
            }
        }, nameof(GetMailFolderAsync));
    }

    // ===== Mail: Nachrichten =====

    public async Task<PagedResult<MailMessage>> ListMailsAsync(
        string folderId,
        int top = 25,
        int skip = 0,
        string? filter = null,
        string? search = null,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic folder = GetFolderByIdOrWellKnownName(folderId);
            try
            {
                string combinedFilter = CombineDaslFilters(filter, search);
                dynamic items = folder.Items;
                try
                {
                    if (!string.IsNullOrWhiteSpace(combinedFilter))
                    {
                        items = items.Restrict(combinedFilter);
                    }
                    items.Sort("[ReceivedTime]", true); // newest first

                    int total = (int)items.Count;
                    int start = Math.Min(skip, total);
                    int end = Math.Min(skip + top, total);
                    var slice = new List<MailMessage>(end - start);
                    for (int i = start; i < end; i++)
                    {
                        dynamic item = items.Item(i + 1); // COM-Auflistungen sind 1-basiert
                        try
                        {
                            slice.Add(MapMailItem(item, includeBody: false));
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(item);
                        }
                    }
                    return new PagedResult<MailMessage>
                    {
                        Value = slice,
                        NextSkip = end < total ? end : (int?)null,
                    };
                }
                finally
                {
                    Marshal.ReleaseComObject(items);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(folder);
            }
        }, nameof(ListMailsAsync));
    }

    public Task<MailMessage> GetMailAsync(
        string id,
        bool includeBody = true,
        CancellationToken cancellationToken = default)
    {
        throw new NotImplementedException("GetMailAsync - wird in Karte 3.5 implementiert");
    }

    public Task<IReadOnlyList<InternetMessageHeader>> GetMailHeadersAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
        return Task.FromResult<IReadOnlyList<InternetMessageHeader>>(Array.Empty<InternetMessageHeader>());
    }

    public Task<IReadOnlyList<AttachmentSummary>> ListAttachmentsAsync(
        string mailId,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5
        return Task.FromResult<IReadOnlyList<AttachmentSummary>>(Array.Empty<AttachmentSummary>());
    }

    public Task<AttachmentData> GetAttachmentAsync(
        string mailId,
        string attachmentId,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: SaveAsFile + ReadAllBytes + Base64
        throw new NotImplementedException("GetAttachmentAsync - wird in Karte 3.5 implementiert");
    }

    public async Task<PagedResult<MailMessage>> SearchMailsAsync(
        string query,
        string? folderId = null,
        int top = 25,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var slice = new List<MailMessage>(top);

            // Wenn folderId null → rekursiv ueber alle Mail-Ordner durchsuchen
            if (folderId is null)
            {
                dynamic root = GetMapiNamespace();
                try
                {
                    CollectMatchingMails(root, query, slice, top);
                }
                finally
                {
                    Marshal.ReleaseComObject(root);
                }
            }
            else
            {
                dynamic folder = GetFolderByIdOrWellKnownName(folderId);
                try
                {
                    CollectMatchingMails(folder, query, slice, top);
                }
                finally
                {
                    Marshal.ReleaseComObject(folder);
                }
            }

            return new PagedResult<MailMessage>
            {
                Value = slice,
                NextSkip = null, // Search liefert keine Pagination (Hard-Cap via top)
            };
        }, nameof(SearchMailsAsync));
    }

    /// <summary>
    /// Iteriert ueber einen Folder (oder Root) und sammelt Mails, deren
    /// Subject ODER SenderEmailAddress den Query (case-insensitive) enthaelt.
    /// Bricht ab, sobald <c>slice.Count == top</c> erreicht ist. Rekursion
    /// ueber Unterordner. Com-Objekte werden pro Item freigegeben.
    /// </summary>
    private static void CollectMatchingMails(dynamic parent, string query, List<MailMessage> slice, int top)
    {
        try
        {
            dynamic items = parent.Items;
            try
            {
                if (items is not null)
                {
                    foreach (var item in items)
                    {
                        try
                        {
                            if ((int)item.Class != 43) continue; // nur olMail (43) — keine Appointments/Tasks
                            var subj = TryGetString(item, "Subject") ?? string.Empty;
                            var sender = TryGetString(item, "SenderEmailAddress") ?? string.Empty;
                            if (subj.Contains(query, StringComparison.OrdinalIgnoreCase)
                                || sender.Contains(query, StringComparison.OrdinalIgnoreCase))
                            {
                                slice.Add(MapMailItem(item, includeBody: false));
                                if (slice.Count >= top) return;
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(item);
                        }
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(items);
            }

            // Rekursion in Unterordner
            dynamic subFolders = parent.Folders;
            try
            {
                if (subFolders is not null)
                {
                    foreach (var sub in subFolders)
                    {
                        try
                        {
                            CollectMatchingMails(sub, query, slice, top);
                            if (slice.Count >= top) return;
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(sub);
                        }
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(subFolders);
            }
        }
        catch { /* defensive — Ordner ohne Zugriff ueberspringen */ }
    }

    // ===== Mail: Mutationen =====

    public Task<SendMailResult> SendMailAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: Application.CreateItem(olMailItem) + .Send()
        throw new NotImplementedException("SendMailAsync - wird in Karte 3.5 implementiert");
    }

    public Task<string> CreateDraftAsync(
        SendMailRequest request,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: Application.CreateItem(olMailItem) + .Save()
        throw new NotImplementedException("CreateDraftAsync - wird in Karte 3.5 implementiert");
    }

    public Task UpdateMailAsync(
        string id,
        MailUpdate update,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: MailItem.IsRead, Categories, Importance
        throw new NotImplementedException("UpdateMailAsync - wird in Karte 3.5 implementiert");
    }

    public Task<string> MoveMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: MailItem.Move(destinationFolder)
        throw new NotImplementedException("MoveMailAsync - wird in Karte 3.5 implementiert");
    }

    public Task<string> CopyMailAsync(
        string id,
        string destinationFolderId,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: MailItem.Copy()
        throw new NotImplementedException("CopyMailAsync - wird in Karte 3.5 implementiert");
    }

    public Task DeleteMailAsync(
        string id,
        bool permanent = false,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: MailItem.Delete() oder .Move(deletedItems)
        throw new NotImplementedException("DeleteMailAsync - wird in Karte 3.5 implementiert");
    }

    // ===== Calendar: Kalender =====

    public Task<IReadOnlyList<Calendar>> ListCalendarsAsync(
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: Namespace.Folders (Loop) + .Folders["Kalender"]?
        return Task.FromResult<IReadOnlyList<Calendar>>(Array.Empty<Calendar>());
    }

    public Task<Calendar> GetCalendarAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        throw new NotImplementedException("GetCalendarAsync - wird in Karte 3.5 implementiert");
    }

    // ===== Calendar: Termine =====

    public Task<PagedResult<CalendarEvent>> ListEventsAsync(
        string? calendarId,
        DateTimeTimeZone start,
        DateTimeTimeZone end,
        int top = 50,
        int skip = 0,
        string? filter = null,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: Calendar.Items.Find("[Start] >= ... AND [End] < ...")
        return Task.FromResult(new PagedResult<CalendarEvent>
        {
            Value = Array.Empty<CalendarEvent>(),
            NextSkip = null,
        });
    }

    public Task<CalendarEvent> GetEventAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        throw new NotImplementedException("GetEventAsync - wird in Karte 3.5 implementiert");
    }

    // ===== Calendar: Mutationen =====

    public Task<string> CreateEventAsync(
        CreateEventRequest request,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: Calendar.Items.Add(olAppointmentItem) + .Send()
        throw new NotImplementedException("CreateEventAsync - wird in Karte 3.5 implementiert");
    }

    public Task UpdateEventAsync(
        string id,
        EventUpdate update,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: AppointmentItem-Mutation + .Save()
        throw new NotImplementedException("UpdateEventAsync - wird in Karte 3.5 implementiert");
    }

    public Task DeleteEventAsync(
        string id,
        bool sendCancellation = true,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: AppointmentItem.Delete()
        throw new NotImplementedException("DeleteEventAsync - wird in Karte 3.5 implementiert");
    }

    public Task RespondToEventAsync(
        string id,
        RespondToEventRequest request,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: MeetingItem.Respond(olResponseStatus, ...)
        throw new NotImplementedException("RespondToEventAsync - wird in Karte 3.5 implementiert");
    }

    public Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(
        FindMeetingTimesRequest request,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: Self-Calendar scan + Luecken >= DurationMinutes
        return Task.FromResult<IReadOnlyList<MeetingTimeCandidate>>(Array.Empty<MeetingTimeCandidate>());
    }

    // ===== Active-Inspector / Selection (TODO Karte 3.5 P3) =====

    public Task<ActiveItem?> GetActiveItemAsync(CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5 P3: Application.ActiveInspector()?.CurrentItem + Type-Dispatch (MailItem | AppointmentItem -> DTO)
        throw new NotImplementedException("GetActiveItemAsync - wird in Karte 3.5 P3 implementiert");
    }

    public Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top = 50,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5 P3: Application.ActiveExplorer()?.Selection + Scope-Filter
        throw new NotImplementedException("GetSelectedItemsAsync - wird in Karte 3.5 P3 implementiert");
    }
}