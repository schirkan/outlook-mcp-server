using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    public async Task<PagedResult<MailFolder>> ListMailFoldersAsync(
        string? parentFolderId = null,
        bool includeHidden = false,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var result = new List<MailFolder>();
            dynamic ns = GetMapiNamespace();

            // Liste der WellKnown-DefaultFolder-COM-Referenzen, die wir zum
            // Vergleich brauchen, um WellKnownName zu setzen.
            var wellKnownByEntryId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (wkName, olId) in OlEnumMappings.KnownFolderMappings)
            {
                try
                {
                    dynamic wkFolder = ns.GetDefaultFolder(olId);
                    string wkEntryId = (string)wkFolder.EntryID;
                    if (!string.IsNullOrEmpty(wkEntryId))
                    {
                        wellKnownByEntryId[wkEntryId] = wkName;
                    }
                    Marshal.ReleaseComObject(wkFolder);
                }
                catch (COMException ex)
                {
                    _logger.LogDebug(ex,
                        "ListMailFolders: WellKnownFolder '{Name}' (olId={OlId}) nicht aufloesbar",
                        wkName, olId);
                }
            }

            // Outlook-Semantik: ns.Folders enthaelt die Postfaecher/Stores
            // (z. B. "Martin.Pietschmann@itergo.com"). Ein Aufrufer ohne
            // parentFolderId will die *sichtbaren Mailordner* sehen — daher
            // iterieren wir bei parentFolderId==null ueber alle Postfaecher
            // und geben deren direkte Kinder (Posteingang, Gesendete Elemente,
            // ...) zurueck. Damit verhaelt sich der Service wie Outlook-UI.
            if (parentFolderId is null)
            {
                dynamic topFolders = ns.Folders;
                try
                {
                    foreach (var storeRoot in topFolders)
                    {
                        dynamic? storeKids = null;
                        try
                        {
                            try { storeKids = storeRoot.Folders; }
                            catch (COMException ex)
                            {
                                _logger.LogWarning(ex,
                                    "ListMailFolders: Postfach '{Name}' hat keine Folders",
                                    (string)storeRoot.Name);
                                continue;
                            }
                            foreach (var folder in storeKids)
                            {
                                try
                                {
                                    var name = (string)folder.Name;
                                    if (!includeHidden && OlEnumMappings.IsOutlookHiddenFolderName(name))
                                    {
                                        continue;
                                    }
                                    string entryId = (string)folder.EntryID;
                                    string? wellKnown = wellKnownByEntryId.TryGetValue(entryId, out var wk) ? wk : null;
                                    result.Add(MapMailFolder(folder, null, wellKnown));
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(folder);
                                }
                            }
                        }
                        finally
                        {
                            if (storeKids is not null) Marshal.ReleaseComObject(storeKids);
                            Marshal.ReleaseComObject(storeRoot);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(topFolders);
                }
            }
            else
            {
                dynamic parent = ns.GetFolderFromID(parentFolderId, Type.Missing);
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
                            string entryId = (string)folder.EntryID;
                            string? wellKnown = wellKnownByEntryId.TryGetValue(entryId, out var wk) ? wk : null;
                            result.Add(MapMailFolder(folder, parentFolderId, wellKnown));
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(folder);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(parent);
                }
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
            dynamic folder = GetFolderByIdOrWellKnownName(folderId)
                ?? throw new OutlookServiceException(
                    ErrorCode.FolderNotFound,
                    $"Mail-Ordner nicht gefunden: '{folderId}'. " +
                    "Bei Well-Known-Namen (inbox, drafts, sentItems, deletedItems, " +
                    "junkEmail, archive, outbox) wird in Multi-Store-Profilen " +
                    "unter allen Stores gesucht — falls keiner passt, ist der " +
                    "Folder im aktiven Profil nicht vorhanden.");
            // WellKnownName nur setzen, wenn der Caller einen Well-Known-Namen
            // verwendet hat (sonst wurde per EntryID aufgeloest, und wir koennen
            // den Folder-Charakter nicht ohne Lookup bestimmen).
            string? wellKnown = WellKnownFolder.IsKnownMailFolder(folderId) ? folderId : null;
            try
            {
                return MapMailFolder(folder, null, wellKnown);
            }
            finally
            {
                ReleaseComObject(folder);
            }
        }, nameof(GetMailFolderAsync));
    }

    public async Task<PagedResult<MailMessage>> ListMailsAsync(
        string folderId,
        int top = 25,
        int skip = 0,
        string? filter = null,
        string? search = null,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        _logger.LogInformation(
            "ListMails START folderId={FolderId} top={Top} skip={Skip} filter={Filter} search={Search}",
            folderId, top, skip, filter, search);

        return await RunComAsync(() =>
        {
            string? normalizedFilter = filter;
            if (WellKnownFolder.IsKnownMailFolder(folderId))
            {
                normalizedFilter = NormalizeDaslFilter(filter);
            }

            dynamic folder = GetFolderByIdOrWellKnownName(folderId)
                ?? throw new OutlookServiceException(
                    ErrorCode.FolderNotFound,
                    $"Mail-Ordner nicht gefunden: '{folderId}'.");

            string? folderName = null;
            try { folderName = (string)folder.Name; } catch { }
            _logger.LogInformation(
                "ListMails folder resolved: displayName={FolderName} folderId={FolderId}",
                folderName, folderId);

            dynamic? items = null;
            try
            {
                items = folder.Items;

                int rawCount = 0;
                try { rawCount = (int)items.Count; } catch { }
                _logger.LogInformation(
                    "ListMails raw item count before Restrict: {RawCount} (folder={FolderName})",
                    rawCount, folderName);

                if (!string.IsNullOrWhiteSpace(normalizedFilter))
                {
                    var originalItems = items;
                    items = RestrictMailItems((object)originalItems, normalizedFilter);
                    if (!ReferenceEquals(originalItems, items)) ReleaseComObject(originalItems);
                }

                try
                {
                    items.Sort("[ReceivedTime]", true);
                }
                catch (COMException sortEx)
                {
                    _logger.LogWarning(sortEx,
                        "ListMails Sort nach [ReceivedTime] fehlgeschlagen (HResult=0x{HResult:X8}); " +
                        "Fallback: unsortierte Iteration. folder={FolderName}",
                        sortEx.HResult, folderName);
                }

                int total = 0;
                try { total = (int)items.Count; } catch { }

                bool hasSearch = !string.IsNullOrWhiteSpace(search);
                int start = hasSearch ? 0 : Math.Min(skip, total);
                int end = hasSearch ? total : start + Math.Min(top, total - start);
                var slice = new List<MailMessage>(top);
                int matched = 0;
                bool moreMatches = false;
                int skippedNonMail = 0;
                int failedItems = 0;

                for (int i = start; i < end; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    dynamic? item = null;
                    try
                    {
                        item = items.Item(i + 1);
                    }
                    catch (COMException itemEx)
                    {
                        failedItems++;
                        _logger.LogWarning(itemEx,
                            "ListMails ueberspringe Item-Index {Index} (HResult=0x{HResult:X8}); " +
                            "Item nicht lesbar im Folder {FolderName}",
                            i + 1, itemEx.HResult, folderName);
                        continue;
                    }

                    try
                    {
                        int itemClass = 0;
                        try { itemClass = (int)item.Class; } catch { }
                        if (itemClass != 43)
                        {
                            skippedNonMail++;
                            continue;
                        }

                        if (hasSearch)
                        {
                            string subject = TryGetString(item, "Subject") ?? string.Empty;
                            if (!subject.Contains(search!, StringComparison.OrdinalIgnoreCase))
                            {
                                string body = TryGetString(item, "Body") ?? string.Empty;
                                if (!body.Contains(search!, StringComparison.OrdinalIgnoreCase)) continue;
                            }
                            if (matched++ < skip) continue;
                            if (slice.Count == top)
                            {
                                moreMatches = true;
                                break;
                            }
                        }
                        slice.Add(MapMailItem(item, null, _logger));
                    }
                    catch (COMException mapEx)
                    {
                        failedItems++;
                        _logger.LogWarning(mapEx,
                            "ListMails ueberspringe Mail-Item {Index} (HResult=0x{HResult:X8}); " +
                            "MapMailItem fehlgeschlagen im Folder {FolderName}",
                            i + 1, mapEx.HResult, folderName);
                    }
                    catch (Exception mapEx)
                    {
                        failedItems++;
                        _logger.LogError(mapEx,
                            "ListMails unerwarteter Fehler beim Mapping von Item {Index} im Folder {FolderName}",
                            i + 1, folderName);
                    }
                    finally
                    {
                        ReleaseComObject(item);
                    }
                }

                _logger.LogInformation(
                    "ListMails DONE folder={FolderName} returned={Returned} skippedNonMail={SkippedNonMail} " +
                    "failedItems={FailedItems} total={Total} start={Start} end={End}",
                    folderName, slice.Count, skippedNonMail, failedItems, total, start, end);

                return new PagedResult<MailMessage>
                {
                    Value = slice,
                    NextSkip = hasSearch ? (moreMatches ? matched - 1 : null) : (end < total ? end : null),
                };
            }
            finally
            {
                ReleaseComObject(items);
                ReleaseComObject(folder);
            }
        }, nameof(ListMailsAsync));
    }

    public async Task<MailMessage> GetMailAsync(
        string id,
        BodyFormat? bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic mail = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                return MapMailItem(mail, bodyFormat, _logger);
            }
            finally
            {
                ReleaseComObject(mail);
            }
        }, nameof(GetMailAsync));
    }

    public async Task<BulkMailResult> GetMailsAsync(
        IReadOnlyList<string> ids,
        BodyFormat? bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        var ns = GetMapiNamespace();
        var found = new List<MailMessage>(ids.Count);
        var notFound = new List<string>();
        foreach (var id in ids)
        {
            dynamic? mail = null;
            try
            {
                mail = ns.GetItemFromID(id, Type.Missing);
                found.Add(MapMailItem(mail, bodyFormat));
            }
            catch (COMException ex)
            {
                _logger.LogInformation(
                    "GetMails: ID '{Id}' nicht aufloesbar (HResult=0x{HResult:X8}: {Msg})",
                    id, ex.HResult, ex.Message);
                notFound.Add(id);
            }
            finally
            {
                ReleaseComObject(mail);
            }
        }
        return new BulkMailResult { Value = found, NotFoundIds = notFound };
    }

    public async Task<IReadOnlyList<InternetMessageHeader>> GetMailHeadersAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var headers = new List<InternetMessageHeader>();
            dynamic mail = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                string raw = (string)mail.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x007D001F");
                foreach (var rawLine in raw.Split('\n'))
                {
                    var line = rawLine.TrimEnd('\r').Trim();
                    if (string.IsNullOrEmpty(line)) continue;
                    int idx = line.IndexOf(':');
                    if (idx <= 0) continue;
                    var name = line.Substring(0, idx).Trim();
                    var value = line.Substring(idx + 1).Trim();
                    if (name.Length == 0) continue;
                    headers.Add(new InternetMessageHeader { Name = name, Value = value });
                }
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x80040401))
            {
                _logger.LogInformation("Mail {Id} hat keine Internet-Header (PropertyAccess nicht vorhanden)", id);
            }
            finally
            {
                ReleaseComObject(mail);
            }
            return (IReadOnlyList<InternetMessageHeader>)headers;
        }, nameof(GetMailHeadersAsync));
    }

    public async Task<IReadOnlyList<AttachmentSummary>> ListAttachmentsAsync(
        string mailId,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var summaries = new List<AttachmentSummary>();
            dynamic mail = GetMapiNamespace().GetItemFromID(mailId, Type.Missing);
            try
            {
                foreach (var attachment in mail.Attachments)
                {
                    try
                    {
                        var position = 0;
                        try { position = (int)attachment.Position; } catch { }
                        summaries.Add(new AttachmentSummary
                        {
                            Id = (string)attachment.EntryID,
                            Name = (string)attachment.DisplayName ?? string.Empty,
                            ContentType = TryGetString(attachment, "ContentType") ?? "application/octet-stream",
                            Size = (long)(int)attachment.Size,
                            IsInline = position > 0,
                        });
                    }
                    finally
                    {
                        ReleaseComObject(attachment);
                    }
                }
            }
            finally
            {
                ReleaseComObject(mail);
            }
            return (IReadOnlyList<AttachmentSummary>)summaries;
        }, nameof(ListAttachmentsAsync));
    }

    public async Task<AttachmentData> GetAttachmentAsync(
        string mailId,
        string attachmentId,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var tempPath = Path.Combine(Path.GetTempPath(), $"outlook-mcp-{Guid.NewGuid():N}.bin");
            try
            {
                dynamic mail = GetMapiNamespace().GetItemFromID(mailId, Type.Missing);
                try
                {
                    AttachmentData? result = null;
                    foreach (var attachment in mail.Attachments)
                    {
                        try
                        {
                            if ((string)attachment.EntryID == attachmentId)
                            {
                                attachment.SaveAsFile(tempPath);
                                var bytes = File.ReadAllBytes(tempPath);
                                result = new AttachmentData
                                {
                                    Id = (string)attachment.EntryID,
                                    Name = (string)attachment.DisplayName ?? "attachment",
                                    ContentType = TryGetString(attachment, "ContentType") ?? "application/octet-stream",
                                    Size = (long)(int)attachment.Size,
                                    ContentBase64 = Convert.ToBase64String(bytes),
                                };
                                break;
                            }
                        }
                        finally
                        {
                            ReleaseComObject(attachment);
                        }
                    }
                    if (result is null)
                    {
                        throw new OutlookServiceException(
                            ErrorCode.AttachmentNotFound,
                            $"Attachment mit EntryID '{attachmentId}' nicht gefunden in Mail '{mailId}'");
                    }
                    return result;
                }
                finally
                {
                    ReleaseComObject(mail);
                }
            }
            finally
            {
                if (File.Exists(tempPath)) File.Delete(tempPath);
            }
        }, nameof(GetAttachmentAsync));
    }

    public async Task<PagedResult<MailMessage>> SearchMailsAsync(
        string query,
        string? folderId = null,
        int top = 25,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        _logger.LogInformation(
            "SearchMails START query={Query} folderId={FolderId} top={Top}",
            query, folderId, top);
        return await RunComAsync(() =>
        {
            var slice = new List<MailMessage>(top);
            int visitedFolders = 0;
            int skippedFolders = 0;
            int skippedNonMail = 0;
            int failedItems = 0;

            if (folderId is null)
            {
                dynamic root = GetMapiNamespace();
                try
                {
                    dynamic stores = root.Stores;
                    bool earlyExit = false;
                    try
                    {
                        foreach (var store in stores)
                        {
                            if (earlyExit) break;
                            try
                            {
                                dynamic storeRoot = store.GetRootFolder();
                                try
                                {
                                    dynamic topFolders = storeRoot.Folders;
                                    try
                                    {
                                        foreach (var f in topFolders)
                                        {
                                            if (earlyExit) break;
                                            try
                                            {
                                                CollectMatchingMails(f, query, slice, top,
                                                    ref visitedFolders, ref skippedFolders,
                                                    ref skippedNonMail, ref failedItems);
                                                if (slice.Count >= top) earlyExit = true;
                                            }
                                            finally
                                            {
                                                ReleaseComObject(f);
                                            }
                                        }
                                    }
                                    finally { ReleaseComObject(topFolders); }
                                }
                                finally { ReleaseComObject(storeRoot); }
                            }
                            finally { ReleaseComObject(store); }
                        }
                    }
                    finally { ReleaseComObject(stores); }
                }
                catch (COMException ex)
                {
                    skippedFolders++;
                    _logger.LogWarning(ex,
                        "SearchMails Fehler im Store-Walk (HResult=0x{HResult:X8})",
                        ex.HResult);
                }
            }
            else
            {
                dynamic folder = GetFolderByIdOrWellKnownName(folderId)
                    ?? throw new OutlookServiceException(
                        ErrorCode.FolderNotFound,
                        $"Mail-Ordner nicht gefunden: '{folderId}'.");
                try
                {
                    CollectMatchingMails(folder, query, slice, top,
                        ref visitedFolders, ref skippedFolders, ref skippedNonMail, ref failedItems);
                }
                finally
                {
                    ReleaseComObject(folder);
                }
            }

            _logger.LogInformation(
                "SearchMails DONE query={Query} returned={Returned} visitedFolders={Visited} " +
                "skippedFolders={SkippedFolders} skippedNonMail={SkippedNonMail} failedItems={FailedItems}",
                query, slice.Count, visitedFolders, skippedFolders, skippedNonMail, failedItems);

            return new PagedResult<MailMessage>
            {
                Value = slice,
                NextSkip = null,
            };
        }, nameof(SearchMailsAsync));
    }

    public async Task<PagedResult<MailMessage>> ListMailsRecursiveAsync(
        IReadOnlyList<string> scope,
        int top,
        string? filter,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        _logger.LogInformation(
            "ListMailsRecursive START scope=[{Scope}] top={Top} filter={Filter}",
            string.Join(",", scope), top, filter);
        return await RunComAsync(() =>
        {
            var resolvedScope = (scope is null || scope.Count == 0)
                ? new List<string>
                  {
                      WellKnownFolder.Inbox,
                      WellKnownFolder.Drafts,
                      WellKnownFolder.SentItems,
                      WellKnownFolder.DeletedItems,
                      WellKnownFolder.JunkEmail,
                      WellKnownFolder.Archive,
                      WellKnownFolder.Outbox,
                  }
                : new List<string>(scope);

            string? normalizedFilter = NormalizeDaslFilter(filter);
            var collected = new List<(string EntryId, DateTime Received, MailMessage Msg)>(top * 2);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            int visitedFolders = 0, skippedFolders = 0, skippedNonMail = 0,
                failedItems = 0, restrictedItems = 0;

            dynamic root = GetMapiNamespace();
            try
            {
                foreach (var name in resolvedScope)
                {
                    dynamic? folder = null;
                    try
                    {
                        var ol = OlEnumMappings.ToOlDefaultFolder(name);
                        if (ol is null)
                        {
                            _logger.LogWarning(
                                "ListMailsRecursive scope-Eintrag unbekannt: {Name} (kein WellKnownFolder)",
                                name);
                            skippedFolders++;
                            continue;
                        }
                        _logger.LogInformation(
                            "ListMailsRecursive rufe GetDefaultFolder(name={Name}, olId={OlId})",
                            name, ol.Value);
                        try
                        {
                            folder = ResolveDefaultFolderByOlIdWithSchemaCheck(ol.Value);
                        }
                        catch (Exception olEx)
                        {
                            _logger.LogWarning(olEx,
                                "ListMailsRecursive GetDefaultFolder({Name}) fehlgeschlagen; uebersprungen",
                                name);
                            skippedFolders++;
                            continue;
                        }
                        if (folder is null)
                        {
                            _logger.LogWarning(
                                "ListMailsRecursive GetDefaultFolder({Name}) lieferte null " +
                                "(kein Store enthaelt den erwarteten Folder-Namen); uebersprungen",
                                name);
                            skippedFolders++;
                            continue;
                        }
                        _logger.LogInformation(
                            "ListMailsRecursive GetDefaultFolder({Name}) -> Folder.Name={FolderName}",
                            name, (string)(TryGetString(folder, "Name") ?? "?"));
                        try
                        {
                            CollectMailsFromFolderRecursive(
                                folder,
                                normalizedFilter,
                                top,
                                collected,
                                seen,
                                ref visitedFolders,
                                ref skippedFolders,
                                ref skippedNonMail,
                                ref failedItems,
                                ref restrictedItems);
                            _logger.LogInformation(
                                "ListMailsRecursive Top-Level {Name} fertig: collected={Collected}/{Top} visitedFolders={Visited}",
                                name, collected.Count, top, visitedFolders);
                            if (collected.Count >= top) break;
                        }
                        catch (OutlookServiceException) { throw; }
                        catch (COMException ex)
                        {
                            skippedFolders++;
                            _logger.LogWarning(ex,
                                "ListMailsRecursive ueberspringe Top-Level-Folder {Name} (HResult=0x{HResult:X8})",
                                name, ex.HResult);
                        }
                        catch (Exception ex)
                        {
                            skippedFolders++;
                            _logger.LogWarning(ex,
                                "ListMailsRecursive ueberspringe Top-Level-Folder {Name}: unerwarteter Fehler",
                                name);
                        }
                    }
                    finally
                    {
                        if (folder is not null)
                        {
                            ReleaseComObject(folder);
                        }
                    }
                }
            }
            finally
            {
                _ = root;
            }

            collected.Sort((a, b) => DateTime.Compare(b.Received, a.Received));
            if (collected.Count > top)
            {
                collected.RemoveRange(top, collected.Count - top);
            }

            var slice = new List<MailMessage>(collected.Count);
            foreach (var t in collected) slice.Add(t.Msg);

            _logger.LogInformation(
                "ListMailsRecursive DONE returned={Returned} visitedFolders={Visited} " +
                "skippedFolders={Skipped} skippedNonMail={SkippedNonMail} failedItems={FailedItems} " +
                "restrictedHits={Restricted}",
                slice.Count, visitedFolders, skippedFolders, skippedNonMail,
                failedItems, restrictedItems);

            return new PagedResult<MailMessage>
            {
                Value = slice,
                NextSkip = null,
            };
        }, nameof(ListMailsRecursiveAsync));
    }

    private void CollectMailsFromFolderRecursive(
        dynamic parent,
        string? filter,
        int top,
        List<(string EntryId, DateTime Received, MailMessage Msg)> collected,
        HashSet<string> seen,
        ref int visitedFolders,
        ref int skippedFolders,
        ref int skippedNonMail,
        ref int failedItems,
        ref int restrictedItems)
    {
        dynamic? items = null;
        try
        {
            try { items = parent.Items; }
            catch (COMException ex)
            {
                skippedFolders++;
                _logger.LogWarning(ex,
                    "ListMailsRecursive ueberspringe Folder (HResult=0x{HResult:X8}): Items nicht lesbar",
                    ex.HResult);
                items = null;
            }
            catch (Exception ex)
            {
                skippedFolders++;
                _logger.LogWarning(ex,
                    "ListMailsRecursive ueberspringe Folder: parent.Items nicht aufrufbar");
                items = null;
            }
            if (items is not null)
            {
                visitedFolders++;
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    var originalItems = items;
                    items = RestrictMailItems((object)originalItems, filter);
                    if (!ReferenceEquals(originalItems, items)) ReleaseComObject(originalItems);
                }
                CollectItemsInto(items, top, collected, seen,
                    ref skippedNonMail, ref failedItems, ref restrictedItems);
            }
        }
        finally
        {
            if (items is not null)
            {
                ReleaseComObject(items);
            }
        }

        if (collected.Count >= top) return;

        dynamic? subFolders = null;
        try
        {
            try { subFolders = parent.Folders; }
            catch (COMException ex)
            {
                skippedFolders++;
                _logger.LogWarning(ex,
                    "ListMailsRecursive ueberspringe Subfolder-Liste (HResult=0x{HResult:X8})",
                    ex.HResult);
                subFolders = null;
            }
            if (subFolders is null) return;
            int subCount;
            try { subCount = (int)subFolders.Count; } catch { subCount = -1; }
            _logger.LogInformation(
                "ListMailsRecursive Subfolder-Walk: parent={Parent} subCount={SubCount}",
                (string)(TryGetString(parent, "Name") ?? "?"), subCount);
            foreach (var sub in subFolders)
            {
                try
                {
                    CollectMailsFromFolderRecursive(
                        sub,
                        filter,
                        top,
                        collected,
                        seen,
                        ref visitedFolders,
                        ref skippedFolders,
                        ref skippedNonMail,
                        ref failedItems,
                        ref restrictedItems);
                    if (collected.Count >= top) return;
                }
                finally
                {
                    ReleaseComObject(sub);
                }
            }
        }
        finally
        {
            if (subFolders is not null)
            {
                ReleaseComObject(subFolders);
            }
        }
    }

    internal static object RestrictMailItems(object collection, string filter)
    {
        try
        {
            return ((dynamic)collection).Restrict(filter);
        }
        catch (COMException ex)
        {
            throw new OutlookServiceException(ErrorCode.InvalidInput,
                "Outlook konnte den Mailfilter nicht anwenden", ex);
        }
    }

    private void CollectItemsInto(
        dynamic items,
        int top,
        List<(string EntryId, DateTime Received, MailMessage Msg)> collected,
        HashSet<string> seen,
        ref int skippedNonMail,
        ref int failedItems,
        ref int restrictedItems)
    {
        int total;
        try { total = (int)items.Count; } catch { total = 0; }
        for (int i = 1; i <= total; i++)
        {
            if (collected.Count >= top) return;
            dynamic? item = null;
            try
            {
                try { item = items.Item(i); }
                catch (COMException)
                {
                    failedItems++;
                    continue;
                }
                int itemClass = 0;
                try { itemClass = (int)item.Class; } catch { }
                if (itemClass != 43)
                {
                    skippedNonMail++;
                    continue;
                }
                string entryId;
                try { entryId = (string)item.EntryID; }
                catch
                {
                    failedItems++;
                    continue;
                }
                if (!seen.Add(entryId))
                {
                    restrictedItems++;
                    continue;
                }
                MailMessage msg;
                try
                {
                    msg = MapMailItem(item, null, _logger);
                }
                catch (COMException)
                {
                    failedItems++;
                    continue;
                }
                DateTime received;
                try { received = (DateTime)item.ReceivedTime; }
                catch { received = DateTime.MinValue; }
                collected.Add((entryId, received, msg));
            }
            finally
            {
                ReleaseComObject(item);
            }
        }
    }

    private void CollectMatchingMails(
        dynamic parent,
        string query,
        List<MailMessage> slice,
        int top,
        ref int visitedFolders,
        ref int skippedFolders,
        ref int skippedNonMail,
        ref int failedItems)
    {
        try
        {
            CollectMatchingMailsFromItems(parent, query, slice, top,
                ref skippedNonMail, ref failedItems);
            if (slice.Count >= top) return;
            CollectMatchingMailsFromSubfolders(parent, query, slice, top,
                ref visitedFolders, ref skippedFolders, ref skippedNonMail, ref failedItems);
        }
        catch (COMException ex)
        {
            skippedFolders++;
            _logger.LogWarning(ex,
                "SearchMails ueberspringe Ordner (HResult=0x{HResult:X8}): Zugriff fehlgeschlagen",
                ex.HResult);
        }
        catch (Exception ex)
        {
            skippedFolders++;
            _logger.LogWarning(ex, "SearchMails ueberspringe Ordner: unerwarteter Fehler");
        }
    }

    private void CollectMatchingMailsFromItems(
        dynamic parent,
        string query,
        List<MailMessage> slice,
        int top,
        ref int skippedNonMail,
        ref int failedItems)
    {
        dynamic items = parent.Items;
        try
        {
            if (items is null) return;
            foreach (var item in items)
            {
                try
                {
                    int itemClass = 0;
                    try { itemClass = (int)item.Class; } catch { }
                    if (itemClass != 43)
                    {
                        skippedNonMail++;
                        continue;
                    }
                    var subj = TryGetString(item, "Subject") ?? string.Empty;
                    var sender = TryGetString(item, "SenderEmailAddress") ?? string.Empty;
                    if (subj.Contains(query, StringComparison.OrdinalIgnoreCase)
                        || sender.Contains(query, StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            slice.Add(MapMailItem(item, null, _logger));
                            if (slice.Count >= top) return;
                        }
                        catch (COMException mapEx)
                        {
                            failedItems++;
                            _logger.LogWarning(mapEx,
                                "SearchMails ueberspringe Mail-Item (HResult=0x{HResult:X8}); " +
                                "MapMailItem fehlgeschlagen",
                                mapEx.HResult);
                        }
                        catch (Exception mapEx)
                        {
                            failedItems++;
                            _logger.LogError(mapEx,
                                "SearchMails unerwarteter Fehler beim Mapping eines Mail-Items");
                        }
                    }
                }
                catch (COMException itemEx)
                {
                    failedItems++;
                    _logger.LogWarning(itemEx,
                        "SearchMails ueberspringe Item (HResult=0x{HResult:X8}); Item nicht lesbar",
                        itemEx.HResult);
                }
                finally
                {
                    ReleaseComObject(item);
                }
            }
        }
        finally
        {
            ReleaseComObject(items);
        }
    }

    private void CollectMatchingMailsFromSubfolders(
        dynamic parent,
        string query,
        List<MailMessage> slice,
        int top,
        ref int visitedFolders,
        ref int skippedFolders,
        ref int skippedNonMail,
        ref int failedItems)
    {
        dynamic subFolders = parent.Folders;
        try
        {
            if (subFolders is null) return;
            foreach (var sub in subFolders)
            {
                try
                {
                    visitedFolders++;
                    CollectMatchingMails(sub, query, slice, top,
                        ref visitedFolders, ref skippedFolders, ref skippedNonMail, ref failedItems);
                    if (slice.Count >= top) return;
                }
                finally
                {
                    ReleaseComObject(sub);
                }
            }
        }
        finally
        {
            ReleaseComObject(subFolders);
        }
    }
}
