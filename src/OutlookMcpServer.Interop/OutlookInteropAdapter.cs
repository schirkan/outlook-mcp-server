using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
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
                StartOutlookExe();
                await WaitForOutlookReadyAsync(_options.Outlook.StartupTimeoutSeconds, ct);
            }

            InitializeMapiNamespace();
            return _outlookApp!;
        }
        finally
        {
            _comLock.Release();
        }
    }

    /// <summary>
    /// Startet outlook.exe (UseShellExecute=true, kein Window-Style gesetzt).
    /// Wird nur aufgerufen, wenn eine aktive Outlook-Application nicht gefunden
    /// wurde UND AutoStartOutlook aktiviert ist.
    /// </summary>
    private void StartOutlookExe()
    {
        _logger.LogInformation("Starte outlook.exe...");
        var psi = new ProcessStartInfo("outlook.exe") { UseShellExecute = true };
        Process.Start(psi);
    }

    /// <summary>
    /// Pollt bis zu <paramref name="timeoutSeconds"/> lang (500ms-Intervall) auf
    /// eine aktive Outlook-Application. Wirft OutlookServiceException(OutlookNotRunning)
    /// wenn der Timeout erreicht wird.
    /// </summary>
    private async Task WaitForOutlookReadyAsync(int timeoutSeconds, CancellationToken ct)
    {
        var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);
        Exception? lastEx = null;
        while (DateTime.UtcNow < deadline)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                _outlookApp = GetOrStartOutlookApplication();
                return;
            }
            catch (COMException ex)
            {
                lastEx = ex;
                await Task.Delay(500, ct);
            }
        }
        throw new OutlookServiceException(
            ErrorCode.OutlookNotRunning,
            $"Outlook konnte nicht innerhalb von {timeoutSeconds}s gestartet werden",
            lastEx);
    }

    /// <summary>
    /// Initialisiert _mapiNamespace aus der Outlook-Application. Setzt
    /// voraus, dass _outlookApp bereits gesetzt ist.
    /// </summary>
    private void InitializeMapiNamespace()
    {
        _mapiNamespace = _outlookApp!.GetNamespace("MAPI");
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
    private dynamic? GetFolderByIdOrWellKnownName(string folderIdOrName)
    {
        var ns = GetMapiNamespace();
        var olFolderId = OlEnumMappings.ToOlDefaultFolder(folderIdOrName);
        if (olFolderId.HasValue)
        {
            return ResolveDefaultFolderSmart(olFolderId.Value);
        }
        // Annahme: folderIdOrName ist eine EntryID
        try
        {
            return ns.GetFolderFromID(folderIdOrName, Type.Missing);
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <summary>
    /// Loest einen Well-Known-Folder per OlDefaultFolder-ID auf. In
    /// Multi-Store-Profilen (z. B. Cached-Mode/Exchange + zusaetzliches PST)
    /// liefert <c>ns.GetDefaultFolder(olId)</c> nicht zwingend den erwarteten
    /// Folder — Outlook nummeriert die olIds je nach Store-Reihenfolge anders.
    ///
    /// Strategie (robust, sprach- und profilunabhaengig):
    ///  1. Iteriere ueber <c>session.Stores</c> und versuche pro Store
    ///     <c>store.GetDefaultFolder(olId)</c>. Der erste Store, der die ID
    ///     unterstuetzt, gewinnt.
    ///  2. Wenn kein Store den Well-Known-Folder enthaelt: Fallback auf
    ///     <c>ns.GetDefaultFolder(olId)</c> (Originalverhalten).
    ///
    /// Liefert <c>null</c>, wenn nirgends etwas gefunden wurde.
    /// </summary>
    private dynamic? ResolveDefaultFolderSmart(int olFolderId)
    {
        // 1) session.Stores durchlaufen — Multi-Store-robust
        try
        {
            dynamic ns = GetMapiNamespace();
            dynamic session = ns.Session;
            dynamic stores = session.Stores;
            try
            {
                foreach (var store in stores)
                {
                    try
                    {
                        dynamic folder = store.GetDefaultFolder(olFolderId);
                        if (folder is not null)
                        {
                            return folder;
                        }
                    }
                    catch (COMException)
                    {
                        // Dieser Store kennt die olId nicht — naechster.
                    }
                }
            }
            finally
            {
                try { Marshal.ReleaseComObject(stores); } catch { }
            }
        }
        catch
        {
            // session/Stores nicht verfuegbar — gleich zu Schritt 2.
        }

        // 2) Fallback: ns.GetDefaultFolder (Originalverhalten)
        try
        {
            dynamic ns = GetMapiNamespace();
            return ns.GetDefaultFolder(olFolderId);
        }
        catch (COMException)
        {
            return null;
        }
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
    // ===== Mail Mapping Helper Records =====
    private record MailRecipients(IReadOnlyList<Recipient> To, IReadOnlyList<Recipient> Cc, IReadOnlyList<Recipient> Bcc);
    private record MailDates(DateTimeOffset? Sent, DateTimeOffset? Received);
    private record MailMetadata(bool HasAttachments, bool IsUnread, Importance Importance, List<string> Categories);

    // Outlook-OOM OlRecipientType:
    //   1 = olOriginator (der From-Absender)
    //   2 = olTo
    //   3 = olCC
    //   4 = olBCC
    private const int OlRecipientTo = 2;
    private const int OlRecipientCC = 3;
    private const int OlRecipientBCC = 4;

    private static MailMessage MapMailItem(dynamic mail, bool includeBody, BodyFormat bodyFormat, ILogger? logger = null)
    {
        var entryId = (string)mail.EntryID;

        var body = TryMapMailBody(mail, includeBody, bodyFormat, logger);
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
    /// Liest den Mail-Body (Text oder HTML). Gibt null zurueck, wenn
    /// <paramref name="includeBody"/>=false (ListMails-Use-Case).
    /// BodyFormat MUSS vor Body/HTMLBody gelesen werden.
    /// <para>
    /// Falls der Body leer erscheint, wird zusaetzlich BodyPreview geliefert,
    /// damit der Client wenigstens eine Vorschau bekommt. COM-Fehler beim
    /// Body-Zugriff werden geloggt (LogMessage wurde zuvor im gesamten
    /// Code-Pfad stillschweigend geschluckt — das war einer der Hauptgruende
    /// fuer die spaerliche Body-Ausgabe).
    /// </para>
    /// </summary>
    private static ItemBody? TryMapMailBody(dynamic mail, bool includeBody, BodyFormat bodyFormat, ILogger? logger = null)
    {
        if (!includeBody) return null;
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

            // Native Outlook-Quelle priorisieren: HTML wenn vorhanden, sonst Text.
            string? nativeSource = null;
            if (!string.IsNullOrEmpty(html))
            {
                nativeSource = html;
            }
            else if (!string.IsNullOrEmpty(text))
            {
                nativeSource = text;
            }

            // Konvertierung in das gewuenschte Zielformat.
            // Bei Html-Request wird der native HTML durchgereicht (1:1).
            // Bei Markdown/Text wird der native HTML konvertiert (auch wenn Outlook
            // nur Text liefert: in diesem Fall wickelt der Converter in <html><body>).
            if (nativeSource is null || nativeSource.Length == 0)
            {
                // Body leer / nicht lesbar → Fallback auf BodyPreview,
                // damit der Client wenigstens eine Vorschau bekommt.
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

            // Wenn Outlook Text liefert und HTML gewuenscht ist, in HTML wrappen
            // (selten, aber konsistent).
            string htmlForConversion = nativeSource == text
                ? $"<html><body><pre>{System.Net.WebUtility.HtmlEncode(text)}</pre></body></html>"
                : nativeSource;

            var (ct, content) = HtmlBodyConverter.Convert(htmlForConversion, bodyFormat);
            return new ItemBody { ContentType = ct, Content = content };
        }
        catch (Exception ex)
        {
            logger?.LogWarning(ex, "TryMapMailBody: Outer try fehlgeschlagen");
        }
        return null;
    }

    /// <summary>
    /// Liest den Sender (Name + SMTP-Adresse). Gibt null zurueck, wenn
    /// Sender leer oder nicht lesbar.
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

    private static MailRecipients TryMapMailRecipients(dynamic mail)
    {
        // Outlook-OOM MailItem hat KEINE separaten To/CC/BCC-Properties — nur eine
        // einheitliche Recipients-Collection, in der jedes Recipient-Element eine
        // Type-Property (2=To, 3=CC, 4=BCC) hat. Der vorherige Versuch, To/CC/BCC
        // per TryGetDynamic zu lesen, hat daher leere Arrays geliefert.
        // Strategie: Recipients einmal holen, dann nach Type filtern.
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
                    // Bei Type=1 (olOriginator) handelt es sich um den From-Absender —
                    // nicht in die Empfänger-Listen aufnehmen, der wird separat
                    // in MapMailFrom() gelesen.
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
                        // Unbekannte Typen (oder Originator) werden ignoriert.
                    }
                }
                finally
                {
                    if (rec is not null) Marshal.ReleaseComObject(rec);
                }
            }
        }
        catch { }
        finally
        {
            try { Marshal.ReleaseComObject(recipients); } catch { }
        }
        return new MailRecipients(to, cc, bcc);
    }

    private static MailDates TryMapMailDates(dynamic mail)
    {
        DateTimeOffset? sent = null;
        DateTimeOffset? received = null;
        try { sent = (DateTimeOffset)mail.SentOn; } catch { }
        try { received = (DateTimeOffset)mail.ReceivedTime; } catch { }
        return new MailDates(sent, received);
    }

    /// <summary>
    /// Liest HasAttachments, IsUnread (Outlook "UnRead" invertiert zu DTO "IsRead"),
    /// Importance und Categories in einem Rutsch.
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

    /// <summary>
    /// Defensive Property-Access-Wrapper fuer COM dynamic-Properties.
    /// Beachte: bei einem <c>System.__ComObject</c>-RCW liefert
    /// <c>obj.GetType().GetProperty(name)</c> NICHT die echten COM/IDispatch-Properties
    /// (Subject/BodyPreview/To/CC/BCC/etc.) — sondern nur die Wrapper-Typ-Properties.
    /// Wir gehen daher den direkten <c>dynamic</c>-Property-Pfad ueber den
    /// C#-dynamic-Dispatcher (welcher intern IDispatch::Invoke nutzt).
    /// </summary>
    private static string? TryGetString(dynamic obj, string propertyName)
    {
        try
        {
            var value = GetDynamicProperty(obj, propertyName);
            return value as string;
        }
        catch { return null; }
    }

    private static dynamic? TryGetDynamic(dynamic obj, string propertyName)
    {
        // Direkter dynamic-Property-Zugriff ueber C#-dynamic-Binder (IDispatch::Invoke).
        try
        {
            return GetDynamicProperty(obj, propertyName);
        }
        catch { return null; }
    }

    /// <summary>
    /// Loest eine COM-Property via IDispatch auf. Da der Parameter dynamic ist,
    /// ruft der C#-Compiler den late-bound GetMember-Binder auf, welcher intern
    /// IDispatch::Invoke (DISPATCH_PROPERTYGET) ausfuehrt. Das ist der einzige
    /// Weg, um Outlook-Properties wie Subject, BodyPreview, To/CC/BCC
    /// (Recipients-Collections) auf einem System.__ComObject-RCW zu lesen —
    /// GetType().GetProperty(name) liefert nur den RCW-Wrapper.
    /// </summary>
    private static object? GetDynamicProperty(dynamic obj, string name)
    {
        var binder = Microsoft.CSharp.RuntimeBinder.Binder.GetMember(
            CSharpBinderFlags.None,
            name,
            null,
            new[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) });
        var callsite = CallSite<Func<CallSite, object, object>>.Create(binder);
        return callsite.Target(callsite, (object)obj);
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
            dynamic folder = GetFolderByIdOrWellKnownName(folderId)
                ?? throw new OutlookServiceException(
                    ErrorCode.FolderNotFound,
                    $"Mail-Ordner nicht gefunden: '{folderId}'. " +
                    "Bei Well-Known-Namen (inbox, drafts, sentItems, deletedItems, " +
                    "junkEmail, archive, outbox) wird in Multi-Store-Profilen " +
                    "unter allen Stores gesucht — falls keiner passt, ist der " +
                    "Folder im aktiven Profil nicht vorhanden.");
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
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        _logger.LogInformation(
            "ListMails START folderId={FolderId} top={Top} skip={Skip} filter={Filter} search={Search}",
            folderId, top, skip, filter, search);
        return await RunComAsync(() =>
        {
            dynamic folder = GetFolderByIdOrWellKnownName(folderId)
                ?? throw new OutlookServiceException(
                    ErrorCode.FolderNotFound,
                    $"Mail-Ordner nicht gefunden: '{folderId}'.");
            string? folderName = null;
            try { folderName = (string)folder.Name; } catch { /* defensive */ }
            _logger.LogInformation(
                "ListMails folder resolved: displayName={FolderName} folderId={FolderId}",
                folderName, folderId);
            try
            {
                string combinedFilter = CombineDaslFilters(filter, search);
                dynamic items = folder.Items;
                int rawCount = 0;
                try { rawCount = (int)items.Count; } catch { }
                _logger.LogInformation(
                    "ListMails raw item count before Restrict: {RawCount} (folder={FolderName})",
                    rawCount, folderName);

                // Wir versuchen zunaechst einen Restrict mit dem kombinierten
                // Filter. Wenn dieser fehlschlaegt (z. B. wegen unsupported
                // DASL-Property auf einem Exchange-Item oder wegen eines
                // Sonderzeichens im Such-Query), fallen wir auf eine
                // unguelfilterte Iteration mit nachtraeglicher In-Memory-Pruefung
                // zurueck. So bekommen wir zumindest ein Ergebnis, statt eines
                // generischen Fehlers.
                try
                {
                    if (!string.IsNullOrWhiteSpace(combinedFilter))
                    {
                        items = items.Restrict(combinedFilter);
                        int restrictedCount = (int)items.Count;
                        _logger.LogInformation(
                            "ListMails Restrict OK: {Count} items nach Filter (folder={FolderName})",
                            restrictedCount, folderName);
                    }
                }
                catch (COMException restrictEx)
                {
                    _logger.LogWarning(restrictEx,
                        "ListMails Restrict fehlgeschlagen (HResult=0x{HResult:X8}); " +
                        "Fallback auf ungefilterte Iteration. filter={Filter} search={Search}",
                        restrictEx.HResult, filter, search);
                    // Items erneut holen, da Restrict-Fehler den Items-Zustand
                    // undefiniert hinterlassen kann.
                    try { Marshal.ReleaseComObject(items); } catch { }
                    items = folder.Items;
                }
                try
                {
                    items.Sort("[ReceivedTime]", true); // newest first

                    int total = (int)items.Count;
                    int start = Math.Min(skip, total);
                    int end = Math.Min(skip + top, total);
                    var slice = new List<MailMessage>(end - start);
                    int skippedNonMail = 0;
                    int failedItems = 0;
                    for (int i = start; i < end; i++)
                    {
                        dynamic item;
                        try
                        {
                            // COM-Auflistungen sind 1-basiert
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
                            // Class-Filter: nur olMail (43). Outlook-Ordner koennen
                            // AppointmentItem (1), TaskItem (2), ContactItem (8),
                            // PostItem (4), ReportItem (5) u. a. enthalten. Ohne
                            // diesen Filter schlaegt MapMailItem fehl, sobald ein
                            // Item eines unerwarteten Typs vorhanden ist.
                            int itemClass = 0;
                            try { itemClass = (int)item.Class; } catch { }
                            if (itemClass != 43)
                            {
                                skippedNonMail++;
                                continue;
                            }
                            slice.Add(MapMailItem(item, false, BodyFormat.Markdown, _logger));
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
                                "ListMails unerwarteter Fehler beim Mapping von Item {Index} " +
                                "im Folder {FolderName}", i + 1, folderName);
                        }
                        finally
                        {
                            try { Marshal.ReleaseComObject(item); } catch { }
                        }
                    }
                    _logger.LogInformation(
                        "ListMails DONE folder={FolderName} returned={Returned} skippedNonMail={SkippedNonMail} " +
                        "failedItems={FailedItems} total={Total} start={Start} end={End}",
                        folderName, slice.Count, skippedNonMail, failedItems, total, start, end);
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

    public async Task<MailMessage> GetMailAsync(
        string id,
        bool includeBody = true,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic mail = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                return MapMailItem(mail, includeBody, bodyFormat, _logger);
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
            }
        }, nameof(GetMailAsync));
    }

    public async Task<BulkMailResult> GetMailsAsync(
        IReadOnlyList<string> ids,
        bool includeBody = false,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        // Bulk-Operation: per-ID-Errors werden nicht als Top-Level-Exception geworfen,
        // sondern in BulkMailResult.NotFoundIds gesammelt. Direkter Zugriff auf
        // den MAPI-Namespace ohne RunComAsync-Wrap, weil wir jede COMException
        // pro ID abfangen und nicht in OutlookServiceException umgewandelt haben wollen.
        var ns = GetMapiNamespace();
        var found = new List<MailMessage>(ids.Count);
        var notFound = new List<string>();
        foreach (var id in ids)
        {
            dynamic? mail = null;
            try
            {
                mail = ns.GetItemFromID(id, Type.Missing);
                found.Add(MapMailItem(mail, includeBody, bodyFormat));
            }
            catch (COMException ex)
            {
                // Item nicht gefunden (MAPI_E_NOT_FOUND) oder andere COM-Fehler
                // -> notFound-Liste. Bulk-Semantik ist per-ID-tolerant.
                _logger.LogInformation(
                    "GetMails: ID '{Id}' nicht aufloesbar (HResult=0x{HResult:X8}: {Msg})",
                    id, ex.HResult, ex.Message);
                notFound.Add(id);
            }
            finally
            {
                if (mail is not null)
                {
                    try { Marshal.ReleaseComObject(mail); } catch { /* defensive */ }
                }
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
                // MAPI_E_NOT_FOUND: Property nicht vorhanden -> leere Liste (manche Mails haben keine Internet-Header)
                _logger.LogInformation("Mail {Id} hat keine Internet-Header (PropertyAccess nicht vorhanden)", id);
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
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
                        // Inline-Detection: Outlook gibt Position 0 fuer
                        // regular Attachments, >0 fuer inline (HTML-Body-embedded).
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
                        Marshal.ReleaseComObject(attachment);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
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
                            Marshal.ReleaseComObject(attachment);
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
                    Marshal.ReleaseComObject(mail);
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

            // Wenn folderId null → rekursiv ueber alle Mail-Ordner durchsuchen.
            // Outlook.NameSpace hat KEIN .Items, nur .Folders — daher iterieren
            // wir erst die Stores, dann pro Store die Top-Level-Folder-Auflistung.
            if (folderId is null)
            {
                dynamic root = GetMapiNamespace();
                // root wird NICHT per Marshal.ReleaseComObject freigegeben —
                // GetMapiNamespace gibt das gecachte _mapiNamespace-Feld zurueck;
                // ein Release hier wuerde die Refcount-Balance brechen und
                // Folge-Calls mit InvalidComObjectException scheitern lassen.
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
                                                try { Marshal.ReleaseComObject(f); } catch { }
                                            }
                                        }
                                    }
                                    finally { Marshal.ReleaseComObject(topFolders); }
                                }
                                finally { Marshal.ReleaseComObject(storeRoot); }
                            }
                            finally { Marshal.ReleaseComObject(store); }
                        }
                    }
                    finally { Marshal.ReleaseComObject(stores); }
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
                    Marshal.ReleaseComObject(folder);
                }
            }

            _logger.LogInformation(
                "SearchMails DONE query={Query} returned={Returned} visitedFolders={Visited} " +
                "skippedFolders={SkippedFolders} skippedNonMail={SkippedNonMail} failedItems={FailedItems}",
                query, slice.Count, visitedFolders, skippedFolders, skippedNonMail, failedItems);

            return new PagedResult<MailMessage>
            {
                Value = slice,
                NextSkip = null, // Search liefert keine Pagination (Hard-Cap via top)
            };
        }, nameof(SearchMailsAsync));
    }

    public async Task<PagedResult<MailMessage>> ListMailsRecursiveAsync(
        IReadOnlyList<string> scope,
        int top,
        string? filter,
        BodyFormat bodyFormat,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        _logger.LogInformation(
            "ListMailsRecursive START scope=[{Scope}] top={Top} filter={Filter}",
            string.Join(",", scope), top, filter);
        return await RunComAsync(() =>
        {
            // scope = leer -> Default = alle erlaubten Well-Known-Mailordner.
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

            // Sammelt Treffer aller Folder, dedupliziert per EntryID und capped bei top.
            var collected = new List<(string EntryId, DateTime Received, MailMessage Msg)>(top * 2);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            int visitedFolders = 0, skippedFolders = 0, skippedNonMail = 0,
                failedItems = 0, restrictedItems = 0, fallbackFolders = 0;

            dynamic root = GetMapiNamespace();
            // root wird NICHT per Marshal.ReleaseComObject freigegeben — siehe
            // Kommentar in SearchMailsAsync und im Methoden-Body oben.
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
                            // Multi-Store-robust: iteriert session.Stores und nimmt
                            // den ersten Store, der die olId unterstuetzt. Fallback
                            // auf ns.GetDefaultFolder ist im Helper enthalten.
                            folder = ResolveDefaultFolderSmart(ol.Value);
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
                                filter,
                                top,
                                bodyFormat,
                                collected,
                                seen,
                                ref visitedFolders,
                                ref skippedFolders,
                                ref skippedNonMail,
                                ref failedItems,
                                ref restrictedItems,
                                ref fallbackFolders);
                            _logger.LogInformation(
                                "ListMailsRecursive Top-Level {Name} fertig: collected={Collected}/{Top} visitedFolders={Visited}",
                                name, collected.Count, top, visitedFolders);
                            if (collected.Count >= top) break;
                        }
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
                            try { Marshal.ReleaseComObject(folder); } catch { }
                        }
                    }
                }
            }
            finally
            {
                // Hier ist nichts zusaetzlich zum root releasen.
                _ = root;
            }

            // Sort neueste zuerst, dann auf top capen.
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
                "restrictedHits={Restricted} fallbackFolders={Fallback}",
                slice.Count, visitedFolders, skippedFolders, skippedNonMail,
                failedItems, restrictedItems, fallbackFolders);

            return new PagedResult<MailMessage>
            {
                Value = slice,
                NextSkip = null,
            };
        }, nameof(ListMailsRecursiveAsync));
    }

    /// <summary>
    /// Sammelt Mails aus einem Folder (rekursiv in Unterordner), wendet
    /// <paramref name="filter"/> per <c>Items.Restrict()</c> an, dedupliziert
    /// per EntryID. Bricht ab, sobald <c>collected.Count &gt;= top</c>.
    /// </summary>
    private void CollectMailsFromFolderRecursive(
        dynamic parent,
        string? filter,
        int top,
        BodyFormat bodyFormat,
        List<(string EntryId, DateTime Received, MailMessage Msg)> collected,
        HashSet<string> seen,
        ref int visitedFolders,
        ref int skippedFolders,
        ref int skippedNonMail,
        ref int failedItems,
        ref int restrictedItems,
        ref int fallbackFolders)
    {
        // 1) Items dieses Folders sammeln.
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
                dynamic? restricted = null;
                try
                {
                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        restricted = items.Restrict(filter!);
                        var c = (int)restricted.Count;
                        _logger.LogInformation(
                            "ListMailsRecursive Restrict OK: folder={FolderName} hits={Hits}",
                            (string)(TryGetString(parent, "Name") ?? "?"), c);
                    }
                    else
                    {
                        restricted = items;
                    }
                }
                catch (COMException restrictEx)
                {
                    fallbackFolders++;
                    _logger.LogWarning(restrictEx,
                        "ListMailsRecursive Restrict fehlgeschlagen (HResult=0x{HResult:X8}); " +
                        "Fallback auf ungefilterte Iteration. filter={Filter}",
                        restrictEx.HResult, filter);
                    try { Marshal.ReleaseComObject(items); } catch { }
                    try { items = parent.Items; } catch { items = null; }
                    restricted = items;
                }
                if (restricted is not null)
                {
                    CollectItemsInto(restricted, top, bodyFormat, collected, seen,
                        ref skippedNonMail, ref failedItems, ref restrictedItems);
                    try { Marshal.ReleaseComObject(restricted); } catch { }
                }
            }
        }
        finally
        {
            if (items is not null)
            {
                try { Marshal.ReleaseComObject(items); } catch { }
            }
        }

        if (collected.Count >= top) return;

        // 2) Rekursion in Unterordner.
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
                        bodyFormat,
                        collected,
                        seen,
                        ref visitedFolders,
                        ref skippedFolders,
                        ref skippedNonMail,
                        ref failedItems,
                        ref restrictedItems,
                        ref fallbackFolders);
                    if (collected.Count >= top) return;
                }
                finally
                {
                    try { Marshal.ReleaseComObject(sub); } catch { }
                }
            }
        }
        finally
        {
            if (subFolders is not null)
            {
                try { Marshal.ReleaseComObject(subFolders); } catch { }
            }
        }
    }

    /// <summary>
    /// Iteriert ueber <paramref name="items"/> (1-basiert), mappt jedes olMail-Item
    /// und nimmt es in <paramref name="collected"/> auf, sofern die EntryID
    /// noch nicht in <paramref name="seen"/> liegt. Bricht ab, sobald
    /// <c>collected.Count &gt;= top</c>.
    /// </summary>
    private void CollectItemsInto(
        dynamic items,
        int top,
        BodyFormat bodyFormat,
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
                    msg = MapMailItem(item, false, BodyFormat.Markdown, _logger);
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
                if (item is not null)
                {
                    try { Marshal.ReleaseComObject(item); } catch { }
                }
            }
        }
    }

    /// <summary>
    /// Iteriert ueber einen Folder (oder Root) und sammelt Mails, deren
    /// Subject ODER SenderEmailAddress den Query (case-insensitive) enthaelt.
    /// Bricht ab, sobald <c>slice.Count == top</c> erreicht ist. Rekursion
    /// ueber Unterordner. Com-Objekte werden pro Item freigegeben.
    /// </summary>
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

    /// <summary>
    /// Iteriert die Items des Folders und sammelt Mails, deren Subject oder
    /// SenderEmailAddress den Query (case-insensitive) enthalten. Bricht ab,
    /// sobald slice.Count >= top.
    /// </summary>
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
                    // Class-Filter: nur olMail (43). Outlook-Ordner koennen
                    // Reports, Appointments, Posts etc. enthalten.
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
                            slice.Add(MapMailItem(item, false, BodyFormat.Markdown, _logger));
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
                    try { Marshal.ReleaseComObject(item); } catch { }
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(items);
        }
    }

    /// <summary>
    /// Rekursion in Unterordner. Bricht ab, sobald slice.Count >= top
    /// (Abbruch-Signal wird ueber den return propagierenden Aufrufer geprueft).
    /// </summary>
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
                    try { Marshal.ReleaseComObject(sub); } catch { }
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(subFolders);
        }
    }

    // ===== Mail: Mutationen =====

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
                // SaveToSentItems: MailItem.DeleteAfterSubmit = !SaveToSentItems
                // (Outlook-Default ist false -> Mail landet in SentItems)
                mail.DeleteAfterSubmit = !request.SaveToSentItems;
                // Deferred delivery (optional, Exchange-Accounts)
                if (request.SendAt is { } sendAt)
                {
                    mail.DeferredDeliveryTime = sendAt.UtcDateTime;
                }
                mail.Send();
                // .Send() verschiebt das Item in den Outbox/SentItems-Pfad;
                // EntryID ist danach nicht mehr stabil verfuegbar.
                return new SendMailResult { Sent = true, Id = null };
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
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
                Marshal.ReleaseComObject(mail);
            }
        }, nameof(CreateDraftAsync));
    }

    /// <summary>
    /// Erstellt ein MailItem aus einem <see cref="SendMailRequest"/>. Behandelt
    /// <c>replyToId</c> (MailItem.Reply / .ReplyAll), <c>forwardFromId</c>
    /// (MailItem.Forward) und Plain-New-Mail (Application.CreateItem(olMailItem)).
    /// Setzt To/Cc/Bcc (Outlook-Syntax: Semikolon-getrennt), Subject (nur bei
    /// neuer Mail; Outlook generiert Subject bei Reply/Forward), Body (HTML oder
    /// Plain), Importance und Attachments (Base64 -&gt; Temp-File).
    /// Setzt <c>_outlookApp</c>/<c>_mapiNamespace</c> via vorherigem
    /// <see cref="GetOutlookApplicationAsync"/> voraus.
    /// </summary>
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
            // Falls mail schon erstellt, aber ein Folgeschritt fehlgeschlagen ist:
            // releasen, sonst Memory-Leak. Caller hat noch keine Referenz.
            if (mail is not null)
            {
                try { Marshal.ReleaseComObject(mail); } catch { }
            }
            throw;
        }
        finally
        {
            if (sourceItem is not null)
            {
                try { Marshal.ReleaseComObject(sourceItem); } catch { }
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

    /// <summary>
    /// Erstellt das Mail-Item gemaess Request: Forward (sourceItem wird geladen),
    /// Reply / ReplyAll (sourceItem wird geladen), oder neue Mail. sourceItem
    /// wird via out-Parameter zurueckgegeben, damit das aeussere try/finally
    /// es sauber releasen kann.
    /// </summary>
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
        return _outlookApp!.CreateItem(0); // 0 = OlItemType.olMailItem
    }

    /// <summary>
    /// Schreibt Subject, To/Cc/Bcc, Body (mit korrekter BodyFormat-Reihenfolge),
    /// Importance und Attachments in das Mail-Item. Subject wird bei Reply/Forward
    /// NICHT gesetzt (Outlook generiert "Re:" / "Fwd:" Prefix automatisch).
    /// </summary>
    private void ApplyMailProperties(dynamic mail, SendMailRequest request, dynamic? sourceItem)
    {
        if (sourceItem is null && !string.IsNullOrEmpty(request.Subject))
        {
            mail.Subject = request.Subject;
        }

        if (request.To.Count > 0) mail.To = string.Join("; ", request.To);
        if (request.Cc.Count > 0) mail.Cc = string.Join("; ", request.Cc);
        if (request.Bcc.Count > 0) mail.Bcc = string.Join("; ", request.Bcc);

        // BodyFormat MUSS vor dem Body-Set gesetzt werden, sonst rendert
        // Outlook die Inhalte in das falsche Format.
        var bodyFormat = OlEnumMappings.ToOlBodyFormat(request.Body.ContentType);
        mail.BodyFormat = bodyFormat;
        if (bodyFormat == 2 /* OlBodyFormat.olFormatHTML */)
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

    /// <summary>
    /// Fuegt Inline-Attachments (Base64-codiert) zu einem MailItem hinzu.
    /// Schreibt jedes Attachment in eine temporaere Datei (Outlook braucht
    /// einen Pfad fuer <c>Attachments.Add</c>), ruft <c>Attachments.Add(path,
    /// OlAttachmentType.olByValue, ...)</c> auf und loescht die Temp-Datei
    /// danach (Outlook kopiert den Inhalt in den Message-Store).
    /// </summary>
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
                // Outlook liest die Datei beim .Add() und kopiert sie in den
                // Message-Store. Source-Datei ist danach nicht mehr noetig.
                mail.Attachments.Add(tempPath, 1 /* OlAttachmentType.olByValue */, Type.Missing, att.Name);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { /* defensive */ }
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
                // PATCH-Semantik: nur gesetzte Felder anfassen.
                if (update.IsRead is { } isRead)
                {
                    // Outlook: UnRead = true heisst ungelesen -> DTO IsRead = !UnRead
                    mail.UnRead = !isRead;
                }
                if (update.Importance is { } importance)
                {
                    mail.Importance = OlEnumMappings.ToOlImportance(importance);
                }
                if (update.Categories is not null)
                {
                    // Outlook: Categories ist ein Komma-getrennter String.
                    // Leere Liste loescht alle Kategorien (""), null-Felder werden uebersprungen.
                    mail.Categories = string.Join(", ", update.Categories);
                }
                mail.Save();
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
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
                    // MailItem.Move() gibt das verschobene Item zurueck (mit neuer EntryID).
                    dynamic moved = mail.Move(destFolder);
                    try
                    {
                        return (string)moved.EntryID;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(moved);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(destFolder);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
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
                // Outlook COM hat kein direktes CopyTo(folder) auf MailItem.
                // Copy() legt eine Kopie im selben Ordner an, danach Move() in den Zielordner.
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
                            Marshal.ReleaseComObject(moved);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(destFolder);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(copy);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
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
                    // Permanent delete: MailItem.Delete() allein reicht nicht, weil
                    // Outlook bei Items ausserhalb von DeletedItems in der Regel
                    // nur einen Soft-Delete macht. Erst in DeletedItems verschieben,
                    // dann dort loeschen (-> kein Recovery mehr moeglich).
                    dynamic deletedItems = GetMapiNamespace().GetDefaultFolder(5 /* OlDefaultFolders.olFolderDeletedItems */);
                    try
                    {
                        dynamic inDeleted = mail.Move(deletedItems);
                        try
                        {
                            inDeleted.Delete();
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(inDeleted);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(deletedItems);
                    }
                }
                else
                {
                    // Soft delete: in DeletedItems verschieben.
                    dynamic deletedItems = GetMapiNamespace().GetDefaultFolder(5 /* OlDefaultFolders.olFolderDeletedItems */);
                    try
                    {
                        mail.Move(deletedItems);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(deletedItems);
                    }
                }
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
            }
        }, nameof(DeleteMailAsync));
    }

    // ===== Calendar: Kalender =====

    public async Task<IReadOnlyList<Calendar>> ListCalendarsAsync(
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var result = new List<Calendar>();
            dynamic session = GetMapiNamespace();
            try
            {
                // Default-Calendar (typischerweise vom Default-Store)
                string? defaultCalendarEntryId = null;
                try
                {
                    dynamic defaultCal = session.GetDefaultFolder(9 /* OlDefaultFolders.olFolderCalendar */);
                    try { defaultCalendarEntryId = (string)defaultCal.EntryID; } catch { }
                    Marshal.ReleaseComObject(defaultCal);
                }
                catch { /* kein Default-Calendar vorhanden */ }

                // Alle Stores durchlaufen (Exchange, PST/OST, etc.)
                foreach (var store in session.Stores)
                {
                    dynamic? storeCal = null;
                    try
                    {
                        try
                        {
                            storeCal = store.GetDefaultFolder(9);
                        }
                        catch
                        {
                            // Store hat keinen Kalender (z. B. reiner Mail-Store)
                            continue;
                        }

                        var isDefault = false;
                        try { isDefault = (string)storeCal.EntryID == defaultCalendarEntryId; } catch { }
                        var entryId = (string)storeCal.EntryID;
                        if (!result.Any(c => c.Id == entryId))
                        {
                            result.Add(MapCalendar(storeCal, isDefault));
                        }

                        // Unter-Kalender (Custom-Calendars im selben Store)
                        foreach (var sub in storeCal.Folders)
                        {
                            try
                            {
                                int defaultItemType = 0;
                                try { defaultItemType = (int)sub.DefaultItemType; } catch { }
                                if (defaultItemType == 1 /* OlItemType.olAppointmentItem */)
                                {
                                    var subEntryId = (string)sub.EntryID;
                                    if (!result.Any(c => c.Id == subEntryId))
                                    {
                                        result.Add(MapCalendar(sub, false));
                                    }
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(sub);
                            }
                        }
                    }
                    finally
                    {
                        if (storeCal is not null) Marshal.ReleaseComObject(storeCal);
                        Marshal.ReleaseComObject(store);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(session);
            }
            return (IReadOnlyList<Calendar>)result;
        }, nameof(ListCalendarsAsync));
    }

    public async Task<Calendar> GetCalendarAsync(
        string id,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic folder = GetMapiNamespace().GetFolderFromID(id, Type.Missing);
            try
            {
                // Pruefen ob es ein Kalender-Ordner ist
                int defaultItemType = 0;
                try { defaultItemType = (int)folder.DefaultItemType; } catch { }
                if (defaultItemType != 1 /* olAppointmentItem */)
                {
                    throw new OutlookServiceException(
                        ErrorCode.CalendarNotFound,
                        $"Ordner mit EntryID '{id}' ist kein Kalender-Ordner (DefaultItemType={defaultItemType})");
                }
                // IsDefaultCalendar-Bestimmung waere Vergleich mit Default-Calendar
                // - hier nicht noetig, da Aufrufer explizit eine ID angibt
                return MapCalendar(folder, isDefault: false);
            }
            finally
            {
                Marshal.ReleaseComObject(folder);
            }
        }, nameof(GetCalendarAsync));
    }

    // ===== Calendar: Termine =====

    public async Task<PagedResult<CalendarEvent>> ListEventsAsync(
        string? calendarId,
        DateTimeTimeZone start,
        DateTimeTimeZone end,
        int top = 50,
        int skip = 0,
        string? filter = null,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic calendar = calendarId is null
                ? GetMapiNamespace().GetDefaultFolder(9 /* OlDefaultFolders.olFolderCalendar */)
                : GetMapiNamespace().GetFolderFromID(calendarId, Type.Missing);
            try
            {
                dynamic items = calendar.Items;
                try
                {
                    // Wichtig: IncludeRecurrences = true expandiert Serien in einzelne Vorkommen.
                    items.IncludeRecurrences = true;

                    // DASL-Filter (JET-Query-Syntax): Datum im US-Format M/d/yyyy h:mm:ss tt.
                    var startDt = DateTime.Parse(start.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                    var endDt = DateTime.Parse(end.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                    var startStr = startDt.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    var endStr = endDt.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    var combined = $"[Start] >= '{startStr}' AND [End] <= '{endStr}'";
                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        combined = $"({combined}) AND ({filter})";
                    }
                    items = items.Restrict(combined);
                    items.Sort("[Start]");

                    int total = (int)items.Count;
                    int startIdx = Math.Min(skip, total);
                    int endIdx = Math.Min(skip + top, total);
                    var slice = new List<CalendarEvent>(endIdx - startIdx);
                    for (int i = startIdx; i < endIdx; i++)
                    {
                        dynamic item = items.Item(i + 1);
                        try
                        {
                            slice.Add(MapAppointmentItem(item, bodyFormat));
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(item);
                        }
                    }
                    return new PagedResult<CalendarEvent>
                    {
                        Value = slice,
                        NextSkip = endIdx < total ? endIdx : (int?)null,
                    };
                }
                finally
                {
                    Marshal.ReleaseComObject(items);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(calendar);
            }
        }, nameof(ListEventsAsync));
    }

    public async Task<CalendarEvent> GetEventAsync(
        string id,
        BodyFormat bodyFormat = BodyFormat.Markdown,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            dynamic item = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                return MapAppointmentItem(item, bodyFormat);
            }
            finally
            {
                Marshal.ReleaseComObject(item);
            }
        }, nameof(GetEventAsync));
    }

    // ===== Calendar: Mapping-Helpers =====

    /// <summary>
    /// Mappt ein COM-MAPIFolder (Kalender-Ordner) auf das <see cref="Calendar"/>-DTO.
    /// Owner wird defensiv mit Fallback null gemappt (PropertyAccess via
    /// store.PropertyAccessor fuer PR_OWNER_NAME_W haengt vom Store-Typ ab).
    /// </summary>
    private static Calendar MapCalendar(dynamic folder, bool isDefault)
    {
        var entryId = (string)folder.EntryID;
        var name = (string)folder.Name;
        bool canEdit = true;
        try { canEdit = (int)folder.Permission != 0 /* olPermissionNone */; } catch { }
        string? owner = null;
        try
        {
            // PR_OWNER_NAME_W (0x3A1F001F als String) via Parent-Store
            dynamic parentStore = folder.Parent;
            try
            {
                if (parentStore is not null)
                {
                    dynamic propAcc = parentStore.PropertyAccessor;
                    try
                    {
                        owner = (string?)propAcc.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A1F001F");
                    }
                    finally { Marshal.ReleaseComObject(propAcc); }
                }
            }
            finally { Marshal.ReleaseComObject(parentStore); }
        }
        catch { /* kein Owner verfuegbar */ }

        return new Calendar
        {
            Id = entryId,
            Name = name,
            IsDefaultCalendar = isDefault,
            CanEdit = canEdit,
            Owner = owner,
        };
    }

    /// <summary>
    /// Mappt ein COM-AppointmentItem auf das <see cref="CalendarEvent"/>-DTO.
    /// Felder werden defensiv gemappt (Outlook-Werte koennen je nach Item-Typ
    /// fehlen, z. B. Body bei Terminen ohne Beschreibung, Organizer bei selbst
    /// angelegten Terminen ohne eingeladene Teilnehmer).
    /// <para>
    /// TimeZone-Hinweis: Outlook liefert nur Windows-TimeZone-IDs (z. B.
    /// "W. Europe Standard Time"), das DTO erwartet aber IANA-Namen (z. B.
    /// "Europe/Berlin"). Wir geben die Windows-ID unveraendert weiter; eine
    /// spaetere Phase koennte eine Mapping-Tabelle fuer haeufige Zonen ergaenzen.
    /// </para>
    /// </summary>
    // ===== Appointment Mapping Helper Records =====
    // Records statt Tuples, weil C# bei dynamic-Rueckgabe keine
    // var-Dekonstruktion ableiten kann (CS8133/CS8130).
    private record AppointmentStartEnd(DateTimeTimeZone? Start, DateTimeTimeZone? End);
    private record AppointmentLocation(Location? Location, List<Location> Locations);
    private record AppointmentFlags(Importance Importance, Sensitivity Sensitivity, ShowAs ShowAs, bool IsCancelled);
    private record AppointmentReminder(bool IsOn, int MinutesBeforeStart);
    private record AppointmentMetadata(bool HasAttachments, string? ICalUId, DateTimeOffset? Created, DateTimeOffset? Modified, string? ChangeKey, EventType Type);

    private static CalendarEvent MapAppointmentItem(dynamic item, BodyFormat bodyFormat = BodyFormat.Markdown)
    {
        var entryId = (string)item.EntryID;

        var body = TryMapAppointmentBody(item, bodyFormat);
        var startEnd = TryMapAppointmentStartEnd(item);
        var loc = TryMapAppointmentLocation(item);
        var organizer = TryMapAppointmentOrganizer(item);
        var attendees = TryMapAppointmentAttendees(item);
        var flags = TryMapAppointmentFlags(item);
        var reminder = TryMapAppointmentReminder(item);
        var categories = TryMapAppointmentCategories(item);
        var meta = TryMapAppointmentMetadata(item);

        return new CalendarEvent
        {
            Id = entryId,
            Subject = TryGetString(item, "Subject"),
            BodyPreview = TryGetString(item, "BodyPreview"),
            Body = body,
            Start = startEnd.Start,
            End = startEnd.End,
            IsAllDay = TryGetAppointmentAllDay(item),
            Location = loc.Location,
            Locations = loc.Locations,
            Organizer = organizer,
            Attendees = attendees,
            Importance = flags.Importance,
            Sensitivity = flags.Sensitivity,
            ShowAs = flags.ShowAs,
            IsCancelled = flags.IsCancelled,
            IsReminderOn = reminder.IsOn,
            ReminderMinutesBeforeStart = reminder.MinutesBeforeStart,
            Categories = categories,
            HasAttachments = meta.HasAttachments,
            Recurrence = null, // Recurrence-Mapping ist komplex (RecurrencePattern mit Type/Interval/DaysOfWeek/etc.); v1 Phase-3f auf null gesetzt
            ICalUId = meta.ICalUId,
            CreatedDateTime = meta.Created,
            LastModifiedDateTime = meta.Modified,
            ChangeKey = meta.ChangeKey,
            Type = meta.Type,
        };
    }

    /// <summary>
    /// Liest den Body (Text oder HTML) aus einem Outlook-Appointment-Item.
    /// BodyFormat MUSS vor Body/HTMLBody gelesen werden.
    /// </summary>
    private static ItemBody? TryMapAppointmentBody(dynamic item, BodyFormat bodyFormat = BodyFormat.Markdown)
    {
        try
        {
            string bodyText = TryGetString(item, "Body") ?? string.Empty;
            string bodyHtml = TryGetString(item, "HTMLBody") ?? string.Empty;
            // Native Quelle priorisieren: HTML wenn vorhanden, sonst Text.
            string? nativeSource = null;
            if (!string.IsNullOrEmpty(bodyHtml))
            {
                nativeSource = bodyHtml;
            }
            else if (!string.IsNullOrEmpty(bodyText))
            {
                nativeSource = bodyText;
            }
            if (nativeSource is null || nativeSource.Length == 0) return null;
            string htmlForConversion = nativeSource == bodyText
                ? $"<html><body><pre>{System.Net.WebUtility.HtmlEncode(bodyText)}</pre></body></html>"
                : nativeSource;
            var (ct, content) = HtmlBodyConverter.Convert(htmlForConversion, bodyFormat);
            return new ItemBody { ContentType = ct, Content = content };
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Liest Start und Ende als DateTimeTimeZone. Outlook liefert nur Windows-TZ-IDs;
    /// das ist eine bekannte Begrenzung des COM-APIs.
    /// </summary>
    private static AppointmentStartEnd TryMapAppointmentStartEnd(dynamic item)
    {
        DateTimeTimeZone? startDttz = null;
        DateTimeTimeZone? endDttz = null;
        try
        {
            var startDt = (DateTime)item.Start;
            var startTz = TryGetWindowsTimeZoneId(item, "StartTimeZone") ?? "UTC";
            startDttz = new DateTimeTimeZone
            {
                DateTime = startDt.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = startTz,
            };
        }
        catch { }
        try
        {
            var endDt = (DateTime)item.End;
            var endTz = TryGetWindowsTimeZoneId(item, "EndTimeZone") ?? "UTC";
            endDttz = new DateTimeTimeZone
            {
                DateTime = endDt.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = endTz,
            };
        }
        catch { }
        return new AppointmentStartEnd(startDttz, endDttz);
    }

    private static bool TryGetAppointmentAllDay(dynamic item)
    {
        try { return (bool)item.AllDayEvent; } catch { return false; }
    }

    /// <summary>
    /// Liest den Ort: ein einzelnes DisplayName-Feld, gemappt auf Location
    /// (Single) und Locations (Collection).
    /// </summary>
    private static AppointmentLocation TryMapAppointmentLocation(dynamic item)
    {
        var locations = new List<Location>();
        try
        {
            var locName = TryGetString(item, "Location");
            if (!string.IsNullOrEmpty(locName))
            {
                var loc = new Location { DisplayName = locName };
                locations.Add(loc);
                return new AppointmentLocation(loc, locations);
            }
        }
        catch { }
        return new AppointmentLocation(null, locations);
    }

    /// <summary>
    /// Liest den Organizer (Name + SMTP-Adresse). Gibt null zurueck, wenn keine
    /// Adresse vorhanden oder Organizer nicht lesbar.
    /// </summary>
    private static Organizer? TryMapAppointmentOrganizer(dynamic item)
    {
        try
        {
            dynamic org = item.Organizer;
            if (org is not null)
            {
                try
                {
                    var orgName = TryGetString(org, "Name");
                    var orgAddr = TryGetString(org, "Address");
                    if (!string.IsNullOrEmpty(orgAddr))
                    {
                        return new Organizer
                        {
                            EmailAddress = new EmailAddress { Name = orgName, Address = orgAddr },
                        };
                    }
                }
                finally { Marshal.ReleaseComObject(org); }
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Liest alle Attendees inkl. Response-Status (Accepted/Tentative/Declined)
    /// und Type (Required/Optional/Resource). OlRecipientType: 1=olRequired,
    /// 2=olOptional, 3=olResource.
    /// </summary>
    private static List<Attendee> TryMapAppointmentAttendees(dynamic item)
    {
        var attendees = new List<Attendee>();
        try
        {
            foreach (var att in item.Recipients)
            {
                try
                {
                    var attName = TryGetString(att, "Name");
                    var attAddr = TryGetString(att, "Address");
                    if (string.IsNullOrEmpty(attAddr)) continue;

                    int responseStatus = 0;
                    try { responseStatus = (int)att.MeetingResponseStatus; } catch { }
                    int typeRaw = 0;
                    try { typeRaw = (int)att.Type; } catch { }
                    var attType = typeRaw switch
                    {
                        2 => AttendeeType.Optional,
                        3 => AttendeeType.Resource,
                        _ => AttendeeType.Required,
                    };

                    attendees.Add(new Attendee
                    {
                        EmailAddress = new EmailAddress { Name = attName, Address = attAddr },
                        Type = attType,
                        Status = new ResponseStatus { Response = OlEnumMappings.ToResponse(responseStatus) },
                    });
                }
                finally { Marshal.ReleaseComObject(att); }
            }
        }
        catch { }
        return attendees;
    }

    /// <summary>
    /// Liest Importance / Sensitivity / ShowAs / Cancelled-Status.
    /// MeetingStatus: 0=olMeeting, 1=olMeetingReceived, 2=olMeetingCanceled,
    /// 3=olMeetingReceivedAndCanceled.
    /// </summary>
    private static AppointmentFlags TryMapAppointmentFlags(dynamic item)
    {
        var importance = Importance.Normal;
        try { importance = OlEnumMappings.ToImportance((int)item.Importance); } catch { }
        var sensitivity = Sensitivity.Normal;
        try { sensitivity = OlEnumMappings.ToSensitivity((int)item.Sensitivity); } catch { }
        var showAs = ShowAs.Busy;
        try { showAs = OlEnumMappings.ToShowAs((int)item.BusyStatus); } catch { }

        bool isCancelled = false;
        try
        {
            var status = (int)item.MeetingStatus;
            // olMeetingCanceled (2) und olMeetingReceivedAndCanceled (3)
            isCancelled = status == 2 || status == 3;
        }
        catch { }

        return new AppointmentFlags(importance, sensitivity, showAs, isCancelled);
    }

    private static AppointmentReminder TryMapAppointmentReminder(dynamic item)
    {
        bool isOn = false;
        try { isOn = (bool)item.ReminderSet; } catch { }
        int minutes = 0;
        try { minutes = (int)item.ReminderMinutesBeforeStart; } catch { }
        return new AppointmentReminder(isOn, minutes);
    }

    private static List<string> TryMapAppointmentCategories(dynamic item)
    {
        var categories = new List<string>();
        try
        {
            var catStr = TryGetString(item, "Categories") ?? string.Empty;
            if (!string.IsNullOrEmpty(catStr))
            {
                foreach (var c in catStr.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    categories.Add(c);
                }
            }
        }
        catch { }
        return categories;
    }

    /// <summary>
    /// Liest Attachment-Flag, iCalUId (GlobalAppointmentID), Erstell-/Aenderungs-
    /// Zeitstempel, ChangeKey und Event-Type (SingleInstance vs. SeriesMaster).
    /// </summary>
    private static AppointmentMetadata TryMapAppointmentMetadata(dynamic item)
    {
        bool hasAttachments = false;
        try { hasAttachments = (int)item.Attachments.Count > 0; } catch { }

        string? iCalUId = null;
        try { iCalUId = TryGetString(item, "GlobalAppointmentID"); } catch { }

        DateTimeOffset? created = null, modified = null;
        try { created = (DateTimeOffset)item.CreationTime; } catch { }
        try { modified = (DateTimeOffset)item.LastModificationTime; } catch { }

        string? changeKey = null;
        try { changeKey = TryGetString(item, "EntryID"); } catch { }

        var type = EventType.SingleInstance;
        try
        {
            if ((bool)item.IsRecurring)
            {
                type = EventType.SeriesMaster;
            }
        }
        catch { }

        return new AppointmentMetadata(hasAttachments, iCalUId, created, modified, changeKey, type);
    }

    /// <summary>
    /// Liest die Windows-TimeZone-ID (z. B. "W. Europe Standard Time") aus einem
    /// Outlook-TimeZone-Objekt (StartTimeZone/EndTimeZone). Outlook liefert hier
    /// keinen IANA-Namen; das ist eine bekannte Begrenzung des COM-APIs.
    /// </summary>
    private static string? TryGetWindowsTimeZoneId(dynamic item, string propertyName)
    {
        try
        {
            var tz = item.GetType().GetProperty(propertyName)?.GetValue(item);
            if (tz is null) return null;
            try
            {
                return (string?)tz.GetType().GetProperty("ID")?.GetValue(tz);
            }
            finally
            {
                Marshal.ReleaseComObject(tz);
            }
        }
        catch { return null; }
    }

    // ===== Calendar: Mutationen =====

    public async Task<string> CreateEventAsync(
        CreateEventRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            // v1: Standardmaessig in den Default-Calendar (kein CalendarId im
            // CreateEventRequest). Folge-Phase koennte Calendar-Parameter ergaenzen.
            dynamic calendar = GetMapiNamespace().GetDefaultFolder(9 /* OlDefaultFolders.olFolderCalendar */);
            dynamic appointment = calendar.Items.Add(1 /* OlItemType.olAppointmentItem */);
            try
            {
                FillAppointmentFromCreateRequest(appointment, request);
                AddAttendees(appointment, request.Attendees);

                if (request.SendInvitations && request.Attendees.Count > 0)
                {
                    // .Send() versendet Einladungen an alle Attendees und persistiert
                    // den Termin im Organizer-Kalender.
                    appointment.Send();
                }
                else
                {
                    // Nur persistieren (kein Mailversand).
                    appointment.Save();
                }

                return (string)appointment.EntryID;
            }
            finally
            {
                Marshal.ReleaseComObject(appointment);
                Marshal.ReleaseComObject(calendar);
            }
        }, nameof(CreateEventAsync));
    }

    public async Task UpdateEventAsync(
        string id,
        EventUpdate update,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic appointment = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                // PATCH-Semantik: nur gesetzte Felder anfassen. BodyFormat MUSS
                // vor .Body/.HTMLBody stehen.
                if (update.Body is { } body)
                {
                    var bodyFormat = OlEnumMappings.ToOlBodyFormat(body.ContentType);
                    appointment.BodyFormat = bodyFormat;
                    if (bodyFormat == 2)
                        appointment.HTMLBody = body.Content;
                    else
                        appointment.Body = body.Content;
                }
                if (update.Subject is { } subject)
                    appointment.Subject = subject;
                if (update.Start is { } start)
                    appointment.Start = DateTime.Parse(start.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                if (update.End is { } end)
                    appointment.End = DateTime.Parse(end.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                if (update.Location is { } location)
                    appointment.Location = location;
                if (update.ReminderMinutesBeforeStart is int rem)
                {
                    appointment.ReminderSet = true;
                    appointment.ReminderMinutesBeforeStart = rem;
                }
                if (update.Categories is not null)
                    appointment.Categories = string.Join(", ", update.Categories);
                if (update.ShowAs is { } showAs)
                    appointment.BusyStatus = OlEnumMappings.ToOlBusyStatus(showAs);
                if (update.Importance is { } importance)
                    appointment.Importance = OlEnumMappings.ToOlImportance(importance);
                if (update.Sensitivity is { } sensitivity)
                    appointment.Sensitivity = OlEnumMappings.ToOlSensitivity(sensitivity);
                if (update.Attendees is not null)
                {
                    // Bestehende Recipients entfernen, dann neu befuellen. Outlook
                    // COM bietet kein direktes ReplaceAll; daher Iteration Remove(1).
                    dynamic recipients = appointment.Recipients;
                    try
                    {
                        while ((int)recipients.Count > 0)
                        {
                            Marshal.ReleaseComObject(recipients.Item(1));
                            recipients.Remove(1);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(recipients);
                    }
                    AddAttendees(appointment, update.Attendees);
                }

                // .Send() benachrichtigt Attendees und speichert; .Save() nur lokal.
                if (update.SendUpdate && (int)appointment.Recipients.Count > 0)
                {
                    appointment.Send();
                }
                else
                {
                    appointment.Save();
                }
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(appointment);
            }
        }, nameof(UpdateEventAsync));
    }

    public async Task DeleteEventAsync(
        string id,
        bool sendCancellation = true,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic appointment = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                // Outlook COM hat kein separates Cancel() auf AppointmentItem.
                // Beim Loeschen eines Termins mit Attendees versendet Outlook
                // automatisch eine Cancellation ueber die Meeting-Infrastruktur
                // (CancellationMessageHandling). Wir unterscheiden den Parameter
                // sendCancellation aktuell nicht; .Delete() deckt beide Faelle ab.
                appointment.Delete();
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(appointment);
            }
        }, nameof(DeleteEventAsync));
    }

    public async Task RespondToEventAsync(
        string id,
        RespondToEventRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        await RunComAsync(() =>
        {
            dynamic appointment = GetMapiNamespace().GetItemFromID(id, Type.Missing);
            try
            {
                // Outlook.OlResponseStatus (1=Organizer, 2=Tentative, 3=Accepted,
                // 4=Declined); .Respond(Response, fAdditionalComments, additionalText)
                // sendet die Antwort-Mail und aktualisiert den Termin-Status.
                var olResponse = OlEnumMappings.ToOlResponseStatus(request.Response);
                bool hasComment = !string.IsNullOrEmpty(request.Comment);
                appointment.Respond(olResponse, hasComment, request.Comment ?? string.Empty);
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(appointment);
            }
        }, nameof(RespondToEventAsync));
    }

    public async Task<IReadOnlyList<MeetingTimeCandidate>> FindMeetingTimesAsync(
        FindMeetingTimesRequest request,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var candidates = new List<MeetingTimeCandidate>();
            dynamic calendar = GetMapiNamespace().GetDefaultFolder(9 /* OlDefaultFolders.olFolderCalendar */);
            try
            {
                var windowStart = DateTime.Parse(request.TimeWindowStart.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                var windowEnd = DateTime.Parse(request.TimeWindowEnd.DateTime, System.Globalization.CultureInfo.InvariantCulture);
                var duration = TimeSpan.FromMinutes(request.DurationMinutes);
                if (windowEnd <= windowStart || duration <= TimeSpan.Zero) return (IReadOnlyList<MeetingTimeCandidate>)candidates;

                dynamic items = calendar.Items;
                try
                {
                    items.IncludeRecurrences = true;
                    var startStr = windowStart.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    var endStr = windowEnd.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture);
                    items = items.Restrict($"[Start] >= '{startStr}' AND [End] <= '{endStr}'");
                    items.Sort("[Start]");

                    // Busy-Intervalle sammeln (AllDay-Events ignorieren wir,
                    // da Outlook sie als Hintergrund busy markiert; fuer freie
                    // Zeit-Slots sind sie aber oft hinderlich. Tunenvariant: mit
                    // AllDay ist eine spaetere Phase).
                    var busyList = new List<(DateTime Start, DateTime End)>();
                    for (int i = 1; i <= (int)items.Count; i++)
                    {
                        dynamic item = items.Item(i);
                        try
                        {
                            bool isAllDay = false;
                            try { isAllDay = (bool)item.AllDayEvent; } catch { }
                            if (isAllDay) continue;

                            int busyStatus = 0;
                            try { busyStatus = (int)item.BusyStatus; } catch { }
                            // 0=olFree, 1=olTentative, 2=olBusy, 3=olOutOfOffice, 4=olWorkingElsewhere
                            // 0 = frei -> kein Gap-Effekt
                            if (busyStatus == 0) continue;

                            var s = (DateTime)item.Start;
                            var e = (DateTime)item.End;
                            if (e <= s) continue;
                            busyList.Add((s, e));
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(item);
                        }
                    }

                    // Gap-Scan im Zeitfenster. Outlook-Termine sind meist sauber
                    // definiert ohne grosse Overlaps; Overlap-Merge waere v1+.
                    var sortedBusy = busyList.OrderBy(b => b.Start).ToList();

                    DateTime cursor = windowStart;
                    foreach (var (bs, be) in sortedBusy)
                    {
                        if (bs >= windowEnd) break;
                        if (bs > cursor)
                        {
                            AddCandidateIfFits(candidates, cursor, bs < windowEnd ? bs : windowEnd, duration);
                        }
                        if (be > cursor) cursor = be;
                        if (cursor >= windowEnd) break;
                    }
                    if (cursor < windowEnd)
                    {
                        AddCandidateIfFits(candidates, cursor, windowEnd, duration);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(items);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(calendar);
            }

            // Ranking: laengste Luecken zuerst, dann fruehester Start.
            var ranked = candidates
                .OrderByDescending(c => (DateTime.Parse(c.End.DateTime, System.Globalization.CultureInfo.InvariantCulture)
                                       - DateTime.Parse(c.Start.DateTime, System.Globalization.CultureInfo.InvariantCulture)))
                .ThenBy(c => c.Start.DateTime)
                .Take(request.MaxCandidates)
                .ToList();
            return (IReadOnlyList<MeetingTimeCandidate>)ranked;
        }, nameof(FindMeetingTimesAsync));
    }

    // ===== Calendar: Mutations-Helpers =====

    /// <summary>
    /// Befuellt ein AppointmentItem mit Feldern aus einem CreateEventRequest.
    /// </summary>
    private static void FillAppointmentFromCreateRequest(dynamic appointment, CreateEventRequest request)
    {
        appointment.Subject = request.Subject;
        if (request.Body is not null)
        {
            var bodyFormat = OlEnumMappings.ToOlBodyFormat(request.Body.ContentType);
            appointment.BodyFormat = bodyFormat;
            if (bodyFormat == 2 /* olFormatHTML */)
                appointment.HTMLBody = request.Body.Content;
            else
                appointment.Body = request.Body.Content;
        }
        appointment.Start = DateTime.Parse(request.Start.DateTime, System.Globalization.CultureInfo.InvariantCulture);
        appointment.End = DateTime.Parse(request.End.DateTime, System.Globalization.CultureInfo.InvariantCulture);
        appointment.AllDayEvent = request.IsAllDay;
        if (!string.IsNullOrEmpty(request.Location))
            appointment.Location = request.Location;
        if (request.ReminderMinutesBeforeStart is int rem)
        {
            appointment.ReminderSet = true;
            appointment.ReminderMinutesBeforeStart = rem;
        }
        if (request.Categories.Count > 0)
            appointment.Categories = string.Join(", ", request.Categories);
        appointment.BusyStatus = OlEnumMappings.ToOlBusyStatus(request.ShowAs);
        appointment.Importance = OlEnumMappings.ToOlImportance(request.Importance);
        appointment.Sensitivity = OlEnumMappings.ToOlSensitivity(request.Sensitivity);
    }

    /// <summary>
    /// Fuegt Event-Attendees als COM-Recipients zu einem AppointmentItem hinzu,
    /// setzt Type (Required/Optional/Resource) und ruft Resolve() pro Empfaenger
    /// und ResolveAll() am Ende auf. Resolve braucht es, damit Outlook SMTP-Adressen
    /// aufloest (sonst kein Send moeglich, .Send wirft).
    /// </summary>
    private static void AddAttendees(dynamic appointment, IReadOnlyList<EventAttendeeInput> attendees)
    {
        if (attendees.Count == 0) return;
        foreach (var att in attendees)
        {
            dynamic recipient = appointment.Recipients.Add(att.Email);
            try
            {
                if (!string.IsNullOrEmpty(att.Name))
                {
                    try { recipient.Name = att.Name; } catch { /* Resolve kann Name ueberschreiben */ }
                }
                // OlRecipientType: 1=olRequired, 2=olOptional, 3=olResource
                int olType = att.Type switch
                {
                    AttendeeType.Optional => 2,
                    AttendeeType.Resource => 3,
                    _ => 1,
                };
                recipient.Type = olType;
                try { recipient.Resolve(); } catch { /* ungeloest -> Skip */ }
            }
            finally
            {
                Marshal.ReleaseComObject(recipient);
            }
        }
        try { appointment.Recipients.ResolveAll(); } catch { }
    }

    /// <summary>
    /// Prueft, ob [start, end) gross genug fuer duration ist, und fuegt
    /// ggf. einen MeetingTimeCandidate ein. Confidence = 100, da lokal berechnet.
    /// </summary>
    private static void AddCandidateIfFits(
        List<MeetingTimeCandidate> candidates,
        DateTime start,
        DateTime end,
        TimeSpan duration)
    {
        if (end - start < duration) return;
        candidates.Add(new MeetingTimeCandidate
        {
            Start = new DateTimeTimeZone
            {
                DateTime = start.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            End = new DateTimeTimeZone
            {
                DateTime = (start + duration).ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC",
            },
            Confidence = 100,
            ConflictsAttendeeCount = 0,
        });
    }

    // ===== Active-Inspector / Selection (Phase 3h) =====

    public async Task<ActiveItem?> GetActiveItemAsync(
        BodyFormat bodyFormat,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            // null ist valides Resultat: kein Inspector offen oder out-of-scope
            // Typ (Tasks/Contacts). Wird defensiv abgefangen, damit kein
            // OutlookServiceException fuer den Normalfall geworfen wird.
            dynamic? inspector = null;
            dynamic? currentItem = null;
            try
            {
                inspector = TryGetActiveInspector();
                if (inspector is null) return (ActiveItem?)null;

                currentItem = TryGetInspectorCurrentItem(inspector);
                if (currentItem is null) return (ActiveItem?)null;

                return MapActiveItemFromCurrentItem(currentItem, bodyFormat, _logger);
            }
            finally
            {
                if (currentItem is not null) Marshal.ReleaseComObject(currentItem);
                if (inspector is not null) Marshal.ReleaseComObject(inspector);
            }
        }, nameof(GetActiveItemAsync));
    }

    /// <summary>
    /// Holt den ActiveInspector defensiv. Gibt null zurueck, wenn kein
    /// Inspector offen ist oder eine COMException geworfen wird.
    /// </summary>
    private dynamic? TryGetActiveInspector()
    {
        try { return _outlookApp!.ActiveInspector(); }
        catch (COMException) { return null; }
    }

    /// <summary>
    /// Holt CurrentItem aus dem Inspector defensiv. Gibt null zurueck, wenn
    /// keine CurrentItem vorhanden oder COMException.
    /// </summary>
    private static dynamic? TryGetInspectorCurrentItem(dynamic inspector)
    {
        try { return inspector.CurrentItem; }
        catch (COMException) { return null; }
    }

    /// <summary>
    /// Mappt CurrentItem auf ActiveItem-DTO basierend auf OlObjectClass.
    /// OlObjectClass: 43=olMail, 26=olAppointment, 48=olTask, 40=olContact, ...
    /// Tasks/Contacts/etc. -> v1.1: null.
    /// </summary>
    private static ActiveItem? MapActiveItemFromCurrentItem(dynamic currentItem, BodyFormat bodyFormat, ILogger? logger = null)
    {
        int objectClass = 0;
        try { objectClass = (int)currentItem.Class; } catch { }
        return objectClass switch
        {
            43 /* olMail */ => new ActiveMail { Item = MapMailItem(currentItem, true, bodyFormat, logger) },
            26 /* olAppointment */ => new ActiveEvent { Item = MapAppointmentItem(currentItem) },
            _ => null,
        };
    }

    public async Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top,
        BodyFormat bodyFormat,
        CancellationToken cancellationToken = default)
    {
        await GetOutlookApplicationAsync(cancellationToken);
        return await RunComAsync(() =>
        {
            var result = new List<ActiveItem>();
            // Aktiver Explorer (Mail-/Calendar-Explorer). ActiveExplorer()==null
            // -> OutlookNotActive (OutlookServiceException durch OutlookServiceException-Wrap).
            dynamic? explorer = null;
            try
            {
                explorer = _outlookApp!.ActiveExplorer();
            }
            catch (COMException ex)
            {
                throw new OutlookServiceException(
                    OlEnumMappings.FromHResult(ex.HResult),
                    $"ActiveExplorer nicht abrufbar: {ex.Message}",
                    ex);
            }
            if (explorer is null)
            {
                throw new OutlookServiceException(
                    ErrorCode.OutlookNotActive,
                    "Outlook ActiveExplorer ist null (kein Explorer-Fenster offen).");
            }

            try
            {
                dynamic selection = explorer.Selection;
                try
                {
                    int count = (int)selection.Count;
                    if (count == 0) return (IReadOnlyList<ActiveItem>)result; // valides Empty-Result

                    int take = Math.Min(count, top);
                    for (int i = 1; i <= take; i++)
                    {
                        dynamic item = selection.Item(i);
                        try
                        {
                            int objectClass = 0;
                            try { objectClass = (int)item.Class; } catch { continue; }

                            bool include = scope switch
                            {
                                SelectionScope.Any => true,
                                SelectionScope.Mail => objectClass == 43,
                                SelectionScope.Calendar => objectClass == 26,
                                _ => false,
                            };
                            if (!include) continue;

                            ActiveItem? mapped = objectClass switch
                            {
                                43 => new ActiveMail { Item = MapMailItem(item, true, bodyFormat, _logger) },
                                26 => new ActiveEvent { Item = MapAppointmentItem(item) },
                                _ => null,
                            };
                            if (mapped is not null) result.Add(mapped);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(item);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(selection);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(explorer);
            }
            return (IReadOnlyList<ActiveItem>)result;
        }, nameof(GetSelectedItemsAsync));
    }
}