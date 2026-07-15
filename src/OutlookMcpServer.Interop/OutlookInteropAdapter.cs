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

    private async Task<T> RunComAsync<T>(Func<Task<T>> func, string opName)
    {
        try
        {
            return await func();
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

    public Task<PagedResult<MailFolder>> ListMailFoldersAsync(
        string? parentFolderId = null,
        bool includeHidden = false,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: echte Implementation
        return Task.FromResult(new PagedResult<MailFolder>
        {
            Value = Array.Empty<MailFolder>(),
            NextSkip = null,
        });
    }

    public Task<MailFolder> GetMailFolderAsync(
        string folderId,
        CancellationToken cancellationToken = default)
    {
        throw new NotImplementedException("GetMailFolderAsync - wird in Karte 3.5 implementiert");
    }

    // ===== Mail: Nachrichten =====

    public Task<PagedResult<MailMessage>> ListMailsAsync(
        string folderId,
        int top = 25,
        int skip = 0,
        string? filter = null,
        string? search = null,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: echte Implementation mit Items.Restrict + Sort
        return Task.FromResult(new PagedResult<MailMessage>
        {
            Value = Array.Empty<MailMessage>(),
            NextSkip = null,
        });
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

    public Task<PagedResult<MailMessage>> SearchMailsAsync(
        string query,
        string? folderId = null,
        int top = 25,
        CancellationToken cancellationToken = default)
    {
        // TODO Karte 3.5: Outlook-InstanteSearch oder AdvancedSearch
        return Task.FromResult(new PagedResult<MailMessage>
        {
            Value = Array.Empty<MailMessage>(),
            NextSkip = null,
        });
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
}