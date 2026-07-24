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
    /// Multi-Store-Profilen (z. B. Exchange + PST/Notes) liefert Outlook fuer
    /// dieselbe olId je nach Store einen anderen Folder. Daher gilt:
    ///
    /// 1. Zuerst `ns.GetDefaultFolder(olId)` probieren.
    /// 2. Danach ueber `session.Stores` iterieren.
    ///
    /// Der erste Treffer wird zurueckgegeben, ohne Schema-Pruefung.
    /// </summary>
    private dynamic? ResolveDefaultFolderSmart(int olFolderId)
    {
        try
        {
            dynamic ns = GetMapiNamespace();
            return ns.GetDefaultFolder(olFolderId);
        }
        catch (COMException)
        {
            // Fallback auf Stores unten.
        }

        try
        {
            dynamic ns = GetMapiNamespace();
            dynamic stores = ns.Session.Stores;
            try
            {
                foreach (var store in stores)
                {
                    try
                    {
                        dynamic folder = store.GetDefaultFolder(olFolderId);
                        if (folder is not null) return folder;
                    }
                    catch (COMException)
                    {
                        // Store kennt die olId nicht.
                    }
                }
            }
            finally
            {
                ReleaseComObject(stores);
            }
        }
        catch
        {
            // kein weiterer Fallback
        }

        return null;
    }

    /// <summary>
    /// Wie <see cref="ResolveDefaultFolderSmart"/>, aber mit Mail-Schema-Pruefung.
    /// Notes-/Journal-/Calendar-Folder werden ignoriert. Falls mehrere Stores
    /// denselben olId-Wert liefern, wird der erste Folder mit Mail-Schema aus
    /// dem Default-Pfad bevorzugt.
    /// </summary>
    private dynamic? ResolveDefaultFolderByOlIdWithSchemaCheck(int olFolderId)
    {
        var candidates = new List<dynamic>();

        try
        {
            dynamic ns = GetMapiNamespace();
            dynamic folder = ns.GetDefaultFolder(olFolderId);
            if (folder is not null) candidates.Add(folder);
        }
        catch (COMException) { }

        try
        {
            dynamic ns = GetMapiNamespace();
            dynamic stores = ns.Session.Stores;
            try
            {
                foreach (var store in stores)
                {
                    try
                    {
                        dynamic folder = store.GetDefaultFolder(olFolderId);
                        if (folder is not null) candidates.Add(folder);
                    }
                    catch (COMException) { }
                }
            }
            finally
            {
                ReleaseComObject(stores);
            }
        }
        catch { }

        foreach (var folder in candidates)
        {
            if (IsMailFolder(folder)) return folder;
            ReleaseComObject(folder);
        }

        return null;
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


    public async Task<ActiveItem?> GetActiveItemAsync(
        BodyFormat? bodyFormat,
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

                int objectClass = 0;
                try { objectClass = (int)currentItem.Class; } catch { }
                return objectClass switch
                {
                    43 /* olMail */ => new ActiveMail
                    {
                        Item = ReloadInspectorMailItem(currentItem, bodyFormat, _logger),
                    },
                    26 /* olAppointment */ => new ActiveEvent
                    {
                        Item = MapAppointmentItem(currentItem, bodyFormat),
                    },
                    _ => null,
                };
            }
            finally
            {
                ReleaseComObject(currentItem);
                ReleaseComObject(inspector);
            }
        }, nameof(GetActiveItemAsync));
    }

    public async Task<IReadOnlyList<ActiveItem>> GetSelectedItemsAsync(
        SelectionScope scope,
        int top,
        BodyFormat? bodyFormat,
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
                                43 => new ActiveMail { Item = MapMailItem(item, bodyFormat, _logger) },
                                26 => new ActiveEvent { Item = MapAppointmentItem(item, bodyFormat) },
                                _ => null,
                            };
                            if (mapped is not null) result.Add(mapped);
                        }
                        finally
                        {
                            ReleaseComObject(item);
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(selection);
                }
            }
            finally
            {
                ReleaseComObject(explorer);
            }
            return (IReadOnlyList<ActiveItem>)result;
        }, nameof(GetSelectedItemsAsync));
    }
}