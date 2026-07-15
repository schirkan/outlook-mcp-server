using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using OutlookMcpServer.Domain.Configuration;
using OutlookMcpServer.Interop;

namespace OutlookMcpServer.IntegrationTests;

/// <summary>
/// Basis fuer alle Integration-Tests, die gegen ein laufendes klassisches
/// Outlook-Profil testen.
/// <para>
/// <b>Outlook-Verfuegbarkeit</b>: Der Constructor prueft via ProgID-Lookup, ob
/// Outlook (klassischer Desktop-Client) installiert und/oder laeuft. Falls nicht,
/// wird <see cref="OutlookAvailable"/> = false gesetzt. Tests muessen dann
/// <see cref="SkipIfOutlookNotAvailable"/> am Anfang aufrufen (typischerweise
/// via <c>Skip.IfNot(OutlookAvailable, ...)</c> aus Xunit.SkippableFact).
/// </para>
/// <para>
/// <b>Adapter-Instanz</b>: Direkte Instanziierung von <see cref="OutlookInteropAdapter"/>
/// ohne DI-Container. <see cref="AutoStartOutlook"/> = false: Tests starten
/// Outlook NICHT automatisch — wenn Outlook nicht laeuft, schlagen Tests
/// sauber fehl bzw. werden uebersprungen statt ungewollt Outlook hochzufahren.
/// </para>
/// </summary>
public abstract class OutlookIntegrationTestBase : IDisposable
{
    protected OutlookInteropAdapter Adapter { get; }
    protected bool OutlookAvailable { get; }
    protected string? OutlookUnavailableReason { get; }

    protected OutlookIntegrationTestBase()
    {
        var options = Options.Create(new OutlookMcpServerOptions
        {
            Outlook = new OutlookOptions
            {
                AutoStartOutlook = false,
                StartupTimeoutSeconds = 5,
                MaxAttachmentBytes = 25 * 1024 * 1024,
                AllowSend = false,
                AllowCreate = false,
                AllowDelete = false,
            },
        });
        ILogger<OutlookInteropAdapter> logger = NullLogger<OutlookInteropAdapter>.Instance;
        Adapter = new OutlookInteropAdapter(options, logger);

        var (available, reason) = DetectOutlook();
        OutlookAvailable = available;
        OutlookUnavailableReason = reason;
    }

    /// <summary>
    /// Prueft Outlook-Verfuegbarkeit via ProgID-Lookup "Outlook.Application".
    /// Wenn ProgID nicht registriert ist: Outlook ist nicht installiert.
    /// Wenn Activator.CreateInstance nicht antwortet (Timeout): COM-Server
    /// nicht erreichbar (z. B. im Sandbox ohne Outlook-Profil) -> false.
    /// Der 5-Sekunden-Timeout ist wichtig, weil Activator.CreateInstance bei
    /// nicht erreichbarem COM-Server sonst endlos haengt und den Testhost
    /// zum Stillstand bringt.
    /// </summary>
    private static (bool available, string? reason) DetectOutlook()
    {
        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType is null)
            {
                return (false, "Outlook ProgID nicht registriert (Outlook nicht installiert)");
            }

            // Aktivierung in Task mit Timeout, weil Activator.CreateInstance
            // bei nicht erreichbarem COM-Server endlos haengen kann.
            var task = System.Threading.Tasks.Task.Run(() =>
            {
                try { return (object?)Activator.CreateInstance(outlookType); }
                catch { return null; }
            });
            if (!task.Wait(TimeSpan.FromSeconds(5)))
            {
                return (false, "Outlook COM-Server reagiert nicht innerhalb von 5s");
            }
            var instance = task.Result;
            if (instance is null)
            {
                return (false, "Activator.CreateInstance lieferte null oder warf Exception");
            }
            Marshal.ReleaseComObject(instance);
            return (true, null);
        }
        catch (Exception ex)
        {
            return (false, $"{ex.GetType().Name}: {ex.Message}");
        }
    }

    /// <summary>
    /// Convenience: ruft <c>Skip.IfNot(OutlookAvailable, ...)</c> auf.
    /// Tests rufen das als erste Zeile auf, damit sie ohne Outlook sauber
    /// uebersprungen werden (nicht fehlschlagen).
    /// </summary>
    protected void SkipIfOutlookNotAvailable()
    {
        Skip.IfNot(OutlookAvailable, OutlookUnavailableReason ?? "Outlook nicht verfuegbar");
    }

    public virtual void Dispose()
    {
        // OutlookInteropAdapter hat aktuell keine Dispose-Logik; _comLock wird
        // bei GetOutlookApplicationAsync erworben. Falls der Adapter spaeter
        // IDisposable implementiert, hier freigeben.
        GC.SuppressFinalize(this);
    }
}