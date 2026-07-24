using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    /// <summary>
    /// Holt den aktuell aktiven Inspector und gibt bei COM-Fehlern null zurueck.
    /// </summary>
    private dynamic? TryGetActiveInspector()
    {
        try { return _outlookApp!.ActiveInspector(); }
        catch (COMException) { return null; }
    }

    /// <summary>
    /// Liest das aktuelle Inspector-Item defensiv aus.
    /// </summary>
    private static dynamic? TryGetInspectorCurrentItem(dynamic inspector)
    {
        try { return inspector.CurrentItem; }
        catch (COMException) { return null; }
    }

    /// <summary>
    /// Mappt ein aktives Outlook-Objekt anhand seiner ObjectClass auf das passende DTO.
    /// </summary>
    private static ActiveItem? MapActiveItemFromCurrentItem(dynamic currentItem, BodyFormat? bodyFormat, ILogger? logger = null)
    {
        int objectClass = 0;
        try { objectClass = (int)currentItem.Class; } catch { }
        return objectClass switch
        {
            43 => new ActiveMail { Item = MapMailItem(currentItem, bodyFormat, logger) },
            26 => new ActiveEvent { Item = MapAppointmentItem(currentItem, bodyFormat) },
            _ => null,
        };
    }

    /// <summary>
    /// Laedt ein im Inspector geoeffnetes MailItem per EntryID neu und mappt es stabil.
    /// </summary>
    private MailMessage ReloadInspectorMailItem(dynamic currentItem, BodyFormat? bodyFormat, ILogger? logger = null)
    {
        string entryId = (string)currentItem.EntryID;
        dynamic reloaded = GetMapiNamespace().GetItemFromID(entryId, Type.Missing);
        try
        {
            return MapMailItem(reloaded, bodyFormat, logger);
        }
        finally
        {
            ReleaseComObject(reloaded);
        }
    }
}
