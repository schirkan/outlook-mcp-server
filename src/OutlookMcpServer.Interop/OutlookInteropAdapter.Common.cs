using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using OutlookMcpServer.Domain.Exceptions;

namespace OutlookMcpServer.Interop;

public sealed partial class OutlookInteropAdapter
{
    /// <summary>
    /// Heuristik fuer Mail-Folder. Outlook.DefaultItemType verwendet fuer
    /// Folder-Inhalte u. a. folgende Werte:
    ///   0 = olMailItem
    ///   1 = olAppointmentItem
    ///   2 = olContactItem
    ///   3 = olTaskItem
    ///   4 = olJournalItem
    ///   5 = olNoteItem
    ///   6 = olPostItem
    ///
    /// Fuer Mail-Folder akzeptieren wir NUR olMailItem=0.
    /// </summary>
    private static bool IsMailFolder(dynamic folder)
    {
        if (folder is null) return false;
        try
        {
            int defaultItemType = (int)folder.DefaultItemType;
            return defaultItemType == 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Kleine serverseitige DASL-Normalisierung fuer die haeufigsten
    /// Outlook-Property-Namen, die in der Praxis oft mit falscher
    /// Gross-/Kleinschreibung geschrieben werden und sonst in Restrict()
    /// mit 0x80020009 enden wuerden.
    /// </summary>
    private static string? NormalizeDaslFilter(string? filter)
    {
        if (string.IsNullOrWhiteSpace(filter)) return filter;
        return System.Text.RegularExpressions.Regex.Replace(
            filter,
            @"\[UnRead\]",
            "[Unread]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    /// <summary>
    /// Defensive Property-Access-Wrapper fuer COM dynamic-Properties.
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

    /// <summary>
    /// Liest eine dynamische COM-Property und gibt bei Fehlern null zurueck.
    /// </summary>
    private static dynamic? TryGetDynamic(dynamic obj, string propertyName)
    {
        try
        {
            return GetDynamicProperty(obj, propertyName);
        }
        catch { return null; }
    }

    /// <summary>
    /// Loest eine COM-Property ueber den C#-Runtime-Binder auf.
    /// </summary>
    private static object? GetDynamicProperty(dynamic obj, string name)
    {
        var binder = Binder.GetMember(
            CSharpBinderFlags.None,
            name,
            null,
            new[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) });
        var callsite = CallSite<Func<CallSite, object, object>>.Create(binder);
        return callsite.Target(callsite, (object)obj);
    }

    /// <summary>
    /// Gibt ein COM-Objekt defensiv frei und ignoriert Release-Fehler.
    /// </summary>
    private static void ReleaseComObject(object? value)
    {
        if (value is not null)
        {
            try { Marshal.ReleaseComObject(value); } catch { }
        }
    }

    /// <summary>
    /// Kombiniert benutzerdefinierten DASL-Filter mit optionalem Volltext-Such-Query.
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
    /// Instanziiert Outlook.Application via COM.
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
}
