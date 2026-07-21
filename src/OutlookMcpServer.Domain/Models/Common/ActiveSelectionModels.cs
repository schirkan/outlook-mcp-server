using System.Text.Json.Serialization;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Domain.Models.Common;

/// <summary>
/// Diskriminierter Union des "gerade in Outlook aktiven Items" (COM-only,
/// kein Graph-Aequivalent). v1-Scope: <c>mail</c> und <c>event</c>.
/// Tasks/Contacts bleiben v1.1.
/// Serialisierung polymorph ueber <see cref="JsonDerivedTypeAttribute"/>
/// mit Diskriminator <c>kind</c> — MCP-Clients koennen ohne Type-Hints
/// auf <c>kind</c> dispatchen.
/// </summary>
[JsonDerivedType(typeof(ActiveMail),  "mail")]
[JsonDerivedType(typeof(ActiveEvent), "event")]
public abstract record ActiveItem
{
    // Get-only (kein init) auf der Basisklasse, damit der
    // System.Text.Json-Source-Generator keine doppelten Initializer
    // fuer ActiveMail/ActiveEvent emittiert (CS1912 bei Trimming/Publish).
    // Wert wird ueber den override in den Sub-Typen festgelegt.
    [JsonPropertyName("kind")]
    public abstract string Kind { get; }
}

/// <summary>ActiveItem-Variante: MailItem in Inspector.</summary>
public sealed record ActiveMail : ActiveItem
{
    public override string Kind { get; } = "mail";

    [JsonPropertyName("item")]
    public required MailMessage Item { get; init; }
}

/// <summary>ActiveItem-Variante: AppointmentItem in Inspector.</summary>
public sealed record ActiveEvent : ActiveItem
{
    public override string Kind { get; } = "event";

    [JsonPropertyName("item")]
    public required CalendarEvent Item { get; init; }
}

/// <summary>
/// Ergebnis von <c>getSelectedItems</c>. <c>Value</c> ist nach
/// <c>scope</c> gefiltert und mit <c>top</c> gecappt; <c>Count</c>
/// ist die Anzahl zurueckgegebener Items (Value.Count).
/// </summary>
public sealed record SelectedItemsResult
{
    [JsonPropertyName("value")]
    public IReadOnlyList<ActiveItem> Value { get; init; } = Array.Empty<ActiveItem>();

    [JsonPropertyName("count")]
    public int Count { get; init; }
}

/// <summary>
/// Filter fuer <c>getSelectedItems</c>. <c>Any</c> (= default) liefert
/// alle Typen; <c>Mail</c>/<c>Calendar</c> filtern auf den jeweils
/// passenden Item-Typ. Outlook selbst hat im Standard-Explorer keine
/// gemischten Listen — der Filter dient als Defensiv-Logik.
/// </summary>
public enum SelectionScope
{
    [JsonPropertyName("mail")]
    Mail,

    [JsonPropertyName("calendar")]
    Calendar,

    [JsonPropertyName("any")]
    Any,
}
