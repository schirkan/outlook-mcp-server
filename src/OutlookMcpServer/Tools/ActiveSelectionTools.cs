using System.ComponentModel;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Tools;

/// <summary>
/// MCP-Tools fuer Active-Inspector + Selection (COM-only).
/// Naming: <c>snake_case</c> (z. B. <c>get_active_item</c>, <c>get_selected_items</c>).
/// </summary>
[McpServerToolType]
public sealed class ActiveSelectionTools
{
    private readonly IOutlookService _service;
    private readonly ILogger<ActiveSelectionTools> _logger;

    public ActiveSelectionTools(IOutlookService service, ILogger<ActiveSelectionTools> logger)
    {
        _service = service;
        _logger = logger;
    }

    [McpServerTool(Name = "get_active_item")]
    [Description("Liefert das aktuell im aktiven Outlook-Inspector-Fenster offene Item (MailItem oder AppointmentItem). null wenn kein Inspector offen oder v1-out-of-scope Typ (Tasks/Contacts). bodyFormat steuert nur den Mail-Body von ActiveMail; AppointmentItem hat keinen Body.")]
    public async Task<ActiveItem?> GetActiveItem(
        [Description(@"Format des Mail-Body (siehe get_mail). Default 'markdown'. Wirkt nur auf ActiveMail — ActiveEvent (AppointmentItem) hat keinen Body.")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        _logger.LogInformation("get_active_item bodyFormat={Bf}", bf);
        return await _service.GetActiveItemAsync(bf, cancellationToken);
    }

    [McpServerTool(Name = "get_selected_items")]
    [Description("Liefert die im aktiven Outlook-Explorer markierten Items, optional gefiltert nach scope (mail/calendar/any) und gecappt durch top (1-250). Selection.Count==0 ist valides Empty-Result (kein Fehler). Wirft OutlookNotActive wenn ActiveExplorer()==null.")]
    public async Task<IReadOnlyList<ActiveItem>> GetSelectedItems(
        [Description("Scope-Filter: 'mail' = nur Mails, 'calendar' = nur Termine, 'any' = alle (Default any).")] string scope = "any",
        [Description("Max Anzahl (1-250, Default 50).")] int top = 50,
        [Description(@"Format des Mail-Body (siehe get_mail). Default 'markdown'. Wirkt nur auf enthaltene ActiveMails.")] string? bodyFormat = null,
        CancellationToken cancellationToken = default)
    {
        var bf = BodyFormatExtensions.ParseBodyFormat(bodyFormat);
        _logger.LogInformation("get_selected_items scope={Scope} top={Top} bodyFormat={Bf}", scope, top, bf);
        var parsedScope = ParseSelectionScope(scope);
        return await _service.GetSelectedItemsAsync(parsedScope, top, bf, cancellationToken);
    }

    private static SelectionScope ParseSelectionScope(string? value) => (value ?? "any").Trim().ToLowerInvariant() switch
    {
        "mail" => SelectionScope.Mail,
        "calendar" => SelectionScope.Calendar,
        _ => SelectionScope.Any,
    };
}