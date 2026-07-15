using System.ComponentModel;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Tools;

/// <summary>
/// MCP-Tools fuer Active-Inspector + Selection. COM-only, kein Graph-Aequivalent.
/// Naming: <c>snake_case</c> wie Microsoft Graph.
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
    [Description("Liefert das aktuell im aktiven Outlook-Inspector-Fenster offene Item (MailItem oder AppointmentItem). null wenn kein Inspector offen oder v1-out-of-scope Typ (Tasks/Contacts).")]
    public async Task<ActiveItem?> GetActiveItem(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("get_active_item");
        return await _service.GetActiveItemAsync(cancellationToken);
    }

    [McpServerTool(Name = "get_selected_items")]
    [Description("Liefert die im aktiven Outlook-Explorer markierten Items, optional gefiltert nach scope (mail/calendar/any) und gecappt durch top (1-250). Selection.Count==0 ist valides Empty-Result (kein Fehler). Wirft OutlookNotActive wenn ActiveExplorer()==null.")]
    public async Task<IReadOnlyList<ActiveItem>> GetSelectedItems(
        [Description("Scope-Filter: 'mail' = nur Mails, 'calendar' = nur Termine, 'any' = alle (Default any).")] string scope = "any",
        [Description("Max Anzahl (1-250, Default 50).")] int top = 50,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("get_selected_items scope={Scope} top={Top}", scope, top);
        var parsedScope = ParseSelectionScope(scope);
        return await _service.GetSelectedItemsAsync(parsedScope, top, cancellationToken);
    }

    private static SelectionScope ParseSelectionScope(string? value) => (value ?? "any").Trim().ToLowerInvariant() switch
    {
        "mail" => SelectionScope.Mail,
        "calendar" => SelectionScope.Calendar,
        _ => SelectionScope.Any,
    };
}