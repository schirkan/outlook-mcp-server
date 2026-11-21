namespace OutlookMcpServer.Domain.Configuration;

/// <summary>
/// Konfigurations-Wurzel fuer den MCP-Server. Bindet gegen
/// <c>appsettings.json</c> Sektion <c>OutlookMcpServer</c>.
/// Wird per <c>Microsoft.Extensions.Options.IOptions&lt;T&gt;</c> injiziert.
/// </summary>
public sealed class OutlookMcpServerOptions
{
    public OutlookOptions Outlook { get; set; } = new();

    public LoggingOptions Logging { get; set; } = new();
}

public sealed class OutlookOptions
{
    /// <summary>Profilname fuer Outlook (null = Default-Profil).</summary>
    public string? ProfileName { get; set; }

    public bool AutoStartOutlook { get; set; } = true;

    public int StartupTimeoutSeconds { get; set; } = 30;

    public bool AllowSend { get; set; } = true;

    public bool AllowDelete { get; set; } = true;

    public bool AllowCreate { get; set; } = true;

    /// <summary>Max. Attachment-Groesse in Bytes (default 25 MB = Outlook-Default).</summary>
    public long MaxAttachmentBytes { get; set; } = 26_214_400;
}

public sealed class LoggingOptions
{
    public LogLevelOptions LogLevel { get; set; } = new();
}

public sealed class LogLevelOptions
{
    public string Default { get; set; } = "Information";

    public string? MicrosoftOfficeInterop { get; set; } = "Warning";

    public string? ModelContextProtocol { get; set; } = "Information";
}