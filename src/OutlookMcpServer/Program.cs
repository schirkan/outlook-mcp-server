using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;
using OutlookMcpServer;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Configuration;
using OutlookMcpServer.Domain.Serialization;
using OutlookMcpServer.Domain.Services;
using OutlookMcpServer.Interop;

// === Konfiguration laden ===
// Hierarchie (hohe Prioritaet zuletzt): CommandLine > Environment (OUTLOOKMCPSERVER_*) > appsettings.json
var configuration = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: true)
    .AddEnvironmentVariables(prefix: "OUTLOOKMCPSERVER_")
    .AddCommandLine(args)
    .Build();

var options = configuration.GetSection("OutlookMcpServer").Get<OutlookMcpServerOptions>()
    ?? new OutlookMcpServerOptions();

// Transport-Info nach stderr (bei stdio wichtig: stdout bleibt sauber fuer MCP-Protokoll)
Console.Error.WriteLine(
    $"[OutlookMcpServer] startup: Transport={options.Transport}, " +
    $"AllowSend={options.Outlook.AllowSend}, AllowDelete={options.Outlook.AllowDelete}, " +
    $"AllowCreate={options.Outlook.AllowCreate}, AutoStartOutlook={options.Outlook.AutoStartOutlook}, " +
    $"StartupTimeoutSeconds={options.Outlook.StartupTimeoutSeconds}");

// === stdio Transport (Default) ===
// HTTP/SSE-Transport folgt in Karte 5.5 — der SDK 2.0.0-preview.3 hat
// WithHttpServerTransport() noch nicht in der erwarteten Namespace-Stelle,
// wird mit der naechsten SDK-Version oder einem Workaround nachgezogen.
var builder = Host.CreateApplicationBuilder(args);
builder.Configuration.AddConfiguration(configuration);
ConfigureCoreServices(builder.Services, builder.Configuration);
builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    // Source-Generation-JSON-Kontext registrieren, damit die Domain-DTOs
    // (insbesondere polymorphe ActiveItem-Varianten) auch nach
    // PublishTrimmed=true korrekt serialisiert werden. Ohne diesen Eintrag
    // wirft der erste Tool-Aufruf mit Active-Inspector-Rueckgabe eine
    // NotSupportedException ("JsonTypeInfo metadata for type ... was not
    // provided"). Der Kontext wird via TypeInfoResolverChain in den
    // SDK-DefaultOptions verankert.
    .WithToolsFromAssembly(serializerOptions: OutlookJsonOptionsFactory.Create());

// Logging: Console (stderr, wegen stdio-Transport) + optional File-Logger
// (zusaetzlich fuer Post-Mortem-Analyse; konfigurierbar per
// "Logging:FilePath" in appsettings.json oder
// OUTLOOKMCPSERVER_Logging__FilePath als ENV-Variable).
var logFilePath = configuration.GetValue<string?>("Logging:FilePath");
if (!string.IsNullOrWhiteSpace(logFilePath))
{
    try
    {
        var dir = Path.GetDirectoryName(logFilePath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        builder.Logging.AddProvider(new FileLoggerProvider(logFilePath));
        Console.Error.WriteLine($"[OutlookMcpServer] log file: {logFilePath}");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"[OutlookMcpServer] WARN: file logger disabled: {ex.Message}");
    }
}
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

var host = builder.Build();
try
{
    Console.Error.WriteLine("[OutlookMcpServer] ready, waiting for MCP requests on stdio...");
    await host.RunAsync();
}
catch (Exception ex)
{
    // Letzte Verteidigungslinie: bei Host-Crash sauber nach stderr loggen,
    // damit der Fehler im MCP-Client sichtbar wird (sonst nur Exit-Code 1
    // ohne jegliche Diagnose).
    Console.Error.WriteLine($"[OutlookMcpServer] FATAL host crashed: {ex.GetType().Name}: {ex.Message}");
    Console.Error.WriteLine(ex.ToString());
    Environment.ExitCode = 1;
}

static void ConfigureCoreServices(IServiceCollection services, IConfiguration configuration)
{
    services.Configure<OutlookMcpServerOptions>(configuration.GetSection("OutlookMcpServer"));
    services.AddSingleton<IInteropOutlookAdapter, OutlookInteropAdapter>();
    services.AddSingleton<IOutlookService, OutlookService>();
}