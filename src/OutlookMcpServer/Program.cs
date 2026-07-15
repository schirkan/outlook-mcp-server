using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Configuration;
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
    $"OutlookMcpServer starting: Transport={options.Transport}, " +
    $"AllowSend={options.Outlook.AllowSend}, AllowDelete={options.Outlook.AllowDelete}, " +
    $"AllowCreate={options.Outlook.AllowCreate}, AutoStartOutlook={options.Outlook.AutoStartOutlook}");

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
    .WithToolsFromAssembly();
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

var host = builder.Build();
await host.RunAsync();

static void ConfigureCoreServices(IServiceCollection services, IConfiguration configuration)
{
    services.Configure<OutlookMcpServerOptions>(configuration.GetSection("OutlookMcpServer"));
    services.AddSingleton<IInteropOutlookAdapter, OutlookInteropAdapter>();
    services.AddSingleton<IOutlookService, OutlookService>();
}