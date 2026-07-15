using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;

var builder = Host.CreateApplicationBuilder(args);

builder.Logging.AddConsole(consoleLogOptions =>
{
    consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
});

// OutlookMcpServer-Block aus appsettings.json ist im builder.Configuration bereits
// automatisch verfuegbar (AddJsonFile("appsettings.json") wird vom
// Host.CreateApplicationBuilder per Default inkludiert).
// Options-Binding und Validierung folgen in der Konfigurations-Karte.

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly(); // Tool-Klassen ([McpServerToolType]) werden per Reflection geladen

// Domain-/Interop-Services werden in den Folge-Karten registriert:
//   builder.Services.AddSingleton<IOutlookService, OutlookService>();
//   builder.Services.AddSingleton<IInteropOutlookAdapter, InteropOutlookAdapter>();
// (Platzhalter, kein Stub noetig fuer Setup-Karte.)

var host = builder.Build();
await host.RunAsync();