using System.Text.Json.Serialization.Metadata;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;
using OutlookMcpServer;
using OutlookMcpServer.Domain.Abstractions;
using OutlookMcpServer.Domain.Configuration;
using OutlookMcpServer.Domain.Services;
using OutlookMcpServer.Interop;

try
{
    using var configuration = ServerConfiguration.Load(AppContext.BaseDirectory, args);
    var settings = configuration.GetSection("OutlookMcpServer");
    var options = settings.Get<OutlookMcpServerOptions>() ?? new();
    var builder = new HostBuilder().UseContentRoot(AppContext.BaseDirectory);

    builder.ConfigureServices(services =>
    {
        services.Configure<OutlookMcpServerOptions>(settings);
        services.AddSingleton<IInteropOutlookAdapter, OutlookInteropAdapter>();
        services.AddSingleton<IOutlookService, OutlookService>();
        // Preserve generated metadata first; allow schema discovery for other tool DTOs
        // (the COM host is published untrimmed).
        var serializerOptions = OutlookJsonOptionsFactory.Create();
        serializerOptions.TypeInfoResolverChain.Add(new DefaultJsonTypeInfoResolver());
        services.AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly(serializerOptions: serializerOptions);
    });
    builder.ConfigureLogging(logging => ConfigureLogging(logging, configuration));

    Console.Error.WriteLine(
        "[OutlookMcpServer] startup: Transport=stdio, " +
        $"AllowSend={options.Outlook.AllowSend}, AllowDelete={options.Outlook.AllowDelete}, " +
        $"AllowCreate={options.Outlook.AllowCreate}, AutoStartOutlook={options.Outlook.AutoStartOutlook}, " +
        $"StartupTimeoutSeconds={options.Outlook.StartupTimeoutSeconds}");
    Console.Error.WriteLine("[OutlookMcpServer] waiting for MCP requests on stdio...");

    using var host = builder.Build();
    await host.RunAsync();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"[OutlookMcpServer] FATAL: {ex}");
    Environment.ExitCode = 1;
}

static void ConfigureLogging(ILoggingBuilder logging, IConfiguration configuration)
{
    logging.AddConfiguration(configuration.GetSection("Logging"));
    logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);
    var logFilePath = configuration.GetValue<string?>("Logging:FilePath");
    if (string.IsNullOrWhiteSpace(logFilePath)) return;

    try
    {
        var dir = Path.GetDirectoryName(logFilePath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        logging.AddProvider(new FileLoggerProvider(logFilePath));
        Console.Error.WriteLine($"[OutlookMcpServer] log file: {logFilePath}");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"[OutlookMcpServer] WARN: file logger disabled: {ex.Message}");
    }
}
