using Microsoft.Extensions.Configuration;

namespace OutlookMcpServer;

internal static class ServerConfiguration
{
    internal static ConfigurationManager Load(string basePath, string[] args)
    {
        var configuration = new ConfigurationManager();
        configuration.SetBasePath(basePath)
            .AddJsonFile("appsettings.json", optional: true)
            // Keep the previously supported single-underscore prefix (including the file logger).
            .AddEnvironmentVariables(prefix: "OUTLOOKMCPSERVER_")
            // Double underscores map OUTLOOKMCPSERVER__OUTLOOK__KEY to OutlookMcpServer:Outlook:Key.
            .AddEnvironmentVariables()
            .AddCommandLine(args);
        return configuration;
    }
}
