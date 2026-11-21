using Xunit;

namespace OutlookMcpServer.Domain.Tests;

public sealed class ServerConfigurationTests
{
    [Fact]
    public void ConfigurationUsesBinaryDirectoryAndCliOverCanonicalEnvOverLegacyEnvOverJson()
    {
        const string key = "OutlookMcpServer:Outlook:ProfileName";
        const string canonical = "OUTLOOKMCPSERVER__OUTLOOK__PROFILENAME";
        const string legacy = "OUTLOOKMCPSERVER_OutlookMcpServer__Outlook__ProfileName";
        var priorCanonical = Environment.GetEnvironmentVariable(canonical);
        var priorLegacy = Environment.GetEnvironmentVariable(legacy);
        var directory = Directory.CreateTempSubdirectory();
        try
        {
            Environment.SetEnvironmentVariable(canonical, null);
            Environment.SetEnvironmentVariable(legacy, null);
            File.WriteAllText(Path.Combine(directory.FullName, "appsettings.json"),
                """{"OutlookMcpServer":{"Outlook":{"ProfileName":"json"}}}""");
            Assert.Equal("json", ReadProfile());

            Environment.SetEnvironmentVariable(legacy, "legacy");
            Assert.Equal("legacy", ReadProfile());

            Environment.SetEnvironmentVariable(canonical, "canonical");
            Assert.Equal("canonical", ReadProfile());
            Assert.Equal("cli", ReadProfile($"--{key}=cli"));
        }
        finally
        {
            Environment.SetEnvironmentVariable(canonical, priorCanonical);
            Environment.SetEnvironmentVariable(legacy, priorLegacy);
            directory.Delete(recursive: true);
        }

        string? ReadProfile(params string[] args)
        {
            using var configuration = ServerConfiguration.Load(directory.FullName, args);
            return configuration[key];
        }
    }
}
