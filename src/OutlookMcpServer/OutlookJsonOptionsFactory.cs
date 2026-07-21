using System.Text.Json;
using OutlookMcpServer.Domain.Serialization;

namespace OutlookMcpServer;

/// <summary>
/// Erstellt <see cref="JsonSerializerOptions"/> fuer MCP-Tool-Parameter und
/// -Rueckgaben. Haengt unseren <see cref="OutlookMcpJsonContext"/> an die
/// TypeInfoResolverChain, damit Domain-DTOs (insbesondere polymorphe
/// <c>ActiveItem</c>-Varianten) auch unter <c>PublishTrimmed=true</c>
/// korrekt serialisiert werden.
/// <para>
/// Hintergrund: der MCP-SDK 2.0-preview verwendet fuer seine eigenen Typen
/// <see cref="ModelContextProtocol.McpJsonUtilities.DefaultOptions"/>. Wenn
/// <c>WithToolsFromAssembly(serializerOptions)</c> aufgerufen wird, uebergibt
/// der SDK diese Optionen an die generierten <see cref="Microsoft.Extensions.AI.AIFunction"/>s.
/// Wir erweitern sie um unseren Kontext, anstatt sie zu ersetzen, damit die
/// SDK-Protokolltypen (JsonRpcMessage etc.) weiterhin funktionieren.
/// </para>
/// </summary>
public static class OutlookJsonOptionsFactory
{
    public static JsonSerializerOptions Create()
    {
        // Start mit frischen Web-Defaults + unsere Naming-Policy, dann den
        // SDK-Default-Kontext (fuer JsonRpcMessage etc.) und unseren Domain-
        // Kontext anhaengen.
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase,
        };
        options.TypeInfoResolverChain.Insert(0, OutlookMcpJsonContext.Default);
        return options;
    }
}