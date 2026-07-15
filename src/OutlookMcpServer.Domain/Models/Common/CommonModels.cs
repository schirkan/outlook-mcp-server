using System.Text.Json.Serialization;

namespace OutlookMcpServer.Domain.Models.Common;

/// <summary>
/// E-Mail-Adresse + optionaler Anzeigename. 1:1 zu Microsoft Graph
/// <c>EmailAddress</c>-Typ.
/// </summary>
public sealed record EmailAddress
{
    [JsonPropertyName("name")]
    public string? Name { get; init; }

    [JsonPropertyName("address")]
    public string Address { get; init; } = string.Empty;
}

/// <summary>
/// Zeitstempel + zugehoerige Zeitzone (IANA-Name, z. B. "Europe/Berlin").
/// 1:1 zu Microsoft Graph <c>DateTimeTimeZone</c>.
/// </summary>
public sealed record DateTimeTimeZone
{
    [JsonPropertyName("dateTime")]
    public string DateTime { get; init; } = string.Empty;

    [JsonPropertyName("timeZone")]
    public string TimeZone { get; init; } = string.Empty;
}

/// <summary>
/// Body einer Mail oder eines Termins. 1:1 zu Microsoft Graph <c>ItemBody</c>.
/// </summary>
public sealed record ItemBody
{
    [JsonPropertyName("contentType")]
    public ItemBodyType ContentType { get; init; } = ItemBodyType.Text;

    [JsonPropertyName("content")]
    public string Content { get; init; } = string.Empty;
}

public enum ItemBodyType
{
    [JsonPropertyName("text")]
    Text,

    [JsonPropertyName("html")]
    Html,
}

/// <summary>
/// Generisches Pagination-Result. <c>NextSkip</c> ist null, wenn keine weiteren
/// Seiten vorhanden sind (siehe <c>API-DESIGN.md</c> Pagination-Konvention).
/// </summary>
public sealed record PagedResult<T>
{
    [JsonPropertyName("value")]
    public IReadOnlyList<T> Value { get; init; } = Array.Empty<T>();

    [JsonPropertyName("nextSkip")]
    public int? NextSkip { get; init; }
}

/// <summary>
/// Ein einzelner Internet-Header (z. B. "X-Mailer", "DKIM-Signature").
/// </summary>
public sealed record InternetMessageHeader
{
    [JsonPropertyName("name")]
    public string Name { get; init; } = string.Empty;

    [JsonPropertyName("value")]
    public string Value { get; init; } = string.Empty;
}