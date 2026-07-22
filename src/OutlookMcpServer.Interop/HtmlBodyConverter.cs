using System.Text.RegularExpressions;
using OutlookMcpServer.Domain.Models.Common;
using ReverseMarkdown;

namespace OutlookMcpServer.Interop;

/// <summary>
/// Konvertiert Outlook-HTML-Mail-Bodies in andere Formate:
/// <list type="bullet">
///   <item><see cref="BodyFormat.Markdown"/>: HTML → GitHub-kompatibles Markdown (default) via ReverseMarkdown.</item>
///   <item><see cref="BodyFormat.Text"/>: HTML → Plain Text (Tags entfernt, Whitespace normalisiert).</item>
///   <item><see cref="BodyFormat.Html"/>: Original-1:1 durchgereicht (Word/Outlook-Styling erhalten).</item>
/// </list>
/// Markdown nutzt <see cref="ReverseMarkdown.Converter"/> (basiert auf HtmlAgilityPack).
/// Text wird per Regex-Strip aus dem Markdown gewonnen — das ist für LLM-Consumer
/// oft handlicher als ein direkter HTML-Strip, weil Block-Struktur erhalten bleibt.
/// </summary>
internal static class HtmlBodyConverter
{
    // ReverseMarkdown-Converter ist thread-safe und stateless → eine Instanz reicht.
    private static readonly Converter MarkdownConverter = new(new Config
    {
        GithubFlavored = true,
        RemoveComments = true,
        SmartHrefHandling = true,
    });
    /// <summary>
    /// Hauptmethode. Liefert <c>(contentType, content)</c> passend zum gewuenschten Format.
    /// Native Quelle ist immer HTML (Outlook speichert Mails als HTML).
    /// </summary>
    public static (ItemBodyType ContentType, string Content) Convert(
        string? htmlBody,
        BodyFormat target)
    {
        if (htmlBody is null) return (ItemBodyType.Text, string.Empty);

        return target switch
        {
            BodyFormat.Html => (ItemBodyType.Html, htmlBody),
            BodyFormat.Markdown => (ItemBodyType.Text, ToMarkdown(htmlBody)),
            BodyFormat.Text => (ItemBodyType.Text, ToText(htmlBody)),
            _ => (ItemBodyType.Text, ToText(htmlBody)),
        };
    }

    private static string ToMarkdown(string html)
    {
        var md = MarkdownConverter.Convert(html);
        return NormalizeWhitespace(md);
    }

    /// <summary>
    /// Plain-Text aus HTML. ReverseMarkdown liefert bereits eine saubere
    /// Markdown-Repräsentation; der finale Strip passiert durch Entfernen der
    /// Markdown-Steuerzeichen. Vorteil ggü. reinem HTML→Text: Block-Struktur
    /// (Listen, Überschriften) bleibt im Plain-Output als Whitespace erkennbar.
    /// </summary>
    private static string ToText(string html)
    {
        var md = MarkdownConverter.Convert(html);
        var s = md;
        // [text](url) → text
        s = Regex.Replace(s, @"\[([^\]]+)\]\([^)]+\)", "$1");
        // Heading-Marker
        s = Regex.Replace(s, @"(?m)^#{1,6}\s*", string.Empty);
        // Bold/Italic/Strike/Code-Marker
        s = Regex.Replace(s, @"\*\*([^*]+)\*\*", "$1");
        s = Regex.Replace(s, @"\*([^*]+)\*", "$1");
        s = Regex.Replace(s, @"~~([^~]+)~~", "$1");
        s = Regex.Replace(s, @"```[^\n]*\n?", string.Empty);
        s = Regex.Replace(s, @"`([^`]+)`", "$1");
        // Blockquote / List-Marker
        s = Regex.Replace(s, @"(?m)^>\s?", string.Empty);
        s = Regex.Replace(s, @"(?m)^\s*[-*]\s+", string.Empty);
        s = Regex.Replace(s, @"(?m)^\s*\d+\.\s+", string.Empty);
        // Table-Pipes
        s = Regex.Replace(s, @"\|", " ");
        return NormalizeWhitespace(s);
    }

    private static string NormalizeWhitespace(string s)
    {
        // 3+ Leerzeichen → 1 (Space-Sequenzen aus Whitespace-only-Lines).
        s = Regex.Replace(s, @"[ \t]{3,}", " ");
        // Mehr als 2 aufeinanderfolgende \n → max 2 (Markdown-Absatz-Trenner).
        s = Regex.Replace(s, @"\n{3,}", "\n\n");
        // \r\n und \r → \n
        s = s.Replace("\r\n", "\n").Replace('\r', '\n');
        return s.Trim();
    }
}
