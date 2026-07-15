using System.Text.RegularExpressions;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Domain.Validation;

/// <summary>
/// Statische Input-Validierung. Wirft <see cref="OutlookServiceException"/>
/// mit Code <see cref="ErrorCode.InvalidInput"/> bzw. <see cref="ErrorCode.AttachmentTooLarge"/>.
/// </summary>
public static partial class ValidationHelpers
{
    // Sehr einfache E-Mail-Validierung (RFC 5322 ist zu komplex, deshalb pragmatisch).
    [GeneratedRegex(@"^[^\s@]+@[^\s@]+\.[^\s@]+$", RegexOptions.CultureInvariant)]
    private static partial Regex EmailRegex();

    public static void ValidateEmail(string email, string paramName = "email")
    {
        if (string.IsNullOrWhiteSpace(email))
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"{paramName}: Wert ist leer oder null");
        }
        if (!EmailRegex().IsMatch(email))
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"{paramName}: ungueltiges E-Mail-Format '{email}'");
        }
    }

    public static void ValidateEmails(IEnumerable<string>? emails, string paramName = "to")
    {
        if (emails is null) return;
        foreach (var email in emails)
        {
            ValidateEmail(email, paramName);
        }
    }

    public static void ValidateDateTimeTimeZone(
        DateTimeTimeZone dateTimeTimeZone,
        string paramName = "start")
    {
        if (dateTimeTimeZone is null)
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"{paramName}: DateTimeTimeZone ist null");
        }
        if (string.IsNullOrWhiteSpace(dateTimeTimeZone.DateTime))
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"{paramName}.dateTime: Wert ist leer");
        }
        if (!string.IsNullOrWhiteSpace(dateTimeTimeZone.TimeZone))
        {
            try
            {
                _ = TimeZoneInfo.FindSystemTimeZoneById(dateTimeTimeZone.TimeZone);
            }
            catch (TimeZoneNotFoundException)
            {
                throw new OutlookServiceException(
                    ErrorCode.InvalidInput,
                    $"{paramName}.timeZone: Zeitzone '{dateTimeTimeZone.TimeZone}' nicht gefunden");
            }
            catch (InvalidTimeZoneException)
            {
                throw new OutlookServiceException(
                    ErrorCode.InvalidInput,
                    $"{paramName}.timeZone: Zeitzone '{dateTimeTimeZone.TimeZone}' hat ungueltige Daten");
            }
        }
    }

    public static void ValidateAttachmentSize(long sizeBytes, long maxBytes)
    {
        if (sizeBytes > maxBytes)
        {
            throw new OutlookServiceException(
                ErrorCode.AttachmentTooLarge,
                $"Attachment-Groesse {sizeBytes:N0} Bytes ueberschreitet Limit {maxBytes:N0} Bytes");
        }
    }

    public static void ValidateStringNotEmpty(string? value, string paramName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"{paramName}: Wert ist leer oder null");
        }
    }

    public static void ValidateRange(int value, int min, int max, string paramName)
    {
        if (value < min || value > max)
        {
            throw new OutlookServiceException(
                ErrorCode.InvalidInput,
                $"{paramName}: Wert {value} ausserhalb des erlaubten Bereichs [{min}, {max}]");
        }
    }

    /// <summary>Konvertiert einen Well-Known-Namen oder EntryID zu EntryID. Stub fuer Karte 2; Karte 3 (Interop-Adapter) implementiert die Aufloesung.</summary>
    public static void ValidateFolderId(string folderId, string paramName = "folderId")
    {
        ValidateStringNotEmpty(folderId, paramName);
    }
}