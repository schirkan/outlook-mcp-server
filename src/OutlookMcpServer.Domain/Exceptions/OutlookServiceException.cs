namespace OutlookMcpServer.Domain.Exceptions;

/// <summary>
/// Fehler-Codes fuer <see cref="OutlookServiceException"/>. 1:1 zu
/// <c>specs/API-DESIGN.md</c> Fehler-Schema.
/// </summary>
public enum ErrorCode
{
    FolderNotFound,
    MailNotFound,
    EventNotFound,
    CalendarNotFound,
    AttachmentNotFound,
    InvalidInput,
    OutlookNotRunning,
    OutlookBusy,
    PermissionDenied,
    AttachmentTooLarge,
    SendDisabled,
    DeleteDisabled,
    InternalError
}

/// <summary>
/// Domain-Exception mit klassifiziertem Fehler-Code. MCP-Tools mappen diese
/// in das JSON-Fehler-Schema.
/// </summary>
public sealed class OutlookServiceException : Exception
{
    public ErrorCode Code { get; }

    public OutlookServiceException(ErrorCode code, string message, Exception? inner = null)
        : base(message, inner)
    {
        Code = code;
    }
}