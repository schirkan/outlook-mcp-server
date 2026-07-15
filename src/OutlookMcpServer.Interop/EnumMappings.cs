using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Domain.Models.Calendar;
using OutlookMcpServer.Domain.Models.Common;
using OutlookMcpServer.Domain.Models.Mail;

namespace OutlookMcpServer.Interop;

/// <summary>
/// Konvertierung zwischen Outlook-Enum-Werten (OlImportance, OlSensitivity, OlBodyFormat,
/// OlBusyStatus, OlResponseStatus, OlDefaultFolders) und Domain-Enum-Werten.
/// </summary>
internal static class OlEnumMappings
{
    // ----- Importance: OlImportance -----
    // 0 = Low, 1 = Normal, 2 = High
    public static Importance ToImportance(int olImportance) => olImportance switch
    {
        0 => Importance.Low,
        2 => Importance.High,
        _ => Importance.Normal,
    };

    public static int ToOlImportance(Importance importance) => importance switch
    {
        Importance.Low => 0,
        Importance.High => 2,
        _ => 1,
    };

    // ----- Sensitivity: OlSensitivity -----
    // 0 = Normal, 1 = Personal, 2 = Private, 3 = Confidential
    public static Sensitivity ToSensitivity(int olSensitivity) => olSensitivity switch
    {
        1 => Sensitivity.Personal,
        2 => Sensitivity.Private,
        3 => Sensitivity.Confidential,
        _ => Sensitivity.Normal,
    };

    public static int ToOlSensitivity(Sensitivity sensitivity) => sensitivity switch
    {
        Sensitivity.Personal => 1,
        Sensitivity.Private => 2,
        Sensitivity.Confidential => 3,
        _ => 0,
    };

    // ----- BodyType -> OlBodyFormat -----
    // OlBodyFormat: 1=Plain, 2=Html, 3=RichText
    public static int ToOlBodyFormat(ItemBodyType type) => type switch
    {
        ItemBodyType.Html => 2,
        _ => 1,
    };

    public static ItemBodyType ToItemBodyType(int olBodyFormat) => olBodyFormat switch
    {
        2 => ItemBodyType.Html,
        _ => ItemBodyType.Text,
    };

    // ----- BusyStatus -> ShowAs -----
    // OlBusyStatus: 0=Free, 1=Tentative, 2=Busy, 3=OutOfOffice, 4=WorkingElsewhere
    public static ShowAs ToShowAs(int olBusyStatus) => olBusyStatus switch
    {
        0 => ShowAs.Free,
        1 => ShowAs.Tentative,
        3 => ShowAs.OutOfOffice,
        4 => ShowAs.WorkingElsewhere,
        _ => ShowAs.Busy,
    };

    public static int ToOlBusyStatus(ShowAs showAs) => showAs switch
    {
        ShowAs.Free => 0,
        ShowAs.Tentative => 1,
        ShowAs.OutOfOffice => 3,
        ShowAs.WorkingElsewhere => 4,
        _ => 2,
    };

    // ----- Response -> OlResponseStatus -----
    // OlResponseStatus: 0=None, 1=Organizer, 2=Tentative, 3=Accepted, 4=Declined, 5=NotResponded
    public static int ToOlResponseStatus(Response response) => response switch
    {
        Response.None => 0,
        Response.Organizer => 1,
        Response.TentativelyAccepted => 2,
        Response.Accepted => 3,
        Response.Declined => 4,
        _ => 5,
    };

    public static Response ToResponse(int olResponseStatus) => olResponseStatus switch
    {
        1 => Response.Organizer,
        2 => Response.TentativelyAccepted,
        3 => Response.Accepted,
        4 => Response.Declined,
        _ => Response.NotResponded,
    };

    // ----- OlDefaultFolders -----
    // 5=olFolderDeletedItems, 6=olFolderOutbox, 9=olFolderCalendar, 10=olFolderContacts,
    // 11=olFolderDrafts, 12=olFolderInbox, 16=olFolderJunk, 18=olFolderSentMail,
    // 23=olFolderArchive
    public static int? ToOlDefaultFolder(string wellKnownName) => wellKnownName switch
    {
        WellKnownFolder.Inbox => 12,
        WellKnownFolder.Drafts => 11,
        WellKnownFolder.SentItems => 18,
        WellKnownFolder.DeletedItems => 5,
        WellKnownFolder.JunkEmail => 16,
        WellKnownFolder.Archive => 23,
        WellKnownFolder.Outbox => 6,
        _ => null,
    };

    // ----- COMException HResult -> ErrorCode -----
    // Wichtige Outlook-Fehler-Codes:
    // - 0x8004010F (RPC_E_DISCONNECTED): Outlook nicht erreichbar
    // - 0x80004005 (E_FAIL): generischer Fehler
    // - 0x80070005 (E_ACCESSDENIED): Permission verweigert
    // - 0x80040119 (MAPI_E_NETWORK_ERROR): Netzwerk-Problem
    // - 0x80040600 (MAPI_E_NOT_ENOUGH_DISK): Disk voll
    // - 0x80040401 (MAPI_E_NOT_FOUND): Item nicht gefunden
    // - 0x8007000E (E_OUTOFMEMORY): Speicher voll
    public static ErrorCode FromHResult(int hResult)
    {
        // Wichtige Outlook-Fehler-Codes:
        // 0x80040119 (-2147221223) MAPI_E_NETWORK_ERROR / RPC_E_DISCONNECTED -> OutlookBusy
        // 0x80070005 (-2147024891) E_ACCESSDENIED -> PermissionDenied
        if (hResult == unchecked((int)0x80040119)) return ErrorCode.OutlookBusy;
        if (hResult == unchecked((int)0x80070005)) return ErrorCode.PermissionDenied;
        return ErrorCode.InternalError;
    }
}