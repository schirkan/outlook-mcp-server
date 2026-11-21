using System.Runtime.InteropServices;
using OutlookMcpServer.Domain.Exceptions;
using OutlookMcpServer.Interop;
using Xunit;

namespace OutlookMcpServer.Domain.Tests;

public sealed class MailInteropRegressionTests
{
    [Theory]
    [InlineData(false, false, true, false)]
    [InlineData(true, false, true, true)]
    [InlineData(false, true, false, false)]
    [InlineData(true, true, false, true)]
    public void DeleteMail_UsesTrashOfOwningStore(bool permanent, bool inTrash, bool moves, bool deletes)
    {
        var store = new Store();
        var mail = new Mail(inTrash ? store.Trash : new Folder("inbox", store));

        OutlookInteropAdapter.DeleteMailItem(mail, permanent);

        Assert.Equal(3, store.RequestedFolderType);
        Assert.Same(moves ? store.Trash : null, mail.MovedTo);
        Assert.Equal(!moves && deletes, mail.Deleted);
        Assert.Equal(moves && deletes, mail.MovedItem?.Deleted == true);
    }

    [Fact]
    public void DeleteMail_RejectsNonMail()
    {
        var store = new Store();
        var wrongItem = new Mail(store.Trash) { Class = 26 };

        var error = Assert.Throws<OutlookServiceException>(() => OutlookInteropAdapter.DeleteMailItem(wrongItem, true));

        Assert.Equal(ErrorCode.MailNotFound, error.Code);
        Assert.Null(store.RequestedFolderType);
        Assert.False(wrongItem.Deleted);
    }

    [Fact]
    public void InvalidFilter_ReportsFailureInsteadOfReturningOriginalCollection()
    {
        var items = new RejectingCollection();

        var error = Assert.Throws<OutlookServiceException>(() => OutlookInteropAdapter.RestrictMailItems(items, "invalid"));

        Assert.Equal(ErrorCode.InvalidInput, error.Code);
        Assert.IsType<COMException>(error.InnerException);
    }

    public sealed class RejectingCollection
    {
        public object Restrict(string filter) => throw new COMException("Filter failed");
    }

    public sealed class Store
    {
        public Store() => Trash = new Folder("trash", this);
        public Folder Trash { get; }
        public int? RequestedFolderType { get; private set; }
        public Folder GetDefaultFolder(int type)
        {
            RequestedFolderType = type;
            return Trash;
        }
    }

    public sealed record Folder(string EntryID, Store Store);

    public sealed class Mail(Folder parent)
    {
        public int Class { get; init; } = 43;
        public Folder Parent { get; } = parent;
        public Folder? MovedTo { get; private set; }
        public Mail? MovedItem { get; private set; }
        public bool Deleted { get; private set; }
        public Mail Move(Folder target)
        {
            MovedTo = target;
            return MovedItem = new Mail(target);
        }
        public void Delete() => Deleted = true;
    }
}
