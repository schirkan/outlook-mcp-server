namespace OutlookMcpServer.IntegrationTests;

/// <summary>
/// Integration-Tests fuer MailFolder-Operationen gegen ein laufendes
/// klassisches Outlook-Profil.
/// </summary>
public sealed class MailFolderIntegrationTests : OutlookIntegrationTestBase
{
    [SkippableFact]
    public async Task ListMailFolders_ReturnsAtLeastInbox()
    {
        SkipIfOutlookNotAvailable();

        // Act
        var result = await Adapter.ListMailFoldersAsync(cancellationToken: CancellationToken.None);

        // Assert
        Assert.NotNull(result);
        Assert.NotNull(result.Value);
        Assert.NotEmpty(result.Value);
        // Outlook legt immer mindestens Posteingang/Inbox an.
        Assert.Contains(result.Value, f =>
            string.Equals(f.WellKnownName, "inbox", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(f.DisplayName, "Posteingang", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(f.DisplayName, "Inbox", StringComparison.OrdinalIgnoreCase));
    }

    [SkippableFact]
    public async Task ListMailFolders_DefaultScope_DoesNotIncludeHidden()
    {
        SkipIfOutlookNotAvailable();

        // Act: includeHidden=false (Default)
        var result = await Adapter.ListMailFoldersAsync(includeHidden: false, cancellationToken: CancellationToken.None);

        // Assert: keine Folder-Namen mit fuehrendem "$"
        Assert.All(result.Value, f =>
            Assert.False(
                !string.IsNullOrEmpty(f.DisplayName) && f.DisplayName.StartsWith("$", StringComparison.Ordinal),
                $"Hidden-Folder {f.DisplayName} wurde trotz includeHidden=false zurueckgegeben"));
    }

    [SkippableFact]
    public async Task GetMailFolder_Inbox_ReturnsValidFolder()
    {
        SkipIfOutlookNotAvailable();

        // Act
        var folder = await Adapter.GetMailFolderAsync(WellKnownFolder.Inbox, cancellationToken: CancellationToken.None);

        // Assert
        Assert.NotNull(folder);
        Assert.False(string.IsNullOrEmpty(folder.Id));
        Assert.Equal(WellKnownFolder.Inbox, folder.WellKnownName ?? WellKnownFolder.Inbox);
        Assert.NotEqual(0, folder.TotalItemCount);
    }
}