namespace MuseDock.Desktop.Models;

public sealed class FolderSummary
{
    public required string RelativePath { get; init; }
    public required string DisplayName { get; init; }
    public required int AssetCount { get; init; }

    public string Label => $"{DisplayName} ({AssetCount})";
}
