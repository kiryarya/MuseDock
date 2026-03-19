namespace FilePilot.Desktop.Models;

public sealed class LibraryMetadataDocument
{
    public Dictionary<string, AssetMetadata> Items { get; set; } = [];
}
