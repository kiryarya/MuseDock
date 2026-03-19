namespace MuseDock.Desktop.Models;

public sealed class LibraryMetadataDocument
{
    public Dictionary<string, AssetMetadata> Items { get; set; } = [];

    public Dictionary<string, double> ColumnWidths { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
