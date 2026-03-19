namespace MuseDock.Desktop.Models;

public sealed class AssetMetadata
{
    public List<string> Tags { get; set; } = [];
    public string Note { get; set; } = string.Empty;
    public bool IsFavorite { get; set; }
}
