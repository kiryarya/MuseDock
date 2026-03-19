namespace MuseDock.Desktop.Models;

public sealed class AppSettingsDocument
{
    public Dictionary<string, string> Shortcuts { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
