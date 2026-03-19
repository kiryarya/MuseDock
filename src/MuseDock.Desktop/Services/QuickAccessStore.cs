using System.IO;
using System.Text.Json;
using MuseDock.Desktop.Models;

namespace MuseDock.Desktop.Services;

public sealed class QuickAccessStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public async Task<QuickAccessDocument> LoadAsync(CancellationToken cancellationToken)
    {
        var path = GetStoragePath();
        if (!File.Exists(path))
        {
            return new QuickAccessDocument();
        }

        await using var stream = File.OpenRead(path);
        var document = await JsonSerializer.DeserializeAsync<QuickAccessDocument>(stream, JsonOptions, cancellationToken)
            .ConfigureAwait(false);
        return document ?? new QuickAccessDocument();
    }

    public async Task SaveAsync(QuickAccessDocument document, CancellationToken cancellationToken)
    {
        var path = GetStoragePath();
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        await using var stream = File.Create(path);
        await JsonSerializer.SerializeAsync(stream, document, JsonOptions, cancellationToken)
            .ConfigureAwait(false);
    }

    public void Save(QuickAccessDocument document)
    {
        var path = GetStoragePath();
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        using var stream = File.Create(path);
        JsonSerializer.Serialize(stream, document, JsonOptions);
    }

    private static string GetStoragePath()
    {
        var appData = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MuseDock");
        return Path.Combine(appData, "directory-groups.json");
    }
}
