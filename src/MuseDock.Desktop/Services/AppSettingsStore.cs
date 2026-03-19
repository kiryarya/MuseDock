using System.IO;
using System.Text.Json;
using MuseDock.Desktop.Models;

namespace MuseDock.Desktop.Services;

public sealed class AppSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public async Task<AppSettingsDocument> LoadAsync(CancellationToken cancellationToken)
    {
        var path = GetStoragePath();
        if (!File.Exists(path))
        {
            return new AppSettingsDocument();
        }

        await using var stream = File.OpenRead(path);
        var document = await JsonSerializer.DeserializeAsync<AppSettingsDocument>(stream, JsonOptions, cancellationToken)
            .ConfigureAwait(false);
        return document ?? new AppSettingsDocument();
    }

    public async Task SaveAsync(AppSettingsDocument document, CancellationToken cancellationToken)
    {
        var path = GetStoragePath();
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        await using var stream = File.Create(path);
        await JsonSerializer.SerializeAsync(stream, document, JsonOptions, cancellationToken)
            .ConfigureAwait(false);
    }

    public void Save(AppSettingsDocument document)
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
        return Path.Combine(appData, "settings.json");
    }
}
