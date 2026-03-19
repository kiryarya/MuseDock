using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using MuseDock.Desktop.Models;

namespace MuseDock.Desktop.Services;

public sealed class LibraryMetadataStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public async Task<LibraryMetadataDocument> LoadAsync(string libraryPath, CancellationToken cancellationToken)
    {
        var metadataPath = GetMetadataPath(libraryPath);
        if (!File.Exists(metadataPath))
        {
            return new LibraryMetadataDocument();
        }

        await using var stream = File.OpenRead(metadataPath);
        var document = await JsonSerializer.DeserializeAsync<LibraryMetadataDocument>(stream, JsonOptions, cancellationToken)
            .ConfigureAwait(false);
        return document ?? new LibraryMetadataDocument();
    }

    public async Task SaveAsync(
        string libraryPath,
        LibraryMetadataDocument document,
        CancellationToken cancellationToken)
    {
        var metadataPath = GetMetadataPath(libraryPath);
        Directory.CreateDirectory(Path.GetDirectoryName(metadataPath)!);

        await using var stream = File.Create(metadataPath);
        await JsonSerializer.SerializeAsync(stream, document, JsonOptions, cancellationToken)
            .ConfigureAwait(false);
    }

    public void Save(string libraryPath, LibraryMetadataDocument document)
    {
        var metadataPath = GetMetadataPath(libraryPath);
        Directory.CreateDirectory(Path.GetDirectoryName(metadataPath)!);

        using var stream = File.Create(metadataPath);
        JsonSerializer.Serialize(stream, document, JsonOptions);
    }

    private static string GetMetadataPath(string libraryPath)
    {
        var appData = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MuseDock");
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(libraryPath)));
        return Path.Combine(appData, $"{hash}.json");
    }
}
