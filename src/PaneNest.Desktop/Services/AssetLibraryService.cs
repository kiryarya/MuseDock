using System.IO;
using PaneNest.Desktop.Models;

namespace PaneNest.Desktop.Services;

public sealed class AssetLibraryService
{
    private static readonly HashSet<string> ImageExtensions =
    [
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg", ".ico", ".avif", ".heic", ".heif", ".tif", ".tiff", ".raw", ".cr2", ".nef", ".arw", ".dng"
    ];

    private static readonly HashSet<string> VideoExtensions =
    [
        ".mp4", ".mov", ".avi", ".wmv", ".webm", ".m4v", ".mkv", ".flv", ".mpeg", ".mpg", ".ts", ".mts", ".m2ts", ".3gp", ".ogv"
    ];

    private static readonly HashSet<string> TextExtensions =
    [
        ".txt", ".md", ".markdown", ".json", ".yaml", ".yml", ".xml", ".csv", ".tsv", ".log", ".ini", ".cfg", ".conf",
        ".cs", ".csproj", ".sln", ".vb", ".fs", ".js", ".jsx", ".ts", ".tsx", ".mjs", ".cjs", ".css", ".scss", ".sass", ".less",
        ".html", ".htm", ".xaml", ".sql", ".ps1", ".psm1", ".bat", ".cmd", ".sh", ".py", ".rb", ".php", ".java", ".kt", ".kts",
        ".swift", ".go", ".rs", ".cpp", ".cc", ".c", ".h", ".hpp", ".toml", ".editorconfig", ".gitignore", ".env"
    ];

    private static readonly HashSet<string> PdfExtensions =
    [
        ".pdf"
    ];

    private static readonly EnumerationOptions VisibleAndHiddenEnumeration = new()
    {
        AttributesToSkip = 0,
        IgnoreInaccessible = true,
        RecurseSubdirectories = false,
        ReturnSpecialDirectories = false
    };

    public IReadOnlyList<NavigationItem> GetAvailableDrives()
    {
        return DriveInfo.GetDrives()
            .Where(drive => drive.IsReady)
            .OrderBy(drive => drive.Name)
            .Select(drive => new NavigationItem
            {
                Path = drive.RootDirectory.FullName,
                Label = $"{drive.Name} {drive.VolumeLabel}".Trim()
            })
            .ToList();
    }

    public IReadOnlyList<NavigationItem> GetChildFolders(string directoryPath)
    {
        try
        {
            return Directory.EnumerateDirectories(directoryPath, "*", VisibleAndHiddenEnumeration)
                .Select(path => new DirectoryInfo(path))
                .OrderBy(info => info.Name, StringComparer.OrdinalIgnoreCase)
                .Select(info => new NavigationItem
                {
                    Path = info.FullName,
                    Label = info.Name
                })
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    public async Task<IReadOnlyList<AssetItem>> GetDirectoryEntriesAsync(
        string rootPath,
        string directoryPath,
        IReadOnlyDictionary<string, AssetMetadata> metadataLookup,
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            var results = new List<AssetItem>();

            foreach (var directory in EnumerateDirectoriesSafe(directoryPath))
            {
                cancellationToken.ThrowIfCancellationRequested();
                results.Add(CreateDirectoryAsset(rootPath, directory, metadataLookup));
            }

            foreach (var file in EnumerateFilesSafe(directoryPath))
            {
                cancellationToken.ThrowIfCancellationRequested();
                results.Add(CreateFileAsset(rootPath, file, metadataLookup));
            }

            return (IReadOnlyList<AssetItem>)results
                .OrderByDescending(item => item.IsDirectory)
                .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }, cancellationToken);
    }

    public AssetKind DetectKind(string extension)
    {
        if (ImageExtensions.Contains(extension))
        {
            return AssetKind.Image;
        }

        if (VideoExtensions.Contains(extension))
        {
            return AssetKind.Video;
        }

        if (TextExtensions.Contains(extension))
        {
            return AssetKind.Text;
        }

        if (PdfExtensions.Contains(extension))
        {
            return AssetKind.Pdf;
        }

        return AssetKind.Other;
    }

    private IEnumerable<string> EnumerateDirectoriesSafe(string directoryPath)
    {
        try
        {
            return Directory.EnumerateDirectories(directoryPath, "*", VisibleAndHiddenEnumeration);
        }
        catch
        {
            return [];
        }
    }

    private IEnumerable<string> EnumerateFilesSafe(string directoryPath)
    {
        try
        {
            return Directory.EnumerateFiles(directoryPath, "*", VisibleAndHiddenEnumeration);
        }
        catch
        {
            return [];
        }
    }

    private AssetItem CreateDirectoryAsset(
        string rootPath,
        string directoryPath,
        IReadOnlyDictionary<string, AssetMetadata> metadataLookup)
    {
        var info = new DirectoryInfo(directoryPath);
        metadataLookup.TryGetValue(directoryPath, out var metadata);

        return new AssetItem
        {
            FilePath = directoryPath,
            Name = info.Name,
            Extension = string.Empty,
            RelativeDirectory = GetRelativeDirectory(rootPath, info.Parent?.FullName ?? rootPath),
            Size = 0,
            LastModified = info.LastWriteTime,
            Kind = AssetKind.Folder,
            IsHidden = IsHidden(info.Attributes),
            IconSource = FileIconService.GetFolderIcon(),
            Tags = metadata?.Tags.ToArray() ?? [],
            Note = metadata?.Note ?? string.Empty,
            IsFavorite = metadata?.IsFavorite ?? false
        };
    }

    private AssetItem CreateFileAsset(
        string rootPath,
        string filePath,
        IReadOnlyDictionary<string, AssetMetadata> metadataLookup)
    {
        var info = new FileInfo(filePath);
        var extension = info.Extension.ToLowerInvariant();
        metadataLookup.TryGetValue(filePath, out var metadata);

        return new AssetItem
        {
            FilePath = filePath,
            Name = info.Name,
            Extension = string.IsNullOrEmpty(extension) ? "(none)" : extension,
            RelativeDirectory = GetRelativeDirectory(rootPath, info.DirectoryName ?? rootPath),
            Size = info.Length,
            LastModified = info.LastWriteTime,
            Kind = DetectKind(extension),
            IsHidden = IsHidden(info.Attributes),
            IconSource = FileIconService.GetFileIcon(filePath, extension),
            Tags = metadata?.Tags.ToArray() ?? [],
            Note = metadata?.Note ?? string.Empty,
            IsFavorite = metadata?.IsFavorite ?? false
        };
    }

    private static bool IsHidden(FileAttributes attributes)
    {
        return attributes.HasFlag(FileAttributes.Hidden);
    }

    private static string GetRelativeDirectory(string rootPath, string directoryPath)
    {
        var relative = Path.GetRelativePath(rootPath, directoryPath).Replace("\\", "/");
        return relative == "." ? string.Empty : relative;
    }
}
