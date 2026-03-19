namespace MuseDock.Desktop.Models;

public sealed class PluginManifest
{
    public string Id { get; set; } = string.Empty;

    public string Name { get; set; } = string.Empty;

    public string Version { get; set; } = "1.0.0";

    public string Description { get; set; } = string.Empty;

    public string Author { get; set; } = string.Empty;

    public List<PluginCommandManifest> Commands { get; set; } = [];
}

public sealed class PluginCommandManifest
{
    public string Id { get; set; } = string.Empty;

    public string Name { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    public string Runner { get; set; } = "powershell";

    public string Entry { get; set; } = string.Empty;

    public List<string> Arguments { get; set; } = [];

    public bool RequiresSelection { get; set; } = true;

    public bool WaitForExit { get; set; } = true;

    public bool RefreshAfterRun { get; set; } = true;

    public List<string> AssetKinds { get; set; } = [];

    public List<string> Extensions { get; set; } = [];
}

public sealed class PluginCatalog
{
    public List<PluginCommandDefinition> Commands { get; } = [];

    public List<string> Errors { get; } = [];
}

public sealed class PluginCommandDefinition
{
    public required string PluginId { get; init; }

    public required string PluginName { get; init; }

    public string PluginDescription { get; init; } = string.Empty;

    public required string PluginDirectory { get; init; }

    public required string CommandId { get; init; }

    public required string CommandName { get; init; }

    public string CommandDescription { get; init; } = string.Empty;

    public required string Runner { get; init; }

    public required string EntryPath { get; init; }

    public IReadOnlyList<string> Arguments { get; init; } = [];

    public bool RequiresSelection { get; init; }

    public bool WaitForExit { get; init; }

    public bool RefreshAfterRun { get; init; }

    public HashSet<string> AssetKinds { get; init; } = new(StringComparer.OrdinalIgnoreCase);

    public HashSet<string> Extensions { get; init; } = new(StringComparer.OrdinalIgnoreCase);

    public bool Matches(PluginInvocationContext context)
    {
        if (RequiresSelection && !context.HasSelection)
        {
            return false;
        }

        if (!context.HasSelection)
        {
            return true;
        }

        if (AssetKinds.Count > 0 && context.SelectedItems.Any(item => !AssetKinds.Contains(item.AssetKind)))
        {
            return false;
        }

        if (Extensions.Count > 0 && context.SelectedItems.Any(item =>
                !item.IsDirectory &&
                !string.IsNullOrWhiteSpace(item.Extension) &&
                !Extensions.Contains(item.Extension)))
        {
            return false;
        }

        return true;
    }
}

public sealed class PluginInvocationContext
{
    public required string CurrentDirectory { get; init; }

    public required string CurrentDrivePath { get; init; }

    public List<PluginSelectedItem> SelectedItems { get; init; } = [];

    public bool HasSelection => SelectedItems.Count > 0;

    public string? PrimarySelectedPath => SelectedItems.FirstOrDefault()?.FilePath;
}

public sealed class PluginSelectedItem
{
    public required string FilePath { get; init; }

    public required string Name { get; init; }

    public required string Extension { get; init; }

    public required string AssetKind { get; init; }

    public required bool IsDirectory { get; init; }

    public string[] Tags { get; init; } = [];

    public string Note { get; init; } = string.Empty;

    public bool IsFavorite { get; init; }
}

public sealed class PluginExecutionResult
{
    public required string ContextPath { get; init; }

    public int? ExitCode { get; init; }
}
