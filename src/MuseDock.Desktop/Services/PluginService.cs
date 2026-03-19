using System.Diagnostics;
using System.IO;
using System.Text.Json;
using MuseDock.Desktop.Models;

namespace MuseDock.Desktop.Services;

public sealed class PluginService
{
    private static readonly JsonSerializerOptions ManifestJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private static readonly JsonSerializerOptions ContextJsonOptions = new()
    {
        WriteIndented = true
    };

    public string UserPluginDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MuseDock",
        "plugins");

    public PluginCatalog LoadCatalog()
    {
        var catalog = new PluginCatalog();
        var seenCommands = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var pluginRoot in GetPluginRoots())
        {
            if (!Directory.Exists(pluginRoot))
            {
                continue;
            }

            IEnumerable<string> manifestPaths;
            try
            {
                manifestPaths = Directory.EnumerateFiles(pluginRoot, "plugin.json", SearchOption.AllDirectories);
            }
            catch (Exception exception)
            {
                catalog.Errors.Add($"{pluginRoot}: {exception.Message}");
                continue;
            }

            foreach (var manifestPath in manifestPaths)
            {
                try
                {
                    using var stream = File.OpenRead(manifestPath);
                    var manifest = JsonSerializer.Deserialize<PluginManifest>(stream, ManifestJsonOptions);
                    if (manifest is null)
                    {
                        catalog.Errors.Add($"{manifestPath}: plugin.json を読み込めません。");
                        continue;
                    }

                    ValidateAndAddManifest(manifest, manifestPath, catalog, seenCommands);
                }
                catch (Exception exception)
                {
                    catalog.Errors.Add($"{manifestPath}: {exception.Message}");
                }
            }
        }

        return catalog;
    }

    public async Task<PluginExecutionResult> ExecuteAsync(
        PluginCommandDefinition command,
        PluginInvocationContext context,
        CancellationToken cancellationToken)
    {
        var runtimeDirectory = Path.Combine(Path.GetTempPath(), "MuseDock", "plugin-context");
        Directory.CreateDirectory(runtimeDirectory);

        var contextPath = Path.Combine(
            runtimeDirectory,
            $"{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.json");

        var payload = new PluginRuntimeContext
        {
            AppName = "MuseDock",
            InvokedAt = DateTimeOffset.Now,
            PluginId = command.PluginId,
            PluginName = command.PluginName,
            CommandId = command.CommandId,
            CommandName = command.CommandName,
            PluginDirectory = command.PluginDirectory,
            AppDirectory = AppContext.BaseDirectory,
            CurrentDirectory = context.CurrentDirectory,
            CurrentDrivePath = context.CurrentDrivePath,
            SelectedItems = context.SelectedItems
        };

        await File.WriteAllTextAsync(
            contextPath,
            JsonSerializer.Serialize(payload, ContextJsonOptions),
            cancellationToken).ConfigureAwait(false);

        using var process = new Process
        {
            StartInfo = CreateStartInfo(command, context, contextPath)
        };

        process.Start();

        if (!command.WaitForExit)
        {
            return new PluginExecutionResult
            {
                ContextPath = contextPath
            };
        }

        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"プラグインが終了コード {process.ExitCode} を返しました。");
        }

        return new PluginExecutionResult
        {
            ContextPath = contextPath,
            ExitCode = process.ExitCode
        };
    }

    private IEnumerable<string> GetPluginRoots()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "Plugins");
        yield return UserPluginDirectory;
    }

    private static void ValidateAndAddManifest(
        PluginManifest manifest,
        string manifestPath,
        PluginCatalog catalog,
        HashSet<string> seenCommands)
    {
        var pluginDirectory = Path.GetDirectoryName(manifestPath)!;
        if (string.IsNullOrWhiteSpace(manifest.Id))
        {
            catalog.Errors.Add($"{manifestPath}: id は必須です。");
            return;
        }

        if (string.IsNullOrWhiteSpace(manifest.Name))
        {
            catalog.Errors.Add($"{manifestPath}: name は必須です。");
            return;
        }

        if (manifest.Commands.Count == 0)
        {
            catalog.Errors.Add($"{manifestPath}: commands が 1 件以上必要です。");
            return;
        }

        foreach (var command in manifest.Commands)
        {
            if (string.IsNullOrWhiteSpace(command.Id) ||
                string.IsNullOrWhiteSpace(command.Name) ||
                string.IsNullOrWhiteSpace(command.Entry))
            {
                catalog.Errors.Add($"{manifestPath}: command の id / name / entry は必須です。");
                continue;
            }

            var normalizedRunner = NormalizeRunner(command.Runner);
            if (normalizedRunner is null)
            {
                catalog.Errors.Add($"{manifestPath}: runner '{command.Runner}' は未対応です。");
                continue;
            }

            var resolvedEntryPath = ResolveEntryPath(normalizedRunner, pluginDirectory, command.Entry);
            if (!IsValidEntryPath(normalizedRunner, resolvedEntryPath, command.Entry))
            {
                catalog.Errors.Add($"{manifestPath}: entry '{command.Entry}' が見つかりません。");
                continue;
            }

            var uniqueKey = $"{manifest.Id}:{command.Id}";
            if (!seenCommands.Add(uniqueKey))
            {
                catalog.Errors.Add($"{manifestPath}: command '{uniqueKey}' が重複しています。");
                continue;
            }

            catalog.Commands.Add(new PluginCommandDefinition
            {
                PluginId = manifest.Id,
                PluginName = manifest.Name,
                PluginDescription = manifest.Description,
                PluginDirectory = pluginDirectory,
                CommandId = command.Id,
                CommandName = command.Name,
                CommandDescription = command.Description,
                Runner = normalizedRunner,
                EntryPath = resolvedEntryPath,
                Arguments = command.Arguments.ToArray(),
                RequiresSelection = command.RequiresSelection,
                WaitForExit = command.WaitForExit,
                RefreshAfterRun = command.RefreshAfterRun,
                AssetKinds = new HashSet<string>(
                    command.AssetKinds
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Select(value => value.Trim().ToLowerInvariant()),
                    StringComparer.OrdinalIgnoreCase),
                Extensions = new HashSet<string>(
                    command.Extensions
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Select(NormalizeExtension),
                    StringComparer.OrdinalIgnoreCase)
            });
        }
    }

    private static string? NormalizeRunner(string? runner)
    {
        return runner?.Trim().ToLowerInvariant() switch
        {
            "powershell" => "powershell",
            "pwsh" => "pwsh",
            "exe" => "exe",
            _ => null
        };
    }

    private static string ResolveEntryPath(string runner, string pluginDirectory, string entry)
    {
        if (Path.IsPathRooted(entry))
        {
            return entry;
        }

        var candidate = Path.Combine(pluginDirectory, entry);
        if (runner != "exe" || File.Exists(candidate))
        {
            return candidate;
        }

        return entry;
    }

    private static bool IsValidEntryPath(string runner, string resolvedEntryPath, string entry)
    {
        if (runner is "powershell" or "pwsh")
        {
            return File.Exists(resolvedEntryPath);
        }

        if (Path.IsPathRooted(entry) ||
            entry.Contains(Path.DirectorySeparatorChar) ||
            entry.Contains(Path.AltDirectorySeparatorChar))
        {
            return File.Exists(resolvedEntryPath);
        }

        return true;
    }

    private static string NormalizeExtension(string extension)
    {
        var normalized = extension.Trim();
        return normalized.StartsWith('.') ? normalized.ToLowerInvariant() : $".{normalized.ToLowerInvariant()}";
    }

    private static ProcessStartInfo CreateStartInfo(
        PluginCommandDefinition command,
        PluginInvocationContext context,
        string contextPath)
    {
        var startInfo = new ProcessStartInfo
        {
            UseShellExecute = false,
            WorkingDirectory = command.PluginDirectory
        };

        startInfo.Environment["MUSEDOCK_CONTEXT_PATH"] = contextPath;
        startInfo.Environment["MUSEDOCK_PLUGIN_DIRECTORY"] = command.PluginDirectory;
        startInfo.Environment["MUSEDOCK_APP_DIRECTORY"] = AppContext.BaseDirectory;
        startInfo.Environment["MUSEDOCK_CURRENT_DIRECTORY"] = context.CurrentDirectory;
        startInfo.Environment["MUSEDOCK_CURRENT_DRIVE"] = context.CurrentDrivePath;
        startInfo.Environment["MUSEDOCK_SELECTED_PATH"] = context.PrimarySelectedPath ?? string.Empty;

        switch (command.Runner)
        {
            case "powershell":
                startInfo.FileName = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.System),
                    "WindowsPowerShell",
                    "v1.0",
                    "powershell.exe");
                startInfo.ArgumentList.Add("-NoProfile");
                startInfo.ArgumentList.Add("-ExecutionPolicy");
                startInfo.ArgumentList.Add("Bypass");
                startInfo.ArgumentList.Add("-File");
                startInfo.ArgumentList.Add(command.EntryPath);
                break;
            case "pwsh":
                startInfo.FileName = "pwsh.exe";
                startInfo.ArgumentList.Add("-NoProfile");
                startInfo.ArgumentList.Add("-File");
                startInfo.ArgumentList.Add(command.EntryPath);
                break;
            default:
                startInfo.FileName = command.EntryPath;
                break;
        }

        foreach (var argument in command.Arguments)
        {
            startInfo.ArgumentList.Add(ExpandArgument(argument, command, context, contextPath));
        }

        return startInfo;
    }

    private static string ExpandArgument(
        string argument,
        PluginCommandDefinition command,
        PluginInvocationContext context,
        string contextPath)
    {
        return argument
            .Replace("{contextPath}", contextPath, StringComparison.OrdinalIgnoreCase)
            .Replace("{pluginDirectory}", command.PluginDirectory, StringComparison.OrdinalIgnoreCase)
            .Replace("{appDirectory}", AppContext.BaseDirectory, StringComparison.OrdinalIgnoreCase)
            .Replace("{currentDirectory}", context.CurrentDirectory, StringComparison.OrdinalIgnoreCase)
            .Replace("{currentDrive}", context.CurrentDrivePath, StringComparison.OrdinalIgnoreCase)
            .Replace("{selectedPath}", context.PrimarySelectedPath ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }
}

internal sealed class PluginRuntimeContext
{
    public string AppName { get; init; } = "MuseDock";

    public DateTimeOffset InvokedAt { get; init; }

    public string PluginId { get; init; } = string.Empty;

    public string PluginName { get; init; } = string.Empty;

    public string CommandId { get; init; } = string.Empty;

    public string CommandName { get; init; } = string.Empty;

    public string PluginDirectory { get; init; } = string.Empty;

    public string AppDirectory { get; init; } = string.Empty;

    public string CurrentDirectory { get; init; } = string.Empty;

    public string CurrentDrivePath { get; init; } = string.Empty;

    public List<PluginSelectedItem> SelectedItems { get; init; } = [];
}
